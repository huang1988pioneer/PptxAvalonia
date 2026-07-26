using System.Collections.ObjectModel;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using PptxAvalonia.Models;
using PptxAvalonia.Services;
using PptxAvalonia.Views;

namespace PptxAvalonia.ViewModels;

public partial class MainViewModel : ViewModelBase
{
    private readonly PptxLoader _loader = new();
    private readonly SlideRenderer _renderer = new();
    private readonly SlideExportService _export = new();
    private readonly RecentFilesService _recent = new();
    private readonly AppSettingsService _settings = new();
    private PptxPresentation? _presentation;
    private Window? _hostWindow;
    private SlideShowWindow? _slideShowWindow;
    private bool _suppressSelection;
    private Size _viewportSize;
    private double _manualZoom = 1.0;
    private readonly DispatcherTimer _autoPlayTimer;
    private readonly DispatcherTimer _hudTimer;
    private readonly DispatcherTimer _autoSaveTimer;
    private bool _autoSaveRunning;

    private const double ViewportPadding = 48;

    public ObservableCollection<SlideItemViewModel> Slides { get; } = [];
    public ObservableCollection<RecentFileItemViewModel> RecentFileItems { get; } = [];
    public ObservableCollection<SlideItemViewModel> FindResults { get; } = [];

    public IReadOnlyList<int> IntervalOptions { get; } = [1, 2, 3, 5, 8, 10, 15, 30];

    [ObservableProperty] private string _title = "PptxAvalonia";
    [ObservableProperty] private string _statusText = "請開啟 .pptx 簡報，或於「介面風格」切換五種外觀。";
    [ObservableProperty] private bool _hasRecentFiles;
    [ObservableProperty] private string _fileName = string.Empty;
    [ObservableProperty] private string _filePath = string.Empty;
    [ObservableProperty] private int _currentIndex = -1;
    [ObservableProperty] private int _slideCount;
    [ObservableProperty] private string _pageLabel = "— / —";
    [ObservableProperty] private double _zoom = 1.0;
    [ObservableProperty] private string _zoomLabel = "符合視窗";
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private bool _hasDocument;
    [ObservableProperty] private bool _isFitMode = true;
    [ObservableProperty] private Control? _currentSlideView;

    /// <summary>
    /// Separate visual tree for fullscreen slideshow.
    /// Avalonia forbids re-parenting the same Control into two windows.
    /// </summary>
    [ObservableProperty] private Control? _slideShowSlideView;
    [ObservableProperty] private double _slideNativeWidth = 1280;
    [ObservableProperty] private double _slideNativeHeight = 720;
    [ObservableProperty] private double _slideDisplayWidth = 1280;
    [ObservableProperty] private double _slideDisplayHeight = 720;
    [ObservableProperty] private SlideItemViewModel? _selectedSlide;
    [ObservableProperty] private bool _isAutoPlaying;
    [ObservableProperty] private int _autoPlayIntervalSeconds = 3;
    [ObservableProperty] private bool _loopAutoPlay = true;
    [ObservableProperty] private string _autoPlayButtonText = "▶ 自動放映";
    [ObservableProperty] private string _autoPlayStatus = string.Empty;

    // View modes & panes
    [ObservableProperty] private AppViewMode _viewMode = AppViewMode.Normal;
    [ObservableProperty] private bool _isNormalView = true;
    [ObservableProperty] private bool _isSorterView;
    [ObservableProperty] private bool _isOutlineView;
    [ObservableProperty] private bool _showNotesPane = true;
    [ObservableProperty] private bool _showThumbnails = true;
    [ObservableProperty] private string _currentNotes = string.Empty;
    [ObservableProperty] private string _currentOutline = string.Empty;
    [ObservableProperty] private string _documentInfo = string.Empty;

    // Find
    [ObservableProperty] private string _findQuery = string.Empty;
    [ObservableProperty] private bool _showFindBar;
    [ObservableProperty] private string _findStatus = string.Empty;
    [ObservableProperty] private int _findMatchCount;

    // Slide show
    [ObservableProperty] private bool _isSlideShowActive;
    [ObservableProperty] private bool _isSlideShowBlank;
    [ObservableProperty] private bool _isBlackScreen;
    [ObservableProperty] private bool _showSlideShowHud = true;
    [ObservableProperty] private IBrush _slideShowBlankBrush = Brushes.Black;

    // Auto-save (every 60 seconds when enabled)
    [ObservableProperty] private bool _autoSaveEnabled;
    [ObservableProperty] private string _autoSaveStatus = "自動儲存：關閉";
    [ObservableProperty] private string _lastAutoSaveText = string.Empty;

    // UI skins (Google / LibreOffice / WPS / FreeOffice / PowerPoint)
    [ObservableProperty] private UiSkin _uiSkin = UiSkin.LibreOffice;
    [ObservableProperty] private bool _isChromeClassic = true;
    [ObservableProperty] private bool _isChromeGoogle;
    [ObservableProperty] private bool _isChromeRibbon;
    [ObservableProperty] private string _skinDisplayName = UiSkinInfo.DisplayName(UiSkin.LibreOffice);
    [ObservableProperty] private string _emptyStateSubtitle = $"{UiSkinInfo.DisplayName(UiSkin.LibreOffice)} 風格簡報檢視器";

    public MainViewModel()
    {
        _autoPlayTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(AutoPlayIntervalSeconds) };
        _autoPlayTimer.Tick += OnAutoPlayTick;
        _hudTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(3) };
        _hudTimer.Tick += (_, _) =>
        {
            if (IsSlideShowActive)
                ShowSlideShowHud = false;
            _hudTimer.Stop();
        };

        // Create timer BEFORE setting AutoSaveEnabled (partial method uses the timer).
        var intervalSec = _settings.AutoSaveIntervalSeconds > 0 ? _settings.AutoSaveIntervalSeconds : 60;
        _autoSaveTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(intervalSec) };
        _autoSaveTimer.Tick += OnAutoSaveTick;

        AutoSaveEnabled = _settings.AutoSaveEnabled;
        UpdateAutoSaveStatusLabel();
        RefreshRecentFiles();

        if (Enum.TryParse<UiSkin>(_settings.UiSkinName, ignoreCase: true, out var savedSkin))
            ApplySkin(savedSkin, persist: false, announce: false);
        else
            ApplySkin(UiSkin.LibreOffice, persist: false, announce: false);
    }

    public void AttachWindow(Window window) => _hostWindow = window;

    /// <summary>Restore last session if auto-save was on and a path still exists.</summary>
    public async Task TryRestoreSessionAsync()
    {
        try
        {
            if (!AutoSaveEnabled) return;
            var session = _settings.LoadSession();
            if (session?.FilePath is not { Length: > 0 } path) return;

            var openPath = File.Exists(path)
                ? path
                : (session.RecoverPath is { } r && File.Exists(r) ? r : null);
            if (openPath is null)
            {
                StatusText = "找不到上次工作階段的檔案。";
                return;
            }

            await LoadPathAsync(openPath);
            if (HasDocument && session.SlideIndex >= 0 && session.SlideIndex < SlideCount)
                GoToIndex(session.SlideIndex);

            if (HasDocument)
                StatusText = $"已還原工作階段：{FileName}（投影片 {CurrentIndex + 1}）";
        }
        catch (Exception ex)
        {
            StatusText = $"還原工作階段失敗：{ex.Message}";
        }
    }

    [RelayCommand]
    private async Task RestoreSessionAsync() => await TryRestoreSessionAsync();

    public void UpdateViewport(Size size)
    {
        if (size.Width <= 1 || size.Height <= 1) return;
        _viewportSize = size;
        if (IsFitMode && HasDocument)
            ApplyFitZoom(refresh: false);
    }

    public void NotifyFitScale(double scale)
    {
        if (!IsFitMode || scale <= 0) return;
        scale = Math.Clamp(scale, 0.05, 4.0);
        if (Math.Abs(Zoom - scale) < 0.001 && ZoomLabel.StartsWith("符合", StringComparison.Ordinal))
            return;
        Zoom = scale;
        ZoomLabel = $"符合 · {(int)Math.Round(scale * 100)}%";
        SlideDisplayWidth = SlideNativeWidth * scale;
        SlideDisplayHeight = SlideNativeHeight * scale;
    }

    private void RefreshRecentFiles()
    {
        RecentFileItems.Clear();
        foreach (var p in _recent.Paths)
            RecentFileItems.Add(new RecentFileItemViewModel(p, OpenRecentCommand));
        HasRecentFiles = RecentFileItems.Count > 0;
    }

    partial void OnSelectedSlideChanged(SlideItemViewModel? value)
    {
        if (_suppressSelection || value is null) return;
        var idx = Slides.IndexOf(value);
        if (idx >= 0 && idx != CurrentIndex)
            GoToIndex(idx);
    }

    partial void OnAutoPlayIntervalSecondsChanged(int value)
    {
        var seconds = Math.Clamp(value, 1, 300);
        if (seconds != value) { AutoPlayIntervalSeconds = seconds; return; }
        _autoPlayTimer.Interval = TimeSpan.FromSeconds(seconds);
        if (IsAutoPlaying) UpdateAutoPlayStatus();
    }

    partial void OnIsAutoPlayingChanged(bool value)
    {
        AutoPlayButtonText = value ? "⏸ 停止放映" : "▶ 自動放映";
        NotifyAutoPlayCommands();
    }

    partial void OnLoopAutoPlayChanged(bool value)
    {
        if (IsAutoPlaying) UpdateAutoPlayStatus();
    }

    partial void OnViewModeChanged(AppViewMode value)
    {
        IsNormalView = value == AppViewMode.Normal;
        IsSorterView = value == AppViewMode.Sorter;
        IsOutlineView = value == AppViewMode.Outline;
        StatusText = value switch
        {
            AppViewMode.Sorter => "投影片瀏覽檢視",
            AppViewMode.Outline => "大綱檢視",
            _ => HasDocument ? $"已開啟 {FileName}" : StatusText
        };
    }

    // ——— File ———

    [RelayCommand]
    private async Task OpenFileAsync()
    {
        if (_hostWindow is null) return;
        var files = await _hostWindow.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "開啟 PowerPoint 簡報",
            AllowMultiple = false,
            FileTypeFilter =
            [
                new FilePickerFileType("PowerPoint 簡報")
                {
                    Patterns = ["*.pptx"],
                    MimeTypes = ["application/vnd.openxmlformats-officedocument.presentationml.presentation"]
                },
                FilePickerFileTypes.All
            ]
        });
        if (files.Count == 0) return;
        var file = files[0];
        var path = file.TryGetLocalPath();
        if (string.IsNullOrEmpty(path))
        {
            // Some providers don't expose a local path — copy stream to temp.
            try
            {
                await using var stream = await file.OpenReadAsync();
                path = Path.Combine(Path.GetTempPath(), $"pptx-open-{Guid.NewGuid():N}.pptx");
                await using var fs = File.Create(path);
                await stream.CopyToAsync(fs);
            }
            catch (Exception ex)
            {
                StatusText = $"無法讀取檔案：{ex.Message}";
                return;
            }
        }

        await LoadPathAsync(path);
    }

    [RelayCommand]
    private async Task OpenRecentAsync(string? path)
    {
        if (string.IsNullOrEmpty(path) || !File.Exists(path))
        {
            StatusText = "最近檔案不存在。";
            RefreshRecentFiles();
            return;
        }
        await LoadPathAsync(path);
    }

    [RelayCommand]
    private void ClearRecent()
    {
        _recent.Clear();
        RefreshRecentFiles();
        StatusText = "已清除最近檔案清單。";
    }

    [RelayCommand(CanExecute = nameof(CanUseDocument))]
    private void CloseDocument()
    {
        StopAutoPlayInternal();
        EndSlideShowInternal();
        _presentation = null;
        Slides.Clear();
        FindResults.Clear();
        HasDocument = false;
        CurrentIndex = -1;
        CurrentSlideView = null;
        FileName = string.Empty;
        FilePath = string.Empty;
        PageLabel = "— / —";
        CurrentNotes = string.Empty;
        CurrentOutline = string.Empty;
        DocumentInfo = string.Empty;
        Title = $"PptxAvalonia — {UiSkinInfo.ShortName(UiSkin)}";
        StatusText = "已關閉簡報。";
        NotifyDocCommands();
    }

    [RelayCommand]
    private void ExitApp()
    {
        if (_hostWindow is not null)
            _hostWindow.Close();
        else if (Application.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
            desktop.Shutdown();
    }

    public async Task LoadPathAsync(string path)
    {
        StopAutoPlayInternal();
        EndSlideShowInternal();
        try
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                StatusText = "路徑為空，無法開啟。";
                return;
            }

            path = Path.GetFullPath(path);
            if (!File.Exists(path))
            {
                StatusText = $"找不到檔案：{path}";
                return;
            }

            if (!path.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
            {
                StatusText = "僅支援 .pptx 檔案。";
                return;
            }

            IsBusy = true;
            StatusText = "正在載入…";
            CurrentSlideView = null;
            Slides.Clear();
            FindResults.Clear();
            HasDocument = false;
            CurrentIndex = -1;

            var presentation = await Task.Run(() => _loader.Load(path));

            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                _presentation = presentation;
                FileName = presentation.FileName;
                FilePath = presentation.FilePath;
                Title = $"{presentation.FileName} — {UiSkinInfo.ShortName(UiSkin)}";
                SlideCount = presentation.Slides.Count;

                var native = _renderer.GetSlidePixelSize(presentation);
                SlideNativeWidth = native.Width;
                SlideNativeHeight = native.Height;
                DocumentInfo = $"{SlideCount} 張投影片 · {native.Width:0}×{native.Height:0} px · 16:9 預覽";

                for (var i = 0; i < presentation.Slides.Count; i++)
                {
                    var slide = presentation.Slides[i];
                    var thumb = _renderer.BuildThumbnail(presentation, slide, 180);
                    Slides.Add(new SlideItemViewModel(slide, thumb, i + 1));
                }

                _recent.Add(path);
                RefreshRecentFiles();

                HasDocument = SlideCount > 0;
                if (HasDocument)
                {
                    IsFitMode = true;
                    ViewMode = AppViewMode.Normal;
                    GoToIndex(0);
                    ApplyFitZoom(refresh: true);
                    StatusText = $"已載入 {SlideCount} 張投影片。F5 放映、Ctrl+F 尋找。";
                    if (AutoSaveEnabled)
                        PerformAutoSave(silent: true);
                }
                else
                {
                    StatusText = "簡報中沒有投影片。";
                    PageLabel = "— / —";
                }
                NotifyDocCommands();
            });
        }
        catch (Exception ex)
        {
            StatusText = $"載入失敗：{ex.Message}";
            HasDocument = false;
        }
        finally
        {
            IsBusy = false;
            NotifyDocCommands();
        }
    }

    // ——— Auto-save ———

    partial void OnAutoSaveEnabledChanged(bool value)
    {
        _settings.AutoSaveEnabled = value;
        _settings.AutoSaveIntervalSeconds = 60;
        _settings.SaveSettings();
        UpdateAutoSaveStatusLabel();

        // Guard: property may be set during construction before timers are fully wired.
        if (_autoSaveTimer is null)
            return;

        if (value)
        {
            _autoSaveTimer.Interval = TimeSpan.FromSeconds(60);
            _autoSaveTimer.Start();
            if (HasDocument)
                PerformAutoSave(silent: false);
            else
                StatusText = "已開啟自動儲存（每 60 秒）。開啟簡報後會自動備份工作階段。";
        }
        else
        {
            _autoSaveTimer.Stop();
            LastAutoSaveText = string.Empty;
            StatusText = "已關閉自動儲存。";
        }
    }

    [RelayCommand]
    private void ToggleAutoSave()
    {
        AutoSaveEnabled = !AutoSaveEnabled;
    }

    [RelayCommand(CanExecute = nameof(CanUseDocument))]
    private void SaveNow()
    {
        PerformAutoSave(silent: false);
    }

    private void OnAutoSaveTick(object? sender, EventArgs e)
    {
        if (!AutoSaveEnabled || !HasDocument || IsBusy)
            return;
        PerformAutoSave(silent: true);
    }

    private void PerformAutoSave(bool silent)
    {
        if (_autoSaveRunning || !HasDocument || string.IsNullOrEmpty(FilePath))
            return;

        _autoSaveRunning = true;
        try
        {
            var recover = _settings.SaveRecoverCopy(FilePath);
            var session = new SessionState
            {
                FilePath = FilePath,
                SlideIndex = CurrentIndex,
                IsFitMode = IsFitMode,
                ManualZoom = _manualZoom,
                ShowNotesPane = ShowNotesPane,
                ShowThumbnails = ShowThumbnails,
                ViewMode = (int)ViewMode,
                RecoverPath = recover,
                SavedAt = DateTimeOffset.Now
            };
            _settings.SaveSession(session);

            var time = DateTime.Now.ToString("HH:mm:ss");
            LastAutoSaveText = $"上次自動儲存 {time}";
            UpdateAutoSaveStatusLabel();

            if (!silent)
                StatusText = $"已自動儲存工作階段與復原複本（{time}）。";
            else
                StatusText = $"自動儲存完成 {time} · {PageLabel}";
        }
        catch (Exception ex)
        {
            if (!silent)
                StatusText = $"自動儲存失敗：{ex.Message}";
        }
        finally
        {
            _autoSaveRunning = false;
        }
    }

    private void UpdateAutoSaveStatusLabel()
    {
        AutoSaveStatus = AutoSaveEnabled
            ? "自動儲存：開啟（每 60 秒）"
            : "自動儲存：關閉";
    }

    // ——— Navigation ———

    [RelayCommand(CanExecute = nameof(CanNavigatePrev))]
    private void PreviousSlide()
    {
        if (CurrentIndex > 0) GoToIndex(CurrentIndex - 1);
        PulseHud();
    }

    [RelayCommand(CanExecute = nameof(CanNavigateNext))]
    private void NextSlide()
    {
        if (CurrentIndex < SlideCount - 1) GoToIndex(CurrentIndex + 1);
        PulseHud();
    }

    [RelayCommand(CanExecute = nameof(CanUseDocument))]
    private void FirstSlide()
    {
        GoToIndex(0);
        PulseHud();
    }

    [RelayCommand(CanExecute = nameof(CanUseDocument))]
    private void LastSlide()
    {
        if (SlideCount > 0) GoToIndex(SlideCount - 1);
        PulseHud();
    }

    [RelayCommand(CanExecute = nameof(CanUseDocument))]
    private async Task GoToSlideAsync()
    {
        if (_hostWindow is null || SlideCount <= 0) return;
        var dlg = new GoToSlideWindow();
        dlg.Configure(CurrentIndex + 1, SlideCount);
        var result = await dlg.ShowDialog<int?>(_hostWindow);
        if (result is int idx && idx >= 0 && idx < SlideCount)
            GoToIndex(idx);
    }

    // ——— Zoom ———

    [RelayCommand]
    private void ZoomIn()
    {
        ExitFitMode();
        _manualZoom = Math.Min(3.0, Math.Round(_manualZoom + 0.1, 2));
        ApplyManualZoom();
    }

    [RelayCommand]
    private void ZoomOut()
    {
        ExitFitMode();
        _manualZoom = Math.Max(0.15, Math.Round(_manualZoom - 0.1, 2));
        ApplyManualZoom();
    }

    [RelayCommand]
    private void ZoomReset()
    {
        ExitFitMode();
        _manualZoom = 1.0;
        ApplyManualZoom();
        StatusText = "縮放：100%。";
    }

    [RelayCommand]
    private void FitToWindow()
    {
        IsFitMode = true;
        ApplyFitZoom(refresh: true);
        StatusText = "縮放：符合視窗。";
    }

    // ——— Views ———

    [RelayCommand]
    private void SetNormalView() => ViewMode = AppViewMode.Normal;

    [RelayCommand]
    private void SetSorterView() => ViewMode = AppViewMode.Sorter;

    [RelayCommand]
    private void SetOutlineView() => ViewMode = AppViewMode.Outline;

    [RelayCommand]
    private void ToggleNotesPane() => ShowNotesPane = !ShowNotesPane;

    [RelayCommand]
    private void ToggleThumbnails() => ShowThumbnails = !ShowThumbnails;

    // ——— Find ———

    [RelayCommand]
    private void ToggleFindBar()
    {
        ShowFindBar = !ShowFindBar;
        if (!ShowFindBar)
        {
            ClearFindHighlights();
            FindStatus = string.Empty;
        }
    }

    [RelayCommand(CanExecute = nameof(CanUseDocument))]
    private void FindNext()
    {
        if (string.IsNullOrWhiteSpace(FindQuery))
        {
            StatusText = "請輸入搜尋關鍵字。";
            return;
        }

        RunFind(forward: true);
    }

    [RelayCommand(CanExecute = nameof(CanUseDocument))]
    private void FindPrevious()
    {
        if (string.IsNullOrWhiteSpace(FindQuery)) return;
        RunFind(forward: false);
    }

    private void RunFind(bool forward)
    {
        var q = FindQuery.Trim();
        FindResults.Clear();
        ClearFindHighlights();

        var matches = new List<int>();
        for (var i = 0; i < Slides.Count; i++)
        {
            var text = Slides[i].OutlineText + "\n" + Slides[i].NotesText + "\n" + Slides[i].Title;
            if (text.Contains(q, StringComparison.OrdinalIgnoreCase))
            {
                matches.Add(i);
                Slides[i].IsFindMatch = true;
                FindResults.Add(Slides[i]);
            }
        }

        FindMatchCount = matches.Count;
        if (matches.Count == 0)
        {
            FindStatus = "找不到符合項目";
            StatusText = $"找不到「{q}」";
            return;
        }

        var start = CurrentIndex;
        int next;
        if (forward)
        {
            next = matches.FirstOrDefault(i => i > start, matches[0]);
        }
        else
        {
            next = matches.Where(i => i < start).DefaultIfEmpty(matches[^1]).Last();
        }

        GoToIndex(next);
        FindStatus = $"{matches.IndexOf(next) + 1} / {matches.Count} 個符合";
        StatusText = $"尋找「{q}」— {FindStatus}";
    }

    private void ClearFindHighlights()
    {
        foreach (var s in Slides)
            s.IsFindMatch = false;
    }

    // ——— Export ———

    [RelayCommand(CanExecute = nameof(CanUseDocument))]
    private async Task ExportCurrentSlideAsync()
    {
        if (_hostWindow is null || _presentation is null || CurrentIndex < 0) return;
        var file = await _hostWindow.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "匯出投影片為 PNG",
            SuggestedFileName = $"slide-{CurrentIndex + 1:D3}.png",
            DefaultExtension = "png",
            FileTypeChoices = [new FilePickerFileType("PNG 影像") { Patterns = ["*.png"] }]
        });
        if (file is null) return;
        var path = file.TryGetLocalPath();
        if (string.IsNullOrEmpty(path)) { StatusText = "無法取得儲存路徑。"; return; }

        try
        {
            IsBusy = true;
            StatusText = "正在匯出…";
            await _export.ExportSlidePngAsync(_presentation, _presentation.Slides[CurrentIndex], path, scale: 1.5);
            StatusText = $"已匯出：{Path.GetFileName(path)}";
        }
        catch (Exception ex)
        {
            StatusText = $"匯出失敗：{ex.Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand(CanExecute = nameof(CanUseDocument))]
    private async Task ExportAllSlidesAsync()
    {
        if (_hostWindow is null || _presentation is null) return;
        var folders = await _hostWindow.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = "選擇匯出資料夾",
            AllowMultiple = false
        });
        if (folders.Count == 0) return;
        var dir = folders[0].TryGetLocalPath();
        if (string.IsNullOrEmpty(dir)) { StatusText = "無法取得資料夾路徑。"; return; }

        try
        {
            IsBusy = true;
            var progress = new Progress<int>(n => StatusText = $"匯出中… {n}/{SlideCount}");
            await _export.ExportAllAsync(_presentation, dir, scale: 1.5, progress);
            StatusText = $"已匯出 {SlideCount} 張 PNG 至 {dir}";
        }
        catch (Exception ex)
        {
            StatusText = $"匯出失敗：{ex.Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    // ——— Slide show (fullscreen) ———

    [RelayCommand(CanExecute = nameof(CanUseDocument))]
    private void StartSlideShow()
    {
        if (_presentation is null || SlideCount <= 0) return;
        OpenSlideShow(fromBeginning: false);
    }

    [RelayCommand(CanExecute = nameof(CanUseDocument))]
    private void StartSlideShowFromBeginning()
    {
        if (_presentation is null || SlideCount <= 0) return;
        GoToIndex(0);
        OpenSlideShow(fromBeginning: true);
    }

    [RelayCommand]
    private void EndSlideShow() => EndSlideShowInternal();

    [RelayCommand]
    private void ToggleBlackScreen()
    {
        if (!IsSlideShowActive) return;
        if (IsSlideShowBlank && IsBlackScreen)
        {
            IsSlideShowBlank = false;
        }
        else
        {
            IsBlackScreen = true;
            SlideShowBlankBrush = Brushes.Black;
            IsSlideShowBlank = true;
        }
        PulseHud();
    }

    [RelayCommand]
    private void ToggleWhiteScreen()
    {
        if (!IsSlideShowActive) return;
        if (IsSlideShowBlank && !IsBlackScreen)
        {
            IsSlideShowBlank = false;
        }
        else
        {
            IsBlackScreen = false;
            SlideShowBlankBrush = Brushes.White;
            IsSlideShowBlank = true;
        }
        PulseHud();
    }

    [RelayCommand]
    private void ClearSlideShowBlank()
    {
        IsSlideShowBlank = false;
        PulseHud();
    }

    [RelayCommand]
    private void ToggleSlideShowHud()
    {
        ShowSlideShowHud = !ShowSlideShowHud;
    }

    private void OpenSlideShow(bool fromBeginning)
    {
        try
        {
            if (_hostWindow is null)
            {
                StatusText = "無法放映：主視窗尚未就緒。";
                return;
            }

            if (_presentation is null || SlideCount <= 0)
            {
                StatusText = "無法放映：沒有投影片。";
                return;
            }

            if (fromBeginning)
                GoToIndex(0);

            EndSlideShowInternal();

            // Build a dedicated visual for the slideshow window (must not share CurrentSlideView).
            RefreshSlideShowSlideView();
            if (SlideShowSlideView is null)
            {
                StatusText = "無法放映：投影片渲染失敗。";
                return;
            }

            IsSlideShowActive = true;
            IsSlideShowBlank = false;
            ShowSlideShowHud = true;
            PulseHud();

            _slideShowWindow = new SlideShowWindow { DataContext = this };
            _slideShowWindow.Closed += (_, _) =>
            {
                IsSlideShowActive = false;
                IsSlideShowBlank = false;
                SlideShowSlideView = null;
                _slideShowWindow = null;
                if (IsAutoPlaying)
                    StopAutoPlayInternal(userStopped: true);
                StatusText = "已結束投影片放映。";
            };
            _slideShowWindow.Show(_hostWindow);
            _slideShowWindow.Activate();
            StatusText = "投影片放映中（Esc 結束）。";
        }
        catch (Exception ex)
        {
            CrashLog.Write("OpenSlideShow failed", ex);
            IsSlideShowActive = false;
            SlideShowSlideView = null;
            StatusText = $"放映失敗：{ex.Message}";
        }
    }

    private void EndSlideShowInternal()
    {
        if (_slideShowWindow is not null)
        {
            var w = _slideShowWindow;
            _slideShowWindow = null;
            try { w.Close(); } catch { /* ignore */ }
        }
        IsSlideShowActive = false;
        IsSlideShowBlank = false;
        SlideShowSlideView = null;
    }

    private void PulseHud()
    {
        if (!IsSlideShowActive) return;
        ShowSlideShowHud = true;
        _hudTimer.Stop();
        _hudTimer.Start();
    }

    // ——— Auto play ———

    [RelayCommand(CanExecute = nameof(CanToggleAutoPlay))]
    private void ToggleAutoPlay()
    {
        if (IsAutoPlaying) StopAutoPlayInternal(userStopped: true);
        else StartAutoPlayInternal(fromBeginning: false);
    }

    [RelayCommand(CanExecute = nameof(CanStartAutoPlay))]
    private void StartAutoPlay() => StartAutoPlayInternal(fromBeginning: false);

    [RelayCommand(CanExecute = nameof(CanStartAutoPlay))]
    private void StartAutoPlayFromBeginning() => StartAutoPlayInternal(fromBeginning: true);

    [RelayCommand(CanExecute = nameof(CanStopAutoPlay))]
    private void StopAutoPlay() => StopAutoPlayInternal(userStopped: true);

    private bool CanToggleAutoPlay() => HasDocument && !IsBusy && SlideCount > 0;
    private bool CanStartAutoPlay() => HasDocument && !IsBusy && SlideCount > 0 && !IsAutoPlaying;
    private bool CanStopAutoPlay() => IsAutoPlaying;
    private bool CanUseDocument() => HasDocument && !IsBusy;
    private bool CanNavigatePrev() => HasDocument && !IsBusy && CurrentIndex > 0;
    private bool CanNavigateNext() => HasDocument && !IsBusy && CurrentIndex >= 0 && CurrentIndex < SlideCount - 1;

    private void StartAutoPlayInternal(bool fromBeginning)
    {
        if (!HasDocument || SlideCount <= 0) return;
        if (fromBeginning || CurrentIndex < 0) GoToIndex(0);
        else if (CurrentIndex >= SlideCount - 1 && !LoopAutoPlay) GoToIndex(0);

        _autoPlayTimer.Interval = TimeSpan.FromSeconds(Math.Clamp(AutoPlayIntervalSeconds, 1, 300));
        _autoPlayTimer.Start();
        IsAutoPlaying = true;
        UpdateAutoPlayStatus();
        StatusText = AutoPlayStatus;
        NotifyAutoPlayCommands();
        PulseHud();
    }

    private void StopAutoPlayInternal(bool userStopped = false)
    {
        if (_autoPlayTimer.IsEnabled) _autoPlayTimer.Stop();
        if (!IsAutoPlaying && !userStopped) return;
        IsAutoPlaying = false;
        AutoPlayStatus = string.Empty;
        if (userStopped) StatusText = "已停止自動放映。";
        NotifyAutoPlayCommands();
    }

    private void OnAutoPlayTick(object? sender, EventArgs e)
    {
        if (!HasDocument || SlideCount <= 0) { StopAutoPlayInternal(); return; }

        if (CurrentIndex < SlideCount - 1)
        {
            GoToIndex(CurrentIndex + 1);
            UpdateAutoPlayStatus();
            StatusText = AutoPlayStatus;
            PulseHud();
            return;
        }

        if (LoopAutoPlay)
        {
            GoToIndex(0);
            UpdateAutoPlayStatus();
            StatusText = AutoPlayStatus + "（已循環）";
            PulseHud();
        }
        else
        {
            StopAutoPlayInternal();
            StatusText = "自動放映結束。";
        }
    }

    private void UpdateAutoPlayStatus()
    {
        var loop = LoopAutoPlay ? "循環" : "播完即停";
        AutoPlayStatus = $"自動放映 · 每 {AutoPlayIntervalSeconds} 秒 · {loop} · {PageLabel}";
    }

    private void NotifyAutoPlayCommands()
    {
        ToggleAutoPlayCommand.NotifyCanExecuteChanged();
        StartAutoPlayCommand.NotifyCanExecuteChanged();
        StartAutoPlayFromBeginningCommand.NotifyCanExecuteChanged();
        StopAutoPlayCommand.NotifyCanExecuteChanged();
    }

    private void NotifyDocCommands()
    {
        NotifyAutoPlayCommands();
        CloseDocumentCommand.NotifyCanExecuteChanged();
        FirstSlideCommand.NotifyCanExecuteChanged();
        LastSlideCommand.NotifyCanExecuteChanged();
        GoToSlideCommand.NotifyCanExecuteChanged();
        FindNextCommand.NotifyCanExecuteChanged();
        FindPreviousCommand.NotifyCanExecuteChanged();
        ExportCurrentSlideCommand.NotifyCanExecuteChanged();
        ExportAllSlidesCommand.NotifyCanExecuteChanged();
        StartSlideShowCommand.NotifyCanExecuteChanged();
        StartSlideShowFromBeginningCommand.NotifyCanExecuteChanged();
        PreviousSlideCommand.NotifyCanExecuteChanged();
        NextSlideCommand.NotifyCanExecuteChanged();
        SaveNowCommand.NotifyCanExecuteChanged();
    }

    private void ExitFitMode()
    {
        if (!IsFitMode) return;
        IsFitMode = false;
        _manualZoom = Zoom > 0.01 ? Zoom : 1.0;
    }

    private void ApplyFitZoom(bool refresh)
    {
        if (_presentation is null) return;
        var native = _renderer.GetSlidePixelSize(_presentation);
        SlideNativeWidth = native.Width;
        SlideNativeHeight = native.Height;

        double scale;
        if (_viewportSize.Width > 1 && _viewportSize.Height > 1)
        {
            var availW = Math.Max(40, _viewportSize.Width - ViewportPadding);
            var availH = Math.Max(40, _viewportSize.Height - ViewportPadding);
            // Reserve space for notes pane roughly
            if (ShowNotesPane && IsNormalView) availH = Math.Max(40, availH - 120);
            scale = Math.Min(availW / native.Width, availH / native.Height);
            scale = Math.Clamp(scale, 0.05, 4.0);
        }
        else scale = 0.7;

        Zoom = scale;
        ZoomLabel = $"符合 · {(int)Math.Round(scale * 100)}%";
        SlideDisplayWidth = native.Width * scale;
        SlideDisplayHeight = native.Height * scale;
        if (refresh) RefreshCurrentSlide();
    }

    private void ApplyManualZoom()
    {
        if (_presentation is null) return;
        var native = _renderer.GetSlidePixelSize(_presentation);
        SlideNativeWidth = native.Width;
        SlideNativeHeight = native.Height;
        Zoom = _manualZoom;
        ZoomLabel = $"{(int)Math.Round(_manualZoom * 100)}%";
        SlideDisplayWidth = native.Width * _manualZoom;
        SlideDisplayHeight = native.Height * _manualZoom;
        RefreshCurrentSlide();
    }

    private void GoToIndex(int index)
    {
        if (_presentation is null || index < 0 || index >= _presentation.Slides.Count)
            return;

        CurrentIndex = index;
        PageLabel = $"{index + 1} / {SlideCount}";

        for (var i = 0; i < Slides.Count; i++)
            Slides[i].IsSelected = i == index;

        _suppressSelection = true;
        try { SelectedSlide = Slides[index]; }
        finally { _suppressSelection = false; }

        var slide = _presentation.Slides[index];
        CurrentNotes = string.IsNullOrWhiteSpace(slide.NotesText)
            ? "（此投影片沒有備忘稿）"
            : slide.NotesText;
        CurrentOutline = string.IsNullOrWhiteSpace(slide.OutlineText)
            ? "（無文字內容）"
            : slide.OutlineText;

        RefreshCurrentSlide();
        PreviousSlideCommand.NotifyCanExecuteChanged();
        NextSlideCommand.NotifyCanExecuteChanged();
        if (IsAutoPlaying) UpdateAutoPlayStatus();
    }

    private void RefreshCurrentSlide()
    {
        if (_presentation is null || CurrentIndex < 0 || CurrentIndex >= _presentation.Slides.Count)
        {
            CurrentSlideView = null;
            if (IsSlideShowActive)
                SlideShowSlideView = null;
            return;
        }

        var slide = _presentation.Slides[CurrentIndex];
        CurrentSlideView = _renderer.BuildSlide(_presentation, slide, scale: 1.0);

        // Keep slideshow visual in sync without reusing the editor tree.
        if (IsSlideShowActive)
            RefreshSlideShowSlideView();

        var native = _renderer.GetSlidePixelSize(_presentation);
        SlideNativeWidth = native.Width;
        SlideNativeHeight = native.Height;
        SlideDisplayWidth = native.Width * Zoom;
        SlideDisplayHeight = native.Height * Zoom;
    }

    private void RefreshSlideShowSlideView()
    {
        if (_presentation is null || CurrentIndex < 0 || CurrentIndex >= _presentation.Slides.Count)
        {
            SlideShowSlideView = null;
            return;
        }

        var slide = _presentation.Slides[CurrentIndex];
        SlideShowSlideView = _renderer.BuildSlide(_presentation, slide, scale: 1.0);
    }

    partial void OnHasDocumentChanged(bool value)
    {
        if (!value) StopAutoPlayInternal();
        NotifyDocCommands();
        SaveNowCommand.NotifyCanExecuteChanged();
    }

    partial void OnIsBusyChanged(bool value) => NotifyDocCommands();

    partial void OnCurrentIndexChanged(int value)
    {
        PreviousSlideCommand.NotifyCanExecuteChanged();
        NextSlideCommand.NotifyCanExecuteChanged();
    }

    [RelayCommand]
    private void ShowAbout()
    {
        StatusText =
            $"PptxAvalonia v1.0.0 — 目前介面：{UiSkinInfo.DisplayName(UiSkin)} · 五種風格可切換 · Avalonia + Open XML";
    }

    // ——— UI skins ———

    [RelayCommand]
    private void SetSkinLibreOffice() => ApplySkin(UiSkin.LibreOffice);

    [RelayCommand]
    private void SetSkinGoogleSlides() => ApplySkin(UiSkin.GoogleSlides);

    [RelayCommand]
    private void SetSkinWps() => ApplySkin(UiSkin.Wps);

    [RelayCommand]
    private void SetSkinFreeOffice() => ApplySkin(UiSkin.FreeOffice);

    [RelayCommand]
    private void SetSkinPowerPoint() => ApplySkin(UiSkin.PowerPoint);

    private void ApplySkin(UiSkin skin, bool persist = true, bool announce = true)
    {
        UiSkin = skin;
        var kind = UiSkinInfo.ChromeKind(skin);
        IsChromeClassic = kind == UiChromeKind.Classic;
        IsChromeGoogle = kind == UiChromeKind.Google;
        IsChromeRibbon = kind == UiChromeKind.Ribbon;
        SkinDisplayName = UiSkinInfo.DisplayName(skin);
        EmptyStateSubtitle = $"{SkinDisplayName} 風格簡報檢視器";

        UiThemeService.Apply(skin);

        if (!HasDocument)
            Title = $"PptxAvalonia — {UiSkinInfo.ShortName(skin)}";

        if (persist)
        {
            _settings.UiSkinName = skin.ToString();
            _settings.SaveSettings();
        }

        if (announce)
            StatusText = $"已切換介面風格：{SkinDisplayName}";
    }
}
