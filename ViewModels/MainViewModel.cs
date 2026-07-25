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
    private PptxPresentation? _presentation;
    private Window? _hostWindow;
    private SlideShowWindow? _slideShowWindow;
    private bool _suppressSelection;
    private Size _viewportSize;
    private double _manualZoom = 1.0;
    private readonly DispatcherTimer _autoPlayTimer;
    private readonly DispatcherTimer _hudTimer;

    private const double ViewportPadding = 48;

    public ObservableCollection<SlideItemViewModel> Slides { get; } = [];
    public ObservableCollection<string> RecentFiles { get; } = [];
    public ObservableCollection<SlideItemViewModel> FindResults { get; } = [];

    public IReadOnlyList<int> IntervalOptions { get; } = [1, 2, 3, 5, 8, 10, 15, 30];

    [ObservableProperty] private string _title = "PptxAvalonia — PowerPoint 預覽";
    [ObservableProperty] private string _statusText = "請開啟 .pptx 檔案。";
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
        RefreshRecentFiles();
    }

    public void AttachWindow(Window window) => _hostWindow = window;

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
        RecentFiles.Clear();
        foreach (var p in _recent.Paths)
            RecentFiles.Add(p);
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
        var path = files[0].TryGetLocalPath();
        if (string.IsNullOrEmpty(path)) { StatusText = "無法取得檔案路徑。"; return; }
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
        Title = "PptxAvalonia — PowerPoint 預覽";
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
                Title = $"{presentation.FileName} — PptxAvalonia";
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
        if (_hostWindow is null) return;
        if (fromBeginning) GoToIndex(0);

        EndSlideShowInternal();

        IsSlideShowActive = true;
        IsSlideShowBlank = false;
        ShowSlideShowHud = true;
        PulseHud();

        _slideShowWindow = new SlideShowWindow { DataContext = this };
        _slideShowWindow.Closed += (_, _) =>
        {
            IsSlideShowActive = false;
            IsSlideShowBlank = false;
            _slideShowWindow = null;
            if (IsAutoPlaying)
                StopAutoPlayInternal(userStopped: true);
            StatusText = "已結束投影片放映。";
        };
        _slideShowWindow.Show(_hostWindow);
        _slideShowWindow.Activate();
        StatusText = "投影片放映中（Esc 結束）。";
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
            return;
        }

        var slide = _presentation.Slides[CurrentIndex];
        CurrentSlideView = _renderer.BuildSlide(_presentation, slide, scale: 1.0);
        var native = _renderer.GetSlidePixelSize(_presentation);
        SlideNativeWidth = native.Width;
        SlideNativeHeight = native.Height;
        SlideDisplayWidth = native.Width * Zoom;
        SlideDisplayHeight = native.Height * Zoom;
    }

    partial void OnHasDocumentChanged(bool value)
    {
        if (!value) StopAutoPlayInternal();
        NotifyDocCommands();
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
        StatusText = "PptxAvalonia v1.0.0 — Avalonia + Open XML PowerPoint 預覽器";
    }
}
