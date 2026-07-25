using System.Collections.ObjectModel;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using PptxAvalonia.Models;
using PptxAvalonia.Services;

namespace PptxAvalonia.ViewModels;

public partial class MainViewModel : ViewModelBase
{
    private readonly PptxLoader _loader = new();
    private readonly SlideRenderer _renderer = new();
    private PptxPresentation? _presentation;
    private Window? _hostWindow;
    private bool _suppressSelection;
    private Size _viewportSize;
    private double _manualZoom = 1.0;
    private readonly DispatcherTimer _autoPlayTimer;

    private const double ViewportPadding = 48;

    public ObservableCollection<SlideItemViewModel> Slides { get; } = [];

    /// <summary>Selectable auto-play intervals (seconds).</summary>
    public IReadOnlyList<int> IntervalOptions { get; } = [1, 2, 3, 5, 8, 10, 15, 30];

    [ObservableProperty]
    private string _title = "PPTX 預覽 — PptxAvalonia";

    [ObservableProperty]
    private string _statusText = "請開啟 .pptx 檔案以預覽投影片。";

    [ObservableProperty]
    private string _fileName = string.Empty;

    [ObservableProperty]
    private int _currentIndex = -1;

    [ObservableProperty]
    private int _slideCount;

    [ObservableProperty]
    private string _pageLabel = "— / —";

    [ObservableProperty]
    private double _zoom = 1.0;

    [ObservableProperty]
    private string _zoomLabel = "符合視窗";

    [ObservableProperty]
    private bool _isBusy;

    [ObservableProperty]
    private bool _hasDocument;

    [ObservableProperty]
    private bool _isFitMode = true;

    [ObservableProperty]
    private Control? _currentSlideView;

    [ObservableProperty]
    private double _slideNativeWidth = 1280;

    [ObservableProperty]
    private double _slideNativeHeight = 720;

    [ObservableProperty]
    private double _slideDisplayWidth = 1280;

    [ObservableProperty]
    private double _slideDisplayHeight = 720;

    [ObservableProperty]
    private SlideItemViewModel? _selectedSlide;

    [ObservableProperty]
    private bool _isAutoPlaying;

    [ObservableProperty]
    private int _autoPlayIntervalSeconds = 3;

    [ObservableProperty]
    private bool _loopAutoPlay = true;

    [ObservableProperty]
    private string _autoPlayButtonText = "▶ 自動放映";

    [ObservableProperty]
    private string _autoPlayStatus = string.Empty;

    public MainViewModel()
    {
        _autoPlayTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(AutoPlayIntervalSeconds) };
        _autoPlayTimer.Tick += OnAutoPlayTick;
    }

    public void AttachWindow(Window window) => _hostWindow = window;

    public void UpdateViewport(Size size)
    {
        if (size.Width <= 1 || size.Height <= 1)
            return;

        _viewportSize = size;
        if (IsFitMode && HasDocument)
            ApplyFitZoom(refresh: false);
    }

    public void NotifyFitScale(double scale)
    {
        if (!IsFitMode || scale <= 0)
            return;

        scale = Math.Clamp(scale, 0.05, 4.0);
        if (Math.Abs(Zoom - scale) < 0.001 && ZoomLabel.StartsWith("符合", StringComparison.Ordinal))
            return;

        Zoom = scale;
        ZoomLabel = $"符合 · {(int)Math.Round(scale * 100)}%";
        SlideDisplayWidth = SlideNativeWidth * scale;
        SlideDisplayHeight = SlideNativeHeight * scale;
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
        if (seconds != value)
        {
            AutoPlayIntervalSeconds = seconds;
            return;
        }

        _autoPlayTimer.Interval = TimeSpan.FromSeconds(seconds);
        if (IsAutoPlaying)
            UpdateAutoPlayStatus();
    }

    partial void OnIsAutoPlayingChanged(bool value)
    {
        AutoPlayButtonText = value ? "⏸ 停止放映" : "▶ 自動放映";
        ToggleAutoPlayCommand.NotifyCanExecuteChanged();
        StartAutoPlayCommand.NotifyCanExecuteChanged();
        StopAutoPlayCommand.NotifyCanExecuteChanged();
    }

    partial void OnLoopAutoPlayChanged(bool value)
    {
        if (IsAutoPlaying)
            UpdateAutoPlayStatus();
    }

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
        if (string.IsNullOrEmpty(path))
        {
            StatusText = "無法取得檔案路徑。";
            return;
        }

        await LoadPathAsync(path);
    }

    public async Task LoadPathAsync(string path)
    {
        StopAutoPlayInternal();

        try
        {
            IsBusy = true;
            StatusText = "正在載入…";
            CurrentSlideView = null;
            Slides.Clear();
            HasDocument = false;
            CurrentIndex = -1;

            var presentation = await Task.Run(() => _loader.Load(path));

            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                _presentation = presentation;
                FileName = presentation.FileName;
                Title = $"{presentation.FileName} — PPTX 預覽";
                SlideCount = presentation.Slides.Count;

                var native = _renderer.GetSlidePixelSize(presentation);
                SlideNativeWidth = native.Width;
                SlideNativeHeight = native.Height;

                for (var i = 0; i < presentation.Slides.Count; i++)
                {
                    var slide = presentation.Slides[i];
                    var thumb = _renderer.BuildThumbnail(presentation, slide, 180);
                    SidesAdd(slide, thumb, i);
                }

                HasDocument = SlideCount > 0;
                if (HasDocument)
                {
                    IsFitMode = true;
                    GoToIndex(0);
                    ApplyFitZoom(refresh: true);
                    StatusText = $"已載入 {SlideCount} 張投影片。可按 F5 自動放映。";
                }
                else
                {
                    StatusText = "簡報中沒有投影片。";
                    PageLabel = "— / —";
                }

                NotifyAutoPlayCommands();
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
            NotifyAutoPlayCommands();
        }
    }

    private void SidesAdd(PptxSlide slide, Control thumb, int i)
        => Slides.Add(new SlideItemViewModel(slide, thumb, i + 1));

    [RelayCommand(CanExecute = nameof(CanNavigatePrev))]
    private void PreviousSlide()
    {
        if (CurrentIndex > 0)
            GoToIndex(CurrentIndex - 1);
    }

    [RelayCommand(CanExecute = nameof(CanNavigateNext))]
    private void NextSlide()
    {
        if (CurrentIndex < SlideCount - 1)
            GoToIndex(CurrentIndex + 1);
    }

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
        StatusText = "縮放：100%（原始大小，可捲動）。";
    }

    [RelayCommand]
    private void FitToWindow()
    {
        IsFitMode = true;
        ApplyFitZoom(refresh: true);
        StatusText = "縮放：符合視窗（完整顯示整頁）。";
    }

    [RelayCommand(CanExecute = nameof(CanToggleAutoPlay))]
    private void ToggleAutoPlay()
    {
        if (IsAutoPlaying)
            StopAutoPlayInternal(userStopped: true);
        else
            StartAutoPlayInternal(fromBeginning: false);
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

    private void StartAutoPlayInternal(bool fromBeginning)
    {
        if (!HasDocument || SlideCount <= 0)
            return;

        if (fromBeginning || CurrentIndex < 0)
            GoToIndex(0);
        else if (CurrentIndex >= SlideCount - 1 && !LoopAutoPlay)
            GoToIndex(0);

        _autoPlayTimer.Interval = TimeSpan.FromSeconds(Math.Clamp(AutoPlayIntervalSeconds, 1, 300));
        _autoPlayTimer.Start();
        IsAutoPlaying = true;
        UpdateAutoPlayStatus();
        StatusText = AutoPlayStatus;
        NotifyAutoPlayCommands();
    }

    private void StopAutoPlayInternal(bool userStopped = false)
    {
        if (_autoPlayTimer.IsEnabled)
            _autoPlayTimer.Stop();

        if (!IsAutoPlaying && !userStopped)
            return;

        IsAutoPlaying = false;
        AutoPlayStatus = string.Empty;
        if (userStopped)
            StatusText = "已停止自動放映。";
        NotifyAutoPlayCommands();
    }

    private void OnAutoPlayTick(object? sender, EventArgs e)
    {
        if (!HasDocument || SlideCount <= 0)
        {
            StopAutoPlayInternal();
            return;
        }

        if (CurrentIndex < SlideCount - 1)
        {
            GoToIndex(CurrentIndex + 1);
            UpdateAutoPlayStatus();
            StatusText = AutoPlayStatus;
            return;
        }

        // Last slide
        if (LoopAutoPlay)
        {
            GoToIndex(0);
            UpdateAutoPlayStatus();
            StatusText = AutoPlayStatus + "（已循環）";
        }
        else
        {
            StopAutoPlayInternal();
            StatusText = "自動放映結束（已到最後一頁）。";
        }
    }

    private void UpdateAutoPlayStatus()
    {
        var loop = LoopAutoPlay ? "循環" : "播完即停";
        AutoPlayStatus = $"自動放映中 · 每 {AutoPlayIntervalSeconds} 秒 · {loop} · {PageLabel}";
    }

    private void NotifyAutoPlayCommands()
    {
        ToggleAutoPlayCommand.NotifyCanExecuteChanged();
        StartAutoPlayCommand.NotifyCanExecuteChanged();
        StartAutoPlayFromBeginningCommand.NotifyCanExecuteChanged();
        StopAutoPlayCommand.NotifyCanExecuteChanged();
    }

    private void ExitFitMode()
    {
        if (!IsFitMode) return;
        IsFitMode = false;
        _manualZoom = Zoom > 0.01 ? Zoom : 1.0;
    }

    private void ApplyFitZoom(bool refresh)
    {
        if (_presentation is null)
            return;

        var native = _renderer.GetSlidePixelSize(_presentation);
        SlideNativeWidth = native.Width;
        SlideNativeHeight = native.Height;

        double scale;
        if (_viewportSize.Width > 1 && _viewportSize.Height > 1)
        {
            var availW = Math.Max(40, _viewportSize.Width - ViewportPadding);
            var availH = Math.Max(40, _viewportSize.Height - ViewportPadding);
            scale = Math.Min(availW / native.Width, availH / native.Height);
            scale = Math.Clamp(scale, 0.05, 4.0);
        }
        else
        {
            scale = 0.7;
        }

        Zoom = scale;
        ZoomLabel = $"符合 · {(int)Math.Round(scale * 100)}%";
        SlideDisplayWidth = native.Width * scale;
        SlideDisplayHeight = native.Height * scale;

        if (refresh)
            RefreshCurrentSlide();
    }

    private void ApplyManualZoom()
    {
        if (_presentation is null)
            return;

        var native = _renderer.GetSlidePixelSize(_presentation);
        SlideNativeWidth = native.Width;
        SlideNativeHeight = native.Height;

        Zoom = _manualZoom;
        ZoomLabel = $"{(int)Math.Round(_manualZoom * 100)}%";
        SlideDisplayWidth = native.Width * _manualZoom;
        SlideDisplayHeight = native.Height * _manualZoom;
        RefreshCurrentSlide();
    }

    private bool CanNavigatePrev() => HasDocument && !IsBusy && CurrentIndex > 0;
    private bool CanNavigateNext() => HasDocument && !IsBusy && CurrentIndex >= 0 && CurrentIndex < SlideCount - 1;

    private void GoToIndex(int index)
    {
        if (_presentation is null || index < 0 || index >= _presentation.Slides.Count)
            return;

        CurrentIndex = index;
        PageLabel = $"{index + 1} / {SlideCount}";

        for (var i = 0; i < Slides.Count; i++)
            Slides[i].IsSelected = i == index;

        _suppressSelection = true;
        try
        {
            SelectedSlide = Slides[index];
        }
        finally
        {
            _suppressSelection = false;
        }

        RefreshCurrentSlide();
        PreviousSlideCommand.NotifyCanExecuteChanged();
        NextSlideCommand.NotifyCanExecuteChanged();

        if (IsAutoPlaying)
            UpdateAutoPlayStatus();
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
        if (!value)
            StopAutoPlayInternal();

        PreviousSlideCommand.NotifyCanExecuteChanged();
        NextSlideCommand.NotifyCanExecuteChanged();
        NotifyAutoPlayCommands();
    }

    partial void OnIsBusyChanged(bool value)
    {
        PreviousSlideCommand.NotifyCanExecuteChanged();
        NextSlideCommand.NotifyCanExecuteChanged();
        NotifyAutoPlayCommands();
    }

    partial void OnCurrentIndexChanged(int value)
    {
        PreviousSlideCommand.NotifyCanExecuteChanged();
        NextSlideCommand.NotifyCanExecuteChanged();
    }
}
