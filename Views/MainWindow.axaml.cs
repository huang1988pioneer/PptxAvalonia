using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using PptxAvalonia.ViewModels;

namespace PptxAvalonia.Views;

public partial class MainWindow : Window
{
    private Border? _previewHost;
    private ScrollViewer? _scroller;
    private Viewbox? _viewbox;
    private MainViewModel? _vm;

    public MainWindow()
    {
        InitializeComponent();
        Opened += OnOpened;
        DataContextChanged += OnDataContextChanged;
        AddHandler(DragDrop.DropEvent, OnDrop);
        AddHandler(DragDrop.DragOverEvent, OnDragOver);
        KeyDown += OnKeyDown;
    }

    private void OnDataContextChanged(object? sender, EventArgs e)
    {
        if (_vm is not null)
            _vm.PropertyChanged -= OnViewModelPropertyChanged;
        _vm = DataContext as MainViewModel;
        if (_vm is not null)
            _vm.PropertyChanged += OnViewModelPropertyChanged;
    }

    private void OnOpened(object? sender, EventArgs e)
    {
        _vm = DataContext as MainViewModel;
        _vm?.AttachWindow(this);
        if (_vm is not null)
            _vm.PropertyChanged += OnViewModelPropertyChanged;

        _previewHost = this.FindControl<Border>("PreviewHost");
        _scroller = this.FindControl<ScrollViewer>("PreviewScroller");
        _viewbox = this.FindControl<Viewbox>("SlideViewbox");

        if (_previewHost is not null)
            _previewHost.SizeChanged += (_, _) => ReportViewportAndLayout();

        Dispatcher.UIThread.Post(ReportViewportAndLayout, DispatcherPriority.Loaded);
    }

    private void OnViewModelPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (e.PropertyName is nameof(MainViewModel.IsFitMode)
            or nameof(MainViewModel.SlideDisplayWidth)
            or nameof(MainViewModel.SlideDisplayHeight)
            or nameof(MainViewModel.HasDocument)
            or nameof(MainViewModel.CurrentSlideView)
            or nameof(MainViewModel.Zoom)
            or nameof(MainViewModel.IsNormalView)
            or nameof(MainViewModel.ShowNotesPane))
        {
            Dispatcher.UIThread.Post(ApplyPreviewLayout, DispatcherPriority.Render);
        }
    }

    private void ReportViewportAndLayout()
    {
        if (_previewHost is null || _vm is null) return;
        var size = _previewHost.Bounds.Size;
        if (size.Width > 1 && size.Height > 1)
            _vm.UpdateViewport(size);
        ApplyPreviewLayout();
    }

    private void ApplyPreviewLayout()
    {
        if (_viewbox is null || _scroller is null || _vm is null || !_vm.HasDocument || !_vm.IsNormalView)
            return;

        if (_vm.IsFitMode)
        {
            var viewport = _scroller.Viewport;
            var host = _previewHost?.Bounds.Size ?? default;
            var pad = 40.0;
            var availW = viewport.Width > 1 ? Math.Max(40, viewport.Width) : Math.Max(40, host.Width - pad);
            var availH = viewport.Height > 1 ? Math.Max(40, viewport.Height) : Math.Max(40, host.Height - pad);

            _viewbox.Width = availW;
            _viewbox.Height = availH;
            _viewbox.HorizontalAlignment = HorizontalAlignment.Center;
            _viewbox.VerticalAlignment = VerticalAlignment.Center;
            _viewbox.Stretch = Avalonia.Media.Stretch.Uniform;
            _scroller.HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled;
            _scroller.VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled;

            if (_vm.SlideNativeWidth > 0 && _vm.SlideNativeHeight > 0)
            {
                var scale = Math.Min(availW / _vm.SlideNativeWidth, availH / _vm.SlideNativeHeight);
                _vm.NotifyFitScale(scale);
            }
        }
        else
        {
            _viewbox.Width = _vm.SlideDisplayWidth;
            _viewbox.Height = _vm.SlideDisplayHeight;
            _viewbox.HorizontalAlignment = HorizontalAlignment.Center;
            _viewbox.VerticalAlignment = VerticalAlignment.Center;
            _viewbox.Stretch = Avalonia.Media.Stretch.Fill;
            _scroller.HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto;
            _scroller.VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto;
        }
    }

    private void OnSorterItemPressed(object? sender, PointerPressedEventArgs e)
    {
        if (sender is not Border { DataContext: SlideItemViewModel item } || _vm is null)
            return;
        _vm.SelectedSlide = item;
        _vm.SetNormalViewCommand.Execute(null);
        e.Handled = true;
    }

    private void OnDragOver(object? sender, DragEventArgs e)
    {
        e.DragEffects = HasPptx(e) ? DragDropEffects.Copy : DragDropEffects.None;
    }

    private async void OnDrop(object? sender, DragEventArgs e)
    {
        if (_vm is null) return;
        var path = await TryGetPptxPathAsync(e);
        if (path is not null)
            await _vm.LoadPathAsync(path);
    }

    private void OnKeyDown(object? sender, KeyEventArgs e)
    {
        if (_vm is null) return;

        // Don't steal keys while typing in TextBox
        if (TopLevel.GetTopLevel(this)?.FocusManager?.GetFocusedElement() is TextBox)
            return;

        switch (e.Key)
        {
            case Key.Left or Key.PageUp or Key.Up:
                _vm.PreviousSlideCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.Right or Key.PageDown or Key.Down:
                _vm.NextSlideCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.Space:
                if (_vm.IsAutoPlaying)
                    _vm.StopAutoPlayCommand.Execute(null);
                else
                    _vm.NextSlideCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.Home:
                _vm.FirstSlideCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.End:
                _vm.LastSlideCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.F5 when e.KeyModifiers.HasFlag(KeyModifiers.Shift):
                _vm.StartSlideShowCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.F5:
                _vm.StartSlideShowFromBeginningCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.Escape:
                if (_vm.IsAutoPlaying)
                {
                    _vm.StopAutoPlayCommand.Execute(null);
                    e.Handled = true;
                }
                break;
            case Key.F when e.KeyModifiers.HasFlag(KeyModifiers.Control):
                _vm.ToggleFindBarCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.F3 when e.KeyModifiers.HasFlag(KeyModifiers.Shift):
                _vm.FindPreviousCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.F3:
                _vm.FindNextCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.G when e.KeyModifiers.HasFlag(KeyModifiers.Control):
                _vm.GoToSlideCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.D1 when e.KeyModifiers.HasFlag(KeyModifiers.Control):
                _vm.SetNormalViewCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.D2 when e.KeyModifiers.HasFlag(KeyModifiers.Control):
                _vm.SetSorterViewCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.D3 when e.KeyModifiers.HasFlag(KeyModifiers.Control):
                _vm.SetOutlineViewCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.OemPlus when e.KeyModifiers.HasFlag(KeyModifiers.Control):
            case Key.Add when e.KeyModifiers.HasFlag(KeyModifiers.Control):
                _vm.ZoomInCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.OemMinus when e.KeyModifiers.HasFlag(KeyModifiers.Control):
            case Key.Subtract when e.KeyModifiers.HasFlag(KeyModifiers.Control):
                _vm.ZoomOutCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.D0 when e.KeyModifiers.HasFlag(KeyModifiers.Control):
                _vm.ZoomResetCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.O when e.KeyModifiers.HasFlag(KeyModifiers.Control):
                _vm.OpenFileCommand.Execute(null);
                e.Handled = true;
                break;
        }
    }

    private static bool HasPptx(DragEventArgs e)
    {
        if (!e.Data.Contains(DataFormats.Files)) return false;
        var items = e.Data.GetFiles();
        if (items is null) return false;
        foreach (var item in items)
        {
            if ((item.Name ?? "").EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    private static async Task<string?> TryGetPptxPathAsync(DragEventArgs e)
    {
        if (!e.Data.Contains(DataFormats.Files)) return null;
        var items = e.Data.GetFiles();
        if (items is null) return null;
        foreach (var item in items)
        {
            if (item is not IStorageFile file) continue;
            if (!(file.Name ?? "").EndsWith(".pptx", StringComparison.OrdinalIgnoreCase)) continue;
            var path = file.TryGetLocalPath();
            if (!string.IsNullOrEmpty(path)) return path;
            await using var stream = await file.OpenReadAsync();
            var temp = Path.Combine(Path.GetTempPath(), $"pptx-preview-{Guid.NewGuid():N}.pptx");
            await using var fs = File.Create(temp);
            await stream.CopyToAsync(fs);
            return temp;
        }
        return null;
    }
}
