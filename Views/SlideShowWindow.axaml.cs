using Avalonia.Controls;
using Avalonia.Input;
using PptxAvalonia.ViewModels;

namespace PptxAvalonia.Views;

public partial class SlideShowWindow : Window
{
    public SlideShowWindow()
    {
        InitializeComponent();
        KeyDown += OnKeyDown;
        PointerPressed += OnPointerPressed;
    }

    private void OnPointerPressed(object? sender, PointerPressedEventArgs e)
    {
        if (DataContext is not MainViewModel vm)
            return;

        var point = e.GetCurrentPoint(this);
        if (point.Properties.IsLeftButtonPressed)
        {
            vm.NextSlideCommand.Execute(null);
            e.Handled = true;
        }
        else if (point.Properties.IsRightButtonPressed)
        {
            vm.PreviousSlideCommand.Execute(null);
            e.Handled = true;
        }
    }

    private void OnKeyDown(object? sender, KeyEventArgs e)
    {
        if (DataContext is not MainViewModel vm)
            return;

        switch (e.Key)
        {
            case Key.Escape:
                vm.EndSlideShowCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.Right:
            case Key.Down:
            case Key.PageDown:
            case Key.Space:
            case Key.Enter:
            case Key.N:
                if (vm.IsSlideShowBlank)
                    vm.ClearSlideShowBlankCommand.Execute(null);
                else
                    vm.NextSlideCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.Left:
            case Key.Up:
            case Key.PageUp:
            case Key.Back:
            case Key.P:
                if (vm.IsSlideShowBlank)
                    vm.ClearSlideShowBlankCommand.Execute(null);
                else
                    vm.PreviousSlideCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.Home:
                vm.FirstSlideCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.End:
                vm.LastSlideCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.B:
                vm.ToggleBlackScreenCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.W:
                vm.ToggleWhiteScreenCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.F5:
                vm.ToggleAutoPlayCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.H:
                vm.ToggleSlideShowHudCommand.Execute(null);
                e.Handled = true;
                break;
        }
    }
}
