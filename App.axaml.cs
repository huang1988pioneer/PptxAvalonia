using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using Avalonia.Threading;
using PptxAvalonia.ViewModels;
using PptxAvalonia.Views;

namespace PptxAvalonia;

public partial class App : Application
{
    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            var vm = new MainViewModel();
            desktop.MainWindow = new MainWindow
            {
                DataContext = vm
            };

            // Support: PptxAvalonia.exe path\to\file.pptx
            var args = desktop.Args ?? [];
            var pptx = args.FirstOrDefault(a =>
                a.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase) && File.Exists(a));
            if (pptx is not null)
            {
                Dispatcher.UIThread.Post(async () => await vm.LoadPathAsync(pptx));
            }
        }

        base.OnFrameworkInitializationCompleted();
    }
}
