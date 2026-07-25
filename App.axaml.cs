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
            var window = new MainWindow
            {
                DataContext = vm
            };
            desktop.MainWindow = window;

            // Only open files from CLI args automatically.
            // Session restore is deferred and failure-safe so it cannot block startup.
            var args = desktop.Args ?? [];
            var pptx = args.FirstOrDefault(a =>
                a.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase) && File.Exists(a));

            Dispatcher.UIThread.Post(async () =>
            {
                try
                {
                    if (pptx is not null)
                        await vm.LoadPathAsync(pptx);
                    else
                        await vm.TryRestoreSessionAsync();
                }
                catch (Exception ex)
                {
                    vm.StatusText = $"啟動開啟檔案失敗：{ex.Message}";
                }
            }, DispatcherPriority.Background);
        }

        base.OnFrameworkInitializationCompleted();
    }
}
