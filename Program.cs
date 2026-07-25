using System;
using Avalonia;

namespace PptxAvalonia;

sealed class Program
{
    [STAThread]
    public static void Main(string[] args)
    {
        // Same pattern as XlsxAvalonia: log beside the exe for easy diagnosis.
        CrashLog.InstallGlobalHandlers();

        try
        {
            BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
        }
        catch (Exception ex)
        {
            CrashLog.Write("Main fatal", ex);
            throw;
        }
    }

    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();
}
