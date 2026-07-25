using System;
using System.IO;
using Avalonia;

namespace PptxAvalonia;

sealed class Program
{
    [STAThread]
    public static void Main(string[] args)
    {
        // Log fatal crashes so "can't open" issues are diagnosable.
        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
        {
            try
            {
                var dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PptxAvalonia");
                Directory.CreateDirectory(dir);
                var path = Path.Combine(dir, "crash.log");
                File.AppendAllText(path,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {e.ExceptionObject}{Environment.NewLine}");
            }
            catch
            {
                // ignore logging failures
            }
        };

        try
        {
            BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
        }
        catch (Exception ex)
        {
            try
            {
                var dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PptxAvalonia");
                Directory.CreateDirectory(dir);
                File.WriteAllText(Path.Combine(dir, "startup-error.log"), ex.ToString());
            }
            catch
            {
                // ignore
            }

            throw;
        }
    }

    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();
}
