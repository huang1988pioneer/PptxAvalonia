using System.Diagnostics;
using System.Text;

namespace PptxAvalonia;

/// <summary>
/// Writes unhandled exceptions next to the executable (crash.log),
/// matching XlsxAvalonia behavior for easier diagnosis.
/// </summary>
public static class CrashLog
{
    private static readonly object Gate = new();
    private static string? _path;

    public static string LogPath =>
        _path ??= Path.Combine(AppContext.BaseDirectory, "crash.log");

    public static void Write(string message, Exception? ex = null)
    {
        try
        {
            var sb = new StringBuilder();
            sb.AppendLine("==================================================");
            sb.AppendLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
            sb.AppendLine(message);
            if (ex is not null)
            {
                sb.AppendLine(ex.GetType().FullName);
                sb.AppendLine(ex.ToString());
            }

            lock (Gate)
            {
                File.AppendAllText(LogPath, sb.ToString());
            }

            Trace.WriteLine(sb.ToString());
        }
        catch
        {
            // Never throw from the logger.
        }
    }

    public static void InstallGlobalHandlers()
    {
        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
        {
            Write("AppDomain.UnhandledException", e.ExceptionObject as Exception);
        };

        TaskScheduler.UnobservedTaskException += (_, e) =>
        {
            Write("TaskScheduler.UnobservedTaskException", e.Exception);
            e.SetObserved();
        };
    }
}
