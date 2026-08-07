using System.Diagnostics;
using System.Text;

namespace PptxAvalonia.Services;

/// <summary>
/// Converts legacy / alternative presentation formats to .pptx for the Open XML loader.
/// Uses LibreOffice (or OpenOffice) headless when available.
/// Supported inputs: .ppt, .pps, .odp (and .pptx passes through).
/// </summary>
public sealed class PresentationFormatConverter
{
    private static readonly string[] NativeExtensions = [".pptx"];
    private static readonly string[] ConvertibleExtensions = [".ppt", ".pps", ".odp"];

    private readonly List<string> _tempFiles = [];
    private string? _cachedSoffice;

    public static bool IsSupported(string path)
    {
        var ext = Path.GetExtension(path);
        return NativeExtensions.Any(e => e.Equals(ext, StringComparison.OrdinalIgnoreCase))
               || ConvertibleExtensions.Any(e => e.Equals(ext, StringComparison.OrdinalIgnoreCase));
    }

    public static bool NeedsConversion(string path)
    {
        var ext = Path.GetExtension(path);
        return ConvertibleExtensions.Any(e => e.Equals(ext, StringComparison.OrdinalIgnoreCase));
    }

    public static string SupportedExtensionsFilterDescription =>
        ".pptx / .ppt / .pps / .odp";

    /// <summary>
    /// Returns a path the Open XML loader can open. May be the original path or a temp .pptx.
    /// Caller should treat the returned path as ephemeral when conversion was needed.
    /// </summary>
    public string EnsurePptx(string path, Action<string>? progress = null)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("簡報檔案不存在。", path);

        if (!NeedsConversion(path))
            return path;

        progress?.Invoke($"正在將 {Path.GetExtension(path).ToLowerInvariant()} 轉換為 .pptx…");

        var soffice = FindSoffice();
        if (soffice is null)
        {
            throw new InvalidOperationException(
                "無法開啟此格式：需要安裝 LibreOffice（或 OpenOffice）以轉換 .ppt / .pps / .odp。" +
                " 請安裝後重試，或改用 .pptx。");
        }

        var workDir = Path.Combine(Path.GetTempPath(), "PptxAvalonia-convert", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workDir);
        try
        {
            // LibreOffice is happier with short ASCII paths; copy into work dir.
            var inputName = SanitizeFileName(Path.GetFileName(path));
            var inputCopy = Path.Combine(workDir, inputName);
            File.Copy(path, inputCopy, overwrite: true);

            var psi = new ProcessStartInfo
            {
                FileName = soffice,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                WorkingDirectory = workDir,
            };
            // --norestore avoids crash-recovery dialogs; --headless for no UI.
            psi.ArgumentList.Add("--headless");
            psi.ArgumentList.Add("--norestore");
            psi.ArgumentList.Add("--nolockcheck");
            psi.ArgumentList.Add("--convert-to");
            psi.ArgumentList.Add("pptx");
            psi.ArgumentList.Add("--outdir");
            psi.ArgumentList.Add(workDir);
            psi.ArgumentList.Add(inputCopy);

            using var proc = Process.Start(psi)
                ?? throw new InvalidOperationException("無法啟動 LibreOffice 轉換程序。");

            var stdout = new StringBuilder();
            var stderr = new StringBuilder();
            proc.OutputDataReceived += (_, e) => { if (e.Data is not null) stdout.AppendLine(e.Data); };
            proc.ErrorDataReceived += (_, e) => { if (e.Data is not null) stderr.AppendLine(e.Data); };
            proc.BeginOutputReadLine();
            proc.BeginErrorReadLine();

            if (!proc.WaitForExit(120_000))
            {
                try { proc.Kill(entireProcessTree: true); } catch { /* ignore */ }
                throw new TimeoutException("格式轉換逾時（LibreOffice 超過 120 秒未完成）。");
            }

            // Base name without original extension + .pptx
            var expected = Path.Combine(workDir, Path.GetFileNameWithoutExtension(inputName) + ".pptx");
            var produced = File.Exists(expected)
                ? expected
                : Directory.GetFiles(workDir, "*.pptx").FirstOrDefault();

            if (produced is null || !File.Exists(produced))
            {
                var detail = string.Join(" ",
                    stdout.ToString().Trim(),
                    stderr.ToString().Trim()).Trim();
                var msg = proc.ExitCode != 0
                    ? $"LibreOffice 轉換失敗（結束代碼 {proc.ExitCode}）。"
                    : "LibreOffice 轉換完成但找不到 .pptx 輸出。";
                if (!string.IsNullOrEmpty(detail))
                    msg += " " + detail;
                throw new InvalidDataException(msg);
            }

            // Move to a stable temp file so we can delete the work directory.
            var dest = Path.Combine(Path.GetTempPath(),
                $"pptx-converted-{Guid.NewGuid():N}.pptx");
            File.Copy(produced, dest, overwrite: true);
            _tempFiles.Add(dest);
            return dest;
        }
        finally
        {
            try
            {
                if (Directory.Exists(workDir))
                    Directory.Delete(workDir, recursive: true);
            }
            catch
            {
                // best-effort cleanup
            }
        }
    }

    public void CleanupTemps()
    {
        foreach (var f in _tempFiles)
        {
            try
            {
                if (File.Exists(f))
                    File.Delete(f);
            }
            catch
            {
                // ignore
            }
        }
        _tempFiles.Clear();
    }

    public string? FindSoffice()
    {
        if (_cachedSoffice is not null && File.Exists(_cachedSoffice))
            return _cachedSoffice;

        // PATH lookup
        foreach (var name in new[] { "soffice", "soffice.exe", "libreoffice", "libreoffice.exe" })
        {
            var fromPath = FindOnPath(name);
            if (fromPath is not null)
            {
                _cachedSoffice = fromPath;
                return fromPath;
            }
        }

        // Common install locations (Windows + typical Linux/macOS)
        var candidates = new List<string>();

        if (OperatingSystem.IsWindows())
        {
            var pf = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            var pf86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            candidates.AddRange(
            [
                Path.Combine(pf, "LibreOffice", "program", "soffice.exe"),
                Path.Combine(pf86, "LibreOffice", "program", "soffice.exe"),
                Path.Combine(pf, "LibreOffice 24", "program", "soffice.exe"),
                Path.Combine(pf, "LibreOffice 25", "program", "soffice.exe"),
                Path.Combine(pf, "OpenOffice 4", "program", "soffice.exe"),
                Path.Combine(pf86, "OpenOffice 4", "program", "soffice.exe"),
                Path.Combine(local, "Programs", "LibreOffice", "program", "soffice.exe"),
            ]);
        }
        else if (OperatingSystem.IsMacOS())
        {
            candidates.AddRange(
            [
                "/Applications/LibreOffice.app/Contents/MacOS/soffice",
                "/Applications/OpenOffice.app/Contents/MacOS/soffice",
            ]);
        }
        else
        {
            candidates.AddRange(
            [
                "/usr/bin/soffice",
                "/usr/bin/libreoffice",
                "/usr/lib/libreoffice/program/soffice",
                "/snap/bin/libreoffice",
            ]);
        }

        foreach (var c in candidates)
        {
            if (File.Exists(c))
            {
                _cachedSoffice = c;
                return c;
            }
        }

        return null;
    }

    private static string? FindOnPath(string fileName)
    {
        var pathEnv = Environment.GetEnvironmentVariable("PATH");
        if (string.IsNullOrEmpty(pathEnv)) return null;
        foreach (var dir in pathEnv.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            try
            {
                var full = Path.Combine(dir.Trim(), fileName);
                if (File.Exists(full))
                    return full;
            }
            catch
            {
                // skip bad PATH entries
            }
        }
        return null;
    }

    private static string SanitizeFileName(string name)
    {
        foreach (var c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        if (string.IsNullOrWhiteSpace(name))
            name = "presentation" + Path.GetExtension(name);
        return name;
    }
}
