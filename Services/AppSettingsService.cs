using System.Text.Json;

namespace PptxAvalonia.Services;

/// <summary>User preferences and auto-save session under LocalAppData.</summary>
public sealed class AppSettingsService
{
    private readonly string _dir;
    private readonly string _settingsPath;
    private readonly string _sessionPath;
    private readonly string _recoverDir;

    public AppSettingsService()
    {
        _dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "PptxAvalonia");
        Directory.CreateDirectory(_dir);
        _settingsPath = Path.Combine(_dir, "settings.json");
        _sessionPath = Path.Combine(_dir, "session.json");
        _recoverDir = Path.Combine(_dir, "AutoRecover");
        Directory.CreateDirectory(_recoverDir);
        LoadSettings();
    }

    public bool AutoSaveEnabled { get; set; }
    public int AutoSaveIntervalSeconds { get; set; } = 60;

    public string RecoverDirectory => _recoverDir;

    public void SaveSettings()
    {
        try
        {
            var dto = new SettingsDto
            {
                AutoSaveEnabled = AutoSaveEnabled,
                AutoSaveIntervalSeconds = AutoSaveIntervalSeconds
            };
            File.WriteAllText(_settingsPath, JsonSerializer.Serialize(dto, JsonOptions));
        }
        catch
        {
            // ignore
        }
    }

    public void SaveSession(SessionState state)
    {
        try
        {
            File.WriteAllText(_sessionPath, JsonSerializer.Serialize(state, JsonOptions));
        }
        catch
        {
            // ignore
        }
    }

    public SessionState? LoadSession()
    {
        try
        {
            if (!File.Exists(_sessionPath)) return null;
            return JsonSerializer.Deserialize<SessionState>(File.ReadAllText(_sessionPath), JsonOptions);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>Copy current presentation to AutoRecover folder; returns recovery path or null.</summary>
    public string? SaveRecoverCopy(string sourcePath)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
                return null;

            var name = Path.GetFileNameWithoutExtension(sourcePath);
            var stamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
            var dest = Path.Combine(_recoverDir, $"{name}.autorecover.pptx");
            // Keep one rolling recover file per source name + latest stamp copy optional
            File.Copy(sourcePath, dest, overwrite: true);

            // Also keep a timestamped backup (trim old ones)
            var stamped = Path.Combine(_recoverDir, $"{name}.{stamp}.pptx");
            File.Copy(sourcePath, stamped, overwrite: true);
            TrimOldRecoveries(name, keep: 3);

            return dest;
        }
        catch
        {
            return null;
        }
    }

    private void TrimOldRecoveries(string baseName, int keep)
    {
        try
        {
            var files = Directory.GetFiles(_recoverDir, $"{baseName}.*.pptx")
                .Where(f => !f.EndsWith(".autorecover.pptx", StringComparison.OrdinalIgnoreCase))
                .Select(f => new FileInfo(f))
                .OrderByDescending(f => f.LastWriteTimeUtc)
                .Skip(keep)
                .ToList();
            foreach (var f in files)
                f.Delete();
        }
        catch
        {
            // ignore
        }
    }

    private void LoadSettings()
    {
        try
        {
            if (!File.Exists(_settingsPath)) return;
            var dto = JsonSerializer.Deserialize<SettingsDto>(File.ReadAllText(_settingsPath), JsonOptions);
            if (dto is null) return;
            AutoSaveEnabled = dto.AutoSaveEnabled;
            AutoSaveIntervalSeconds = dto.AutoSaveIntervalSeconds is > 0 and <= 3600
                ? dto.AutoSaveIntervalSeconds
                : 60;
        }
        catch
        {
            // defaults
        }
    }

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private sealed class SettingsDto
    {
        public bool AutoSaveEnabled { get; set; }
        public int AutoSaveIntervalSeconds { get; set; } = 60;
    }
}

public sealed class SessionState
{
    public string? FilePath { get; set; }
    public int SlideIndex { get; set; }
    public bool IsFitMode { get; set; } = true;
    public double ManualZoom { get; set; } = 1.0;
    public bool ShowNotesPane { get; set; } = true;
    public bool ShowThumbnails { get; set; } = true;
    public int ViewMode { get; set; }
    public string? RecoverPath { get; set; }
    public DateTimeOffset SavedAt { get; set; }
}
