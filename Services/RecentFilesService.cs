using System.Text.Json;

namespace PptxAvalonia.Services;

/// <summary>Persists recent .pptx paths under the user app data folder.</summary>
public sealed class RecentFilesService
{
    private const int MaxItems = 12;
    private readonly string _storePath;
    private readonly List<string> _paths = [];

    public RecentFilesService()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "PptxAvalonia");
        Directory.CreateDirectory(dir);
        _storePath = Path.Combine(dir, "recent.json");
        Load();
    }

    public IReadOnlyList<string> Paths => _paths.Where(File.Exists).ToList();

    public void Add(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return;

        path = Path.GetFullPath(path);
        _paths.RemoveAll(p => string.Equals(p, path, StringComparison.OrdinalIgnoreCase));
        _paths.Insert(0, path);
        while (_paths.Count > MaxItems)
            _paths.RemoveAt(_paths.Count - 1);
        Save();
    }

    public void Clear()
    {
        _paths.Clear();
        Save();
    }

    private void Load()
    {
        try
        {
            if (!File.Exists(_storePath))
                return;
            var json = File.ReadAllText(_storePath);
            var list = JsonSerializer.Deserialize<List<string>>(json);
            if (list is null) return;
            _paths.Clear();
            _paths.AddRange(list.Where(p => !string.IsNullOrWhiteSpace(p)));
        }
        catch
        {
            // ignore corrupt store
        }
    }

    private void Save()
    {
        try
        {
            var json = JsonSerializer.Serialize(_paths, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_storePath, json);
        }
        catch
        {
            // ignore
        }
    }
}
