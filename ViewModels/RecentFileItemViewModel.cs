using CommunityToolkit.Mvvm.Input;

namespace PptxAvalonia.ViewModels;

/// <summary>One entry in the Recent Files menu (path + open command).</summary>
public sealed class RecentFileItemViewModel
{
    public RecentFileItemViewModel(string path, IAsyncRelayCommand<string?> openCommand)
    {
        Path = path;
        DisplayName = System.IO.Path.GetFileName(path);
        ToolTip = path;
        OpenCommand = openCommand;
    }

    public string Path { get; }
    public string DisplayName { get; }
    public string ToolTip { get; }
    public IAsyncRelayCommand<string?> OpenCommand { get; }
}
