using Avalonia.Controls;
using CommunityToolkit.Mvvm.ComponentModel;
using PptxAvalonia.Models;

namespace PptxAvalonia.ViewModels;

public partial class SlideItemViewModel : ViewModelBase
{
    public SlideItemViewModel(PptxSlide slide, Control thumbnail, int displayNumber)
    {
        Slide = slide;
        Thumbnail = thumbnail;
        DisplayNumber = displayNumber;
        Title = slide.Name;
        OutlineText = slide.OutlineText;
        NotesText = slide.NotesText;
        HasNotes = !string.IsNullOrWhiteSpace(slide.NotesText);
    }

    public PptxSlide Slide { get; }
    public Control Thumbnail { get; }
    public int DisplayNumber { get; }
    public string OutlineText { get; }
    public string NotesText { get; }
    public bool HasNotes { get; }

    [ObservableProperty]
    private string _title;

    [ObservableProperty]
    private bool _isSelected;

    [ObservableProperty]
    private bool _isFindMatch;
}
