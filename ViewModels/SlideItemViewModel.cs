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
    }

    public PptxSlide Slide { get; }
    public Control Thumbnail { get; }
    public int DisplayNumber { get; }

    [ObservableProperty]
    private string _title;

    [ObservableProperty]
    private bool _isSelected;
}
