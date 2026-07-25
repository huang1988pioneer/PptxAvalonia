using Avalonia;
using Avalonia.Controls;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using PptxAvalonia.Models;

namespace PptxAvalonia.Services;

/// <summary>Renders slides to PNG images for export.</summary>
public sealed class SlideExportService
{
    private readonly SlideRenderer _renderer = new();

    public async Task ExportSlidePngAsync(
        PptxPresentation presentation,
        PptxSlide slide,
        string filePath,
        double scale = 1.0)
    {
        await Dispatcher.UIThread.InvokeAsync(() =>
        {
            var control = _renderer.BuildSlide(presentation, slide, scale: 1.0);
            var size = _renderer.GetSlidePixelSize(presentation);
            var pixelW = Math.Max(1, (int)Math.Round(size.Width * scale));
            var pixelH = Math.Max(1, (int)Math.Round(size.Height * scale));

            control.Width = size.Width;
            control.Height = size.Height;
            control.Measure(new Size(size.Width, size.Height));
            control.Arrange(new Rect(0, 0, size.Width, size.Height));
            control.UpdateLayout();

            var bitmap = new RenderTargetBitmap(
                new PixelSize(pixelW, pixelH),
                new Vector(96 * scale, 96 * scale));
            bitmap.Render(control);

            bitmap.Save(filePath);
        });
    }

    public async Task ExportAllAsync(
        PptxPresentation presentation,
        string directory,
        double scale = 1.0,
        IProgress<int>? progress = null)
    {
        Directory.CreateDirectory(directory);
        for (var i = 0; i < presentation.Slides.Count; i++)
        {
            var path = Path.Combine(directory, $"slide-{i + 1:D3}.png");
            await ExportSlidePngAsync(presentation, presentation.Slides[i], path, scale);
            progress?.Report(i + 1);
        }
    }
}
