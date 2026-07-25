using System.IO;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Controls.Shapes;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using PptxAvalonia.Models;
using AvPath = Avalonia.Controls.Shapes.Path;

namespace PptxAvalonia.Services;

/// <summary>
/// Builds Avalonia visual trees that approximate PowerPoint slides.
/// </summary>
public sealed class SlideRenderer
{
    private const double EmuPerPixel = 9525.0;

    public Size GetSlidePixelSize(PptxPresentation presentation)
    {
        var w = presentation.SlideWidthEmu / EmuPerPixel;
        var h = presentation.SlideHeightEmu / EmuPerPixel;
        return new Size(w, h);
    }

    public Control BuildSlide(PptxPresentation presentation, PptxSlide slide, double scale = 1.0)
    {
        var size = GetSlidePixelSize(presentation);
        var canvas = new Canvas
        {
            Width = size.Width,
            Height = size.Height,
            Background = new SolidColorBrush(slide.Background.ToAvalonia()),
            ClipToBounds = true
        };

        foreach (var element in slide.Elements)
        {
            var visual = CreateElement(element);
            if (visual is null) continue;
            canvas.Children.Add(visual);
        }

        if (Math.Abs(scale - 1.0) < 0.001)
            return canvas;

        var viewbox = new Viewbox
        {
            Width = size.Width * scale,
            Height = size.Height * scale,
            Stretch = Stretch.Fill,
            Child = canvas
        };
        return viewbox;
    }

    public Control BuildThumbnail(PptxPresentation presentation, PptxSlide slide, double maxWidth = 160)
    {
        var size = GetSlidePixelSize(presentation);
        var scale = maxWidth / size.Width;
        return BuildSlide(presentation, slide, scale);
    }

    private Control? CreateElement(SlideElement element)
    {
        return element switch
        {
            ShapeElement shape => CreateShape(shape),
            ImageElement image => CreateImage(image),
            LineElement line => CreateLine(line),
            _ => null
        };
    }

    private Control CreateShape(ShapeElement shape)
    {
        var root = new Canvas
        {
            Width = shape.Width,
            Height = shape.Height
        };
        Canvas.SetLeft(root, shape.X);
        Canvas.SetTop(root, shape.Y);

        if (Math.Abs(shape.Rotation) > 0.01)
        {
            root.RenderTransformOrigin = new RelativePoint(0.5, 0.5, RelativeUnit.Relative);
            root.RenderTransform = new RotateTransform(shape.Rotation);
        }

        var geometry = CreateGeometry(shape);
        if (geometry is not null)
        {
            var path = new AvPath
            {
                Data = geometry,
                Width = shape.Width,
                Height = shape.Height,
                Stretch = Stretch.Fill
            };

            if (shape.Fill is { } fill)
                path.Fill = new SolidColorBrush(fill.ToAvalonia());

            if (shape.Stroke is { } stroke && shape.StrokeWidth > 0)
            {
                path.Stroke = new SolidColorBrush(stroke.ToAvalonia());
                path.StrokeThickness = shape.StrokeWidth;
            }

            root.Children.Add(path);
        }
        else if (shape.Fill is { } fillOnly)
        {
            // Fallback rectangle
            var rect = new Border
            {
                Width = shape.Width,
                Height = shape.Height,
                Background = new SolidColorBrush(fillOnly.ToAvalonia()),
                CornerRadius = shape.Kind == ShapeKind.RoundedRectangle
                    ? new CornerRadius(shape.CornerRadius)
                    : default
            };
            if (shape.Stroke is { } stroke && shape.StrokeWidth > 0)
            {
                rect.BorderBrush = new SolidColorBrush(stroke.ToAvalonia());
                rect.BorderThickness = new Thickness(shape.StrokeWidth);
            }
            root.Children.Add(rect);
        }

        if (shape.Text is { } text && text.Paragraphs.Count > 0)
        {
            var textBlock = BuildText(text, shape.Width, shape.Height);
            root.Children.Add(textBlock);
        }

        return root;
    }

    private static Geometry? CreateGeometry(ShapeElement shape)
    {
        var w = shape.Width;
        var h = shape.Height;

        return shape.Kind switch
        {
            ShapeKind.Ellipse => new EllipseGeometry(new Rect(0, 0, w, h)),
            ShapeKind.RoundedRectangle => new RectangleGeometry(
                new Rect(0, 0, w, h), shape.CornerRadius, shape.CornerRadius),
            ShapeKind.Triangle => Geometry.Parse($"M {w / 2},0 L {w},{h} L 0,{h} Z"),
            ShapeKind.Diamond => Geometry.Parse($"M {w / 2},0 L {w},{h / 2} L {w / 2},{h} L 0,{h / 2} Z"),
            ShapeKind.Rectangle or ShapeKind.Other => new RectangleGeometry(new Rect(0, 0, w, h)),
            _ => new RectangleGeometry(new Rect(0, 0, w, h))
        };
    }

    private Control CreateImage(ImageElement image)
    {
        var bitmap = image.CachedBitmap;
        if (bitmap is null)
        {
            try
            {
                using var ms = new MemoryStream(image.ImageBytes);
                bitmap = new Bitmap(ms);
                image.CachedBitmap = bitmap;
            }
            catch
            {
                var placeholder = new Border
                {
                    Width = image.Width,
                    Height = image.Height,
                    Background = new SolidColorBrush(Color.FromRgb(230, 230, 230)),
                    Child = new TextBlock
                    {
                        Text = "圖片",
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center,
                        Foreground = Brushes.Gray
                    }
                };
                Canvas.SetLeft(placeholder, image.X);
                Canvas.SetTop(placeholder, image.Y);
                return placeholder;
            }
        }

        var img = new Image
        {
            Source = bitmap,
            Width = image.Width,
            Height = image.Height,
            Stretch = Stretch.Fill
        };

        if (Math.Abs(image.Rotation) > 0.01)
        {
            img.RenderTransformOrigin = new RelativePoint(0.5, 0.5, RelativeUnit.Relative);
            img.RenderTransform = new RotateTransform(image.Rotation);
        }

        Canvas.SetLeft(img, image.X);
        Canvas.SetTop(img, image.Y);
        return img;
    }

    private static Control CreateLine(LineElement line)
    {
        var path = new Avalonia.Controls.Shapes.Line
        {
            StartPoint = new Point(0, 0),
            EndPoint = new Point(line.Width, line.Height),
            Stroke = new SolidColorBrush(line.Stroke.ToAvalonia()),
            StrokeThickness = line.StrokeWidth,
            Width = Math.Max(line.Width, 1),
            Height = Math.Max(line.Height, 1)
        };
        Canvas.SetLeft(path, line.X);
        Canvas.SetTop(path, line.Y);
        return path;
    }

    private static Control BuildText(TextContent text, double width, double height)
    {
        var panel = new StackPanel
        {
            Orientation = Orientation.Vertical,
            Width = Math.Max(0, width - 8),
            Margin = new Thickness(4)
        };

        foreach (var para in text.Paragraphs)
        {
            var tb = new TextBlock
            {
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(para.IndentLevel * 18, 0, 0, 2),
                TextAlignment = para.Align switch
                {
                    HorizontalAlign.Center => TextAlignment.Center,
                    HorizontalAlign.Right => TextAlignment.Right,
                    HorizontalAlign.Justify => TextAlignment.Justify,
                    _ => TextAlignment.Left
                }
            };

            foreach (var run in para.Runs)
            {
                var inline = new Run(run.Text)
                {
                    FontSize = Math.Max(6, run.FontSize),
                    FontFamily = new FontFamily(run.FontFamily),
                    FontWeight = run.Bold ? FontWeight.Bold : FontWeight.Normal,
                    FontStyle = run.Italic ? FontStyle.Italic : FontStyle.Normal,
                    Foreground = new SolidColorBrush(run.Color.ToAvalonia())
                };

                if (run.Underline)
                    inline.TextDecorations = TextDecorations.Underline;

                tb.Inlines?.Add(inline);
            }

            panel.Children.Add(tb);
        }

        var host = new Border
        {
            Width = width,
            Height = height,
            Child = new ScrollViewer
            {
                HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
                VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
                Content = AlignTextPanel(panel, text.VerticalAlign)
            }
        };

        return host;
    }

    private static Control AlignTextPanel(Control panel, VerticalAlign align)
    {
        return align switch
        {
            VerticalAlign.Middle => new Grid
            {
                Children =
                {
                    new ContentControl
                    {
                        Content = panel,
                        VerticalAlignment = VerticalAlignment.Center,
                        HorizontalAlignment = HorizontalAlignment.Stretch
                    }
                }
            },
            VerticalAlign.Bottom => new Grid
            {
                Children =
                {
                    new ContentControl
                    {
                        Content = panel,
                        VerticalAlignment = VerticalAlignment.Bottom,
                        HorizontalAlignment = HorizontalAlignment.Stretch
                    }
                }
            },
            _ => panel
        };
    }
}
