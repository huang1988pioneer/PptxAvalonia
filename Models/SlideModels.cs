using Avalonia.Media.Imaging;

namespace PptxAvalonia.Models;

/// <summary>In-memory presentation model loaded from a .pptx package.</summary>
public sealed class PptxPresentation
{
    public required string FilePath { get; init; }
    public required string FileName { get; init; }
    public long SlideWidthEmu { get; init; } = 12_192_000; // 13.333" default 16:9
    public long SlideHeightEmu { get; init; } = 6_858_000; // 7.5"
    public IReadOnlyList<PptxSlide> Slides { get; init; } = [];
}

public sealed class PptxSlide
{
    public required int Index { get; init; }
    public required string Name { get; init; }
    public ColorRgba Background { get; init; } = ColorRgba.White;
    public string? BackgroundImagePart { get; init; }
    public IReadOnlyList<SlideElement> Elements { get; init; } = [];
    /// <summary>Speaker notes (備忘稿) plain text.</summary>
    public string NotesText { get; init; } = string.Empty;
    /// <summary>All text content concatenated for outline / find.</summary>
    public string OutlineText { get; init; } = string.Empty;
}

public enum AppViewMode
{
    Normal,
    Sorter,
    Outline
}

public abstract class SlideElement
{
    public double X { get; init; }
    public double Y { get; init; }
    public double Width { get; init; }
    public double Height { get; init; }
    public double Rotation { get; init; }
}

public sealed class ShapeElement : SlideElement
{
    public ShapeKind Kind { get; init; } = ShapeKind.Rectangle;
    public ColorRgba? Fill { get; init; }
    public ColorRgba? Stroke { get; init; }
    public double StrokeWidth { get; init; } = 1;
    public double CornerRadius { get; init; }
    public TextContent? Text { get; init; }
}

public sealed class ImageElement : SlideElement
{
    public required byte[] ImageBytes { get; init; }
    public string? ContentType { get; init; }
    public Bitmap? CachedBitmap { get; set; }
}

public sealed class LineElement : SlideElement
{
    public double X2 { get; init; }
    public double Y2 { get; init; }
    public ColorRgba Stroke { get; init; } = ColorRgba.Black;
    public double StrokeWidth { get; init; } = 1.5;
}

public sealed class TextContent
{
    public IReadOnlyList<TextParagraph> Paragraphs { get; init; } = [];
    public VerticalAlign VerticalAlign { get; init; } = VerticalAlign.Top;
}

public sealed class TextParagraph
{
    public HorizontalAlign Align { get; init; } = HorizontalAlign.Left;
    public double SpaceBefore { get; init; }
    public double SpaceAfter { get; init; }
    public double IndentLevel { get; init; }
    public IReadOnlyList<TextRun> Runs { get; init; } = [];
}

public sealed class TextRun
{
    public string Text { get; init; } = string.Empty;
    public double FontSize { get; init; } = 18;
    public string FontFamily { get; init; } = "Segoe UI";
    public bool Bold { get; init; }
    public bool Italic { get; init; }
    public bool Underline { get; init; }
    public ColorRgba Color { get; init; } = ColorRgba.Black;
}

public enum ShapeKind
{
    Rectangle,
    RoundedRectangle,
    Ellipse,
    Triangle,
    Diamond,
    Other
}

public enum HorizontalAlign
{
    Left,
    Center,
    Right,
    Justify
}

public enum VerticalAlign
{
    Top,
    Middle,
    Bottom
}

public readonly record struct ColorRgba(byte R, byte G, byte B, byte A = 255)
{
    public static ColorRgba White => new(255, 255, 255);
    public static ColorRgba Black => new(0, 0, 0);
    public static ColorRgba Transparent => new(0, 0, 0, 0);

    public static ColorRgba FromHex(string hex)
    {
        hex = hex.TrimStart('#');
        if (hex.Length == 6)
        {
            return new ColorRgba(
                Convert.ToByte(hex[..2], 16),
                Convert.ToByte(hex[2..4], 16),
                Convert.ToByte(hex[4..6], 16));
        }

        if (hex.Length == 8)
        {
            return new ColorRgba(
                Convert.ToByte(hex[2..4], 16),
                Convert.ToByte(hex[4..6], 16),
                Convert.ToByte(hex[6..8], 16),
                Convert.ToByte(hex[..2], 16));
        }

        return Black;
    }

    public Avalonia.Media.Color ToAvalonia() => Avalonia.Media.Color.FromArgb(A, R, G, B);
}
