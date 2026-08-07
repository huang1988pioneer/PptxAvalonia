using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PptxAvalonia.Models;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxAvalonia.Services;

/// <summary>
/// Loads a .pptx package into an in-memory model suitable for Avalonia rendering.
/// Coordinates use pixels at 96 DPI (EMU → px: / 9525).
/// </summary>
public sealed class PptxLoader
{
    private const double EmuPerPixel = 9525.0; // 914400 EMU/inch / 96 DPI

    public PptxPresentation Load(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException("PPTX file not found.", filePath);

        using var doc = PresentationDocument.Open(filePath, false);
        var presentationPart = doc.PresentationPart
            ?? throw new InvalidDataException("Invalid PPTX: missing presentation part.");

        var (slideW, slideH) = ReadSlideSize(presentationPart);
        var themeColors = ReadThemeColors(presentationPart);
        var slides = new List<PptxSlide>();

        var slideIds = presentationPart.Presentation?.SlideIdList?.Elements<SlideId>().ToList()
                       ?? [];

        var index = 0;
        foreach (var slideId in slideIds)
        {
            var relId = slideId.RelationshipId?.Value;
            if (string.IsNullOrEmpty(relId))
                continue;

            var slidePart = (SlidePart?)presentationPart.GetPartById(relId);
            if (slidePart is null)
                continue;

            slides.Add(ParseSlide(slidePart, index, themeColors));
            index++;
        }

        return new PptxPresentation
        {
            FilePath = filePath,
            FileName = Path.GetFileName(filePath),
            SlideWidthEmu = slideW,
            SlideHeightEmu = slideH,
            Slides = slides
        };
    }

    private static (long Width, long Height) ReadSlideSize(PresentationPart part)
    {
        var size = part.Presentation?.SlideSize;
        var w = size?.Cx?.Value ?? 12_192_000L;
        var h = size?.Cy?.Value ?? 6_858_000L;
        return (w, h);
    }

    private static Dictionary<string, ColorRgba> ReadThemeColors(PresentationPart presentationPart)
    {
        var map = new Dictionary<string, ColorRgba>(StringComparer.OrdinalIgnoreCase)
        {
            ["dk1"] = ColorRgba.Black,
            ["lt1"] = ColorRgba.White,
            ["dk2"] = ColorRgba.FromHex("1F4E79"),
            ["lt2"] = ColorRgba.FromHex("EEECE1"),
            ["accent1"] = ColorRgba.FromHex("4472C4"),
            ["accent2"] = ColorRgba.FromHex("ED7D31"),
            ["accent3"] = ColorRgba.FromHex("A5A5A5"),
            ["accent4"] = ColorRgba.FromHex("FFC000"),
            ["accent5"] = ColorRgba.FromHex("5B9BD5"),
            ["accent6"] = ColorRgba.FromHex("70AD47"),
            ["hlink"] = ColorRgba.FromHex("0563C1"),
            ["folHlink"] = ColorRgba.FromHex("954F72"),
            ["tx1"] = ColorRgba.Black,
            ["bg1"] = ColorRgba.White,
            ["tx2"] = ColorRgba.FromHex("1F4E79"),
            ["bg2"] = ColorRgba.FromHex("EEECE1"),
        };

        try
        {
            var themePart = presentationPart.ThemePart
                            ?? presentationPart.SlideMasterParts.FirstOrDefault()?.ThemePart;
            if (themePart?.Theme?.ThemeElements?.ColorScheme is not { } scheme)
                return map;

            void Put(string key, A.Color2Type? c)
            {
                var color = ResolveColor2(c);
                if (color is { } col)
                    map[key] = col;
            }

            Put("dk1", scheme.Dark1Color);
            Put("lt1", scheme.Light1Color);
            Put("dk2", scheme.Dark2Color);
            Put("lt2", scheme.Light2Color);
            Put("accent1", scheme.Accent1Color);
            Put("accent2", scheme.Accent2Color);
            Put("accent3", scheme.Accent3Color);
            Put("accent4", scheme.Accent4Color);
            Put("accent5", scheme.Accent5Color);
            Put("accent6", scheme.Accent6Color);
            Put("hlink", scheme.Hyperlink);
            Put("folHlink", scheme.FollowedHyperlinkColor);

            map["tx1"] = map["dk1"];
            map["bg1"] = map["lt1"];
            map["tx2"] = map["dk2"];
            map["bg2"] = map["lt2"];
        }
        catch
        {
            // Theme is optional; fall back to defaults.
        }

        return map;
    }

    private static ColorRgba? ResolveColor2(A.Color2Type? color)
    {
        if (color is null) return null;
        if (color.RgbColorModelHex?.Val?.Value is { } hex)
            return ColorRgba.FromHex(hex);
        if (color.SystemColor?.LastColor?.Value is { } sysHex)
            return ColorRgba.FromHex(sysHex);
        return null;
    }

    private static PptxSlide ParseSlide(
        SlidePart slidePart,
        int index,
        Dictionary<string, ColorRgba> theme)
    {
        var elements = new List<SlideElement>();
        var background = ColorRgba.White;

        // Master → layout → slide layering
        if (slidePart.SlideLayoutPart is { } layoutPart)
        {
            if (layoutPart.SlideMasterPart is { } masterPart)
            {
                var masterEls = new List<SlideElement>();
                CollectShapes(masterPart.SlideMaster?.CommonSlideData?.ShapeTree, masterPart, theme, masterEls);
                elements.AddRange(masterEls);

                if (slidePart.Slide?.CommonSlideData?.Background is null
                    && layoutPart.SlideLayout?.CommonSlideData?.Background is null)
                {
                    background = ReadBackground(
                        masterPart.SlideMaster?.CommonSlideData?.Background, theme, background);
                }
            }

            CollectShapes(layoutPart.SlideLayout?.CommonSlideData?.ShapeTree, layoutPart, theme, elements);

            if (slidePart.Slide?.CommonSlideData?.Background is null)
            {
                background = ReadBackground(
                    layoutPart.SlideLayout?.CommonSlideData?.Background, theme, background);
            }
        }

        background = ReadBackground(slidePart.Slide?.CommonSlideData?.Background, theme, background);
        CollectShapes(slidePart.Slide?.CommonSlideData?.ShapeTree, slidePart, theme, elements);

        var outline = ExtractOutlineText(elements);
        var notes = ExtractNotesText(slidePart);
        var title = ExtractTitle(elements) ?? $"投影片 {index + 1}";

        return new PptxSlide
        {
            Index = index,
            Name = title,
            Background = background,
            Elements = elements,
            OutlineText = outline,
            NotesText = notes
        };
    }

    private static string? ExtractTitle(List<SlideElement> elements)
    {
        foreach (var el in elements)
        {
            if (el is not ShapeElement { Text: { Paragraphs.Count: > 0 } text })
                continue;

            var line = string.Join("", text.Paragraphs
                .SelectMany(p => p.Runs)
                .Select(r => r.Text)).Trim();
            if (line.Length == 0)
                continue;

            // Prefer larger title-like text near the top
            if (el.Y < 200 && text.Paragraphs[0].Runs.Any(r => r.FontSize >= 24))
                return line.Length > 60 ? line[..60] + "…" : line;
        }

        foreach (var el in elements)
        {
            if (el is not ShapeElement { Text: { Paragraphs.Count: > 0 } text })
                continue;
            var line = string.Join("", text.Paragraphs.SelectMany(p => p.Runs).Select(r => r.Text)).Trim();
            if (line.Length > 0)
                return line.Length > 60 ? line[..60] + "…" : line;
        }

        return null;
    }

    private static string ExtractOutlineText(List<SlideElement> elements)
    {
        var sb = new System.Text.StringBuilder();
        foreach (var el in elements)
        {
            if (el is not ShapeElement { Text: { Paragraphs: { Count: > 0 } paras } })
                continue;

            foreach (var p in paras)
            {
                var line = string.Join("", p.Runs.Select(r => r.Text)).Trim();
                if (line.Length == 0) continue;
                if (p.IndentLevel > 0)
                    sb.Append(new string(' ', (int)p.IndentLevel * 2));
                sb.AppendLine(line);
            }
        }

        return sb.ToString().TrimEnd();
    }

    private static string ExtractNotesText(SlidePart slidePart)
    {
        try
        {
            var notesPart = slidePart.NotesSlidePart;
            if (notesPart?.NotesSlide?.CommonSlideData?.ShapeTree is not { } tree)
                return string.Empty;

            var sb = new System.Text.StringBuilder();
            foreach (var shape in tree.Descendants<P.Shape>())
            {
                // Skip slide image placeholder shapes that only mirror the slide
                try
                {
                    var ph = shape.NonVisualShapeProperties?
                        .ApplicationNonVisualDrawingProperties?
                        .PlaceholderShape;
                    var typeName = ph?.Type?.ToString() ?? string.Empty;
                    if (typeName.Contains("SlideImage", StringComparison.OrdinalIgnoreCase) ||
                        typeName.Equals("sldImg", StringComparison.OrdinalIgnoreCase))
                        continue;
                }
                catch
                {
                    // ignore placeholder detection failures
                }

                var body = shape.TextBody;
                if (body is null) continue;

                foreach (var p in body.Elements<A.Paragraph>())
                {
                    var line = string.Join("", p.Descendants<A.Text>().Select(t => t.Text ?? string.Empty)).Trim();
                    if (line.Length == 0) continue;
                    sb.AppendLine(line);
                }
            }

            return sb.ToString().TrimEnd();
        }
        catch
        {
            return string.Empty;
        }
    }

    private static ColorRgba ReadBackground(
        P.Background? bg,
        Dictionary<string, ColorRgba> theme,
        ColorRgba fallback)
    {
        if (bg?.BackgroundProperties is not { } props)
            return fallback;

        var solid = props.GetFirstChild<A.SolidFill>();
        if (solid is not null)
            return ResolveFillColor(solid, theme) ?? fallback;

        return fallback;
    }

    private static void CollectShapes(
        P.ShapeTree? tree,
        OpenXmlPart containerPart,
        Dictionary<string, ColorRgba> theme,
        List<SlideElement> elements)
    {
        if (tree is null) return;

        foreach (var child in tree.ChildElements)
        {
            switch (child)
            {
                case P.Shape shape:
                    if (TryParseShape(shape, theme) is { } se)
                        elements.Add(se);
                    break;

                case P.Picture picture:
                    if (TryParsePicture(picture, containerPart) is { } pe)
                        elements.Add(pe);
                    break;

                case P.ConnectionShape conn:
                    if (TryParseConnection(conn, theme) is { } le)
                        elements.Add(le);
                    break;

                case P.GroupShape group:
                    CollectGroup(group, containerPart, theme, elements);
                    break;
            }
        }
    }

    private static void CollectGroup(
        P.GroupShape group,
        OpenXmlPart containerPart,
        Dictionary<string, ColorRgba> theme,
        List<SlideElement> elements)
    {
        foreach (var child in group.ChildElements)
        {
            switch (child)
            {
                case P.Shape shape:
                    if (TryParseShape(shape, theme) is { } se)
                        elements.Add(se);
                    break;
                case P.Picture picture:
                    if (TryParsePicture(picture, containerPart) is { } pe)
                        elements.Add(pe);
                    break;
                case P.ConnectionShape conn:
                    if (TryParseConnection(conn, theme) is { } le)
                        elements.Add(le);
                    break;
                case P.GroupShape nested:
                    CollectGroup(nested, containerPart, theme, elements);
                    break;
            }
        }
    }

    private static ShapeElement? TryParseShape(P.Shape shape, Dictionary<string, ColorRgba> theme)
    {
        var xfrm = shape.ShapeProperties?.Transform2D
                   ?? shape.ShapeProperties?.GetFirstChild<A.Transform2D>();
        if (xfrm?.Offset is null || xfrm.Extents is null)
            return null;

        var x = EmuToPx(xfrm.Offset.X?.Value ?? 0);
        var y = EmuToPx(xfrm.Offset.Y?.Value ?? 0);
        var w = EmuToPx(xfrm.Extents.Cx?.Value ?? 0);
        var h = EmuToPx(xfrm.Extents.Cy?.Value ?? 0);
        if (w <= 0 || h <= 0)
            return null;

        var rot = (xfrm.Rotation?.Value ?? 0) / 60000.0;

        var prstGeom = shape.ShapeProperties?.GetFirstChild<A.PresetGeometry>();
        var kind = MapShapeKind(prstGeom?.Preset);

        ColorRgba? fill = null;
        ColorRgba? stroke = null;
        double strokeWidth = 0;
        double cornerRadius = kind == ShapeKind.RoundedRectangle ? Math.Min(w, h) * 0.1 : 0;

        if (shape.ShapeProperties is { } spPr)
        {
            if (spPr.GetFirstChild<A.SolidFill>() is { } sf)
                fill = ResolveFillColor(sf, theme);
            else if (spPr.GetFirstChild<A.NoFill>() is not null)
                fill = null;

            if (spPr.GetFirstChild<A.Outline>() is { } ln)
            {
                strokeWidth = ln.Width is not null ? EmuToPx(ln.Width.Value) : 1.0;
                if (strokeWidth < 0.5 && ln.Width is null) strokeWidth = 1.0;
                if (ln.GetFirstChild<A.SolidFill>() is { } lsf)
                    stroke = ResolveFillColor(lsf, theme);
                if (ln.GetFirstChild<A.NoFill>() is not null)
                {
                    stroke = null;
                    strokeWidth = 0;
                }
            }
        }

        var text = ParseTextBody(shape.TextBody, theme);

        // Skip completely invisible shapes (common on masters / empty placeholders)
        var hasText = text is { Paragraphs.Count: > 0 } &&
                      text.Paragraphs.Any(p => p.Runs.Any(r => !string.IsNullOrWhiteSpace(r.Text)));
        if (fill is null && (stroke is null || strokeWidth <= 0) && !hasText)
            return null;

        return new ShapeElement
        {
            X = x,
            Y = y,
            Width = w,
            Height = h,
            Rotation = rot,
            Kind = kind,
            Fill = fill,
            Stroke = stroke,
            StrokeWidth = strokeWidth,
            CornerRadius = cornerRadius,
            Text = text
        };
    }

    private static ImageElement? TryParsePicture(P.Picture picture, OpenXmlPart containerPart)
    {
        var xfrm = picture.ShapeProperties?.Transform2D;
        if (xfrm?.Offset is null || xfrm.Extents is null)
            return null;

        var x = EmuToPx(xfrm.Offset.X?.Value ?? 0);
        var y = EmuToPx(xfrm.Offset.Y?.Value ?? 0);
        var w = EmuToPx(xfrm.Extents.Cx?.Value ?? 0);
        var h = EmuToPx(xfrm.Extents.Cy?.Value ?? 0);
        if (w <= 0 || h <= 0)
            return null;

        var embed = picture.BlipFill?.Blip?.Embed?.Value;
        if (string.IsNullOrEmpty(embed))
            return null;

        try
        {
            var part = containerPart.GetPartById(embed);
            using var stream = part.GetStream();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            var bytes = ms.ToArray();
            if (bytes.Length == 0)
                return null;

            return new ImageElement
            {
                X = x,
                Y = y,
                Width = w,
                Height = h,
                Rotation = (xfrm.Rotation?.Value ?? 0) / 60000.0,
                ImageBytes = bytes,
                ContentType = part.ContentType
            };
        }
        catch
        {
            return null;
        }
    }

    private static LineElement? TryParseConnection(P.ConnectionShape conn, Dictionary<string, ColorRgba> theme)
    {
        var xfrm = conn.ShapeProperties?.Transform2D;
        if (xfrm?.Offset is null || xfrm.Extents is null)
            return null;

        var x = EmuToPx(xfrm.Offset.X?.Value ?? 0);
        var y = EmuToPx(xfrm.Offset.Y?.Value ?? 0);
        var w = EmuToPx(xfrm.Extents.Cx?.Value ?? 0);
        var h = EmuToPx(xfrm.Extents.Cy?.Value ?? 0);

        var stroke = ColorRgba.Black;
        double strokeWidth = 1.5;
        if (conn.ShapeProperties?.GetFirstChild<A.Outline>() is { } ln)
        {
            strokeWidth = ln.Width is not null ? Math.Max(0.5, EmuToPx(ln.Width.Value)) : 1.5;
            if (ln.GetFirstChild<A.SolidFill>() is { } sf)
                stroke = ResolveFillColor(sf, theme) ?? stroke;
        }

        return new LineElement
        {
            X = x,
            Y = y,
            Width = w,
            Height = h,
            X2 = x + w,
            Y2 = y + h,
            Stroke = stroke,
            StrokeWidth = strokeWidth
        };
    }

    private static TextContent? ParseTextBody(P.TextBody? body, Dictionary<string, ColorRgba> theme)
    {
        if (body is null) return null;

        var paragraphs = new List<TextParagraph>();
        var vAlign = VerticalAlign.Top;

        var bodyPr = body.BodyProperties;
        if (bodyPr?.Anchor is not null)
        {
            if (bodyPr.Anchor == A.TextAnchoringTypeValues.Center)
                vAlign = VerticalAlign.Middle;
            else if (bodyPr.Anchor == A.TextAnchoringTypeValues.Bottom)
                vAlign = VerticalAlign.Bottom;
        }

        foreach (var p in body.Elements<A.Paragraph>())
        {
            var align = HorizontalAlign.Left;
            var pPr = p.ParagraphProperties;
            if (pPr?.Alignment is not null)
            {
                if (pPr.Alignment == A.TextAlignmentTypeValues.Center)
                    align = HorizontalAlign.Center;
                else if (pPr.Alignment == A.TextAlignmentTypeValues.Right)
                    align = HorizontalAlign.Right;
                else if (pPr.Alignment == A.TextAlignmentTypeValues.Justified)
                    align = HorizontalAlign.Justify;
            }

            var level = p.ParagraphProperties?.Level?.Value ?? 0;
            var runs = new List<TextRun>();

            foreach (var child in p.ChildElements)
            {
                if (child is A.Run run)
                {
                    var text = run.Text?.Text ?? string.Empty;
                    if (text.Length == 0) continue;

                    var rPr = run.RunProperties;
                    var fontSize = rPr?.FontSize is not null ? rPr.FontSize.Value / 100.0 : 18.0;
                    var bold = rPr?.Bold?.Value ?? false;
                    var italic = rPr?.Italic?.Value ?? false;
                    var underline = rPr?.Underline?.Value is not null
                                    && rPr.Underline.Value != A.TextUnderlineValues.None;

                    var color = ColorRgba.Black;
                    if (rPr?.GetFirstChild<A.SolidFill>() is { } sf)
                        color = ResolveFillColor(sf, theme) ?? color;

                    var font = rPr?.GetFirstChild<A.LatinFont>()?.Typeface?.Value
                               ?? rPr?.GetFirstChild<A.EastAsianFont>()?.Typeface?.Value
                               ?? "Segoe UI";

                    runs.Add(new TextRun
                    {
                        Text = text,
                        FontSize = fontSize,
                        FontFamily = font,
                        Bold = bold,
                        Italic = italic,
                        Underline = underline,
                        Color = color
                    });
                }
                else if (child is A.Break)
                {
                    runs.Add(new TextRun { Text = "\n", FontSize = 18 });
                }
                else if (child is A.Field field)
                {
                    var text = field.GetFirstChild<A.Text>()?.Text ?? string.Empty;
                    if (text.Length == 0) continue;
                    runs.Add(new TextRun { Text = text, FontSize = 12, Color = ColorRgba.FromHex("666666") });
                }
            }

            if (runs.Count == 0)
                runs.Add(new TextRun { Text = " ", FontSize = 12 });

            paragraphs.Add(new TextParagraph
            {
                Align = align,
                IndentLevel = level,
                Runs = runs
            });
        }

        if (paragraphs.Count == 0)
            return null;

        return new TextContent
        {
            Paragraphs = paragraphs,
            VerticalAlign = vAlign
        };
    }

    private static ColorRgba? ResolveFillColor(A.SolidFill solid, Dictionary<string, ColorRgba> theme)
    {
        if (solid.RgbColorModelHex?.Val?.Value is { } hex)
            return ApplyTransforms(ColorRgba.FromHex(hex), solid.RgbColorModelHex);

        if (solid.SchemeColor is { } sc)
        {
            var key = sc.Val?.ToString() ?? string.Empty;
            key = NormalizeSchemeKey(key);
            if (theme.TryGetValue(key, out var baseColor))
                return ApplyTransforms(baseColor, sc);
        }

        if (solid.SystemColor?.LastColor?.Value is { } sysHex)
            return ColorRgba.FromHex(sysHex);

        if (solid.PresetColor?.Val is { } preset)
        {
            var name = preset.ToString() ?? string.Empty;
            if (name.Equals("Black", StringComparison.OrdinalIgnoreCase)) return ColorRgba.Black;
            if (name.Equals("White", StringComparison.OrdinalIgnoreCase)) return ColorRgba.White;
            if (name.Equals("Red", StringComparison.OrdinalIgnoreCase)) return ColorRgba.FromHex("FF0000");
            if (name.Equals("Blue", StringComparison.OrdinalIgnoreCase)) return ColorRgba.FromHex("0000FF");
            if (name.Equals("Green", StringComparison.OrdinalIgnoreCase)) return ColorRgba.FromHex("00B050");
            if (name.Equals("Gray", StringComparison.OrdinalIgnoreCase)) return ColorRgba.FromHex("808080");
            return ColorRgba.Black;
        }

        return null;
    }

    private static string NormalizeSchemeKey(string key)
    {
        key = key.Trim();
        return key.ToLowerInvariant() switch
        {
            "dark1" => "dk1",
            "dark2" => "dk2",
            "light1" => "lt1",
            "light2" => "lt2",
            "text1" => "tx1",
            "text2" => "tx2",
            "background1" => "bg1",
            "background2" => "bg2",
            "hyperlink" => "hlink",
            "followedhyperlink" => "folHlink",
            var s => s
        };
    }

    private static ColorRgba ApplyTransforms(ColorRgba color, OpenXmlElement? colorEl)
    {
        if (colorEl is null) return color;

        double r = color.R / 255.0;
        double g = color.G / 255.0;
        double b = color.B / 255.0;
        double a = color.A / 255.0;

        foreach (var child in colorEl.ChildElements)
        {
            switch (child)
            {
                case A.LuminanceModulation lumMod when lumMod.Val is not null:
                {
                    var f = lumMod.Val.Value / 100_000.0;
                    r *= f; g *= f; b *= f;
                    break;
                }
                case A.LuminanceOffset lumOff when lumOff.Val is not null:
                {
                    var f = lumOff.Val.Value / 100_000.0;
                    r += (1 - r) * f;
                    g += (1 - g) * f;
                    b += (1 - b) * f;
                    break;
                }
                case A.Alpha alpha when alpha.Val is not null:
                    a = alpha.Val.Value / 100_000.0;
                    break;
                case A.Shade shade when shade.Val is not null:
                {
                    var f = shade.Val.Value / 100_000.0;
                    r *= f; g *= f; b *= f;
                    break;
                }
                case A.Tint tint when tint.Val is not null:
                {
                    var f = tint.Val.Value / 100_000.0;
                    r = r * f + (1 - f);
                    g = g * f + (1 - f);
                    b = b * f + (1 - f);
                    break;
                }
            }
        }

        static byte Clamp(double v) => (byte)Math.Clamp((int)Math.Round(v * 255), 0, 255);
        return new ColorRgba(Clamp(r), Clamp(g), Clamp(b), Clamp(a));
    }

    private static ShapeKind MapShapeKind(DocumentFormat.OpenXml.EnumValue<A.ShapeTypeValues>? prst)
    {
        if (prst is null || !prst.HasValue)
            return ShapeKind.Rectangle;

        if (prst == A.ShapeTypeValues.Ellipse)
            return ShapeKind.Ellipse;
        if (prst == A.ShapeTypeValues.RoundRectangle)
            return ShapeKind.RoundedRectangle;
        if (prst == A.ShapeTypeValues.Triangle || prst == A.ShapeTypeValues.RightTriangle)
            return ShapeKind.Triangle;
        if (prst == A.ShapeTypeValues.Diamond)
            return ShapeKind.Diamond;
        if (prst == A.ShapeTypeValues.Rectangle)
            return ShapeKind.Rectangle;

        return ShapeKind.Other;
    }

    private static double EmuToPx(long emu) => emu / EmuPerPixel;
}
