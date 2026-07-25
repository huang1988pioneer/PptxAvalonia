using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

var outPath = args.Length > 0
    ? args[0]
    : Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "Samples", "demo.pptx"));

Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outPath))!);

using (var doc = PresentationDocument.Create(outPath, PresentationDocumentType.Presentation))
{
    var presentationPart = doc.AddPresentationPart();
    presentationPart.Presentation = new Presentation();

    var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>("rId1");
    slideMasterPart.SlideMaster = new SlideMaster(
        new CommonSlideData(CreateEmptyTree()),
        new P.ColorMap
        {
            Background1 = A.ColorSchemeIndexValues.Light1,
            Text1 = A.ColorSchemeIndexValues.Dark1,
            Background2 = A.ColorSchemeIndexValues.Light2,
            Text2 = A.ColorSchemeIndexValues.Dark2,
            Accent1 = A.ColorSchemeIndexValues.Accent1,
            Accent2 = A.ColorSchemeIndexValues.Accent2,
            Accent3 = A.ColorSchemeIndexValues.Accent3,
            Accent4 = A.ColorSchemeIndexValues.Accent4,
            Accent5 = A.ColorSchemeIndexValues.Accent5,
            Accent6 = A.ColorSchemeIndexValues.Accent6,
            Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
            FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
        });
    slideMasterPart.SlideMaster.Save();

    var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>("rId1");
    slideLayoutPart.SlideLayout = new SlideLayout(
        new CommonSlideData(CreateEmptyTree()),
        new ColorMapOverride(new A.MasterColorMapping()));
    slideLayoutPart.SlideLayout.Save();

    var slideIdList = new SlideIdList();
    uint slideId = 256;
    var rel = 2;

    void AddSlide(string title, string body, string accentHex, Action<ShapeTree>? extra = null)
    {
        var relId = "rId" + rel++;
        var slidePart = presentationPart.AddNewPart<SlidePart>(relId);
        slidePart.AddPart(slideLayoutPart);

        var tree = CreateEmptyTree();
        tree.Append(MakeRect(2, 0, 0, 12_192_000, 1_200_000, accentHex, null));
        tree.Append(MakeTextShape(3, 500_000, 250_000, 11_000_000, 900_000, title, 40, true, "FFFFFF",
            A.TextAlignmentTypeValues.Left));
        tree.Append(MakeRect(4, 800_000, 1_800_000, 10_500_000, 4_200_000, "FFFFFF", "D0D0D0"));
        tree.Append(MakeTextShape(5, 1_100_000, 2_100_000, 9_900_000, 3_600_000, body, 22, false, "333333",
            A.TextAlignmentTypeValues.Left));
        extra?.Invoke(tree);

        slidePart.Slide = new Slide(
            new CommonSlideData(tree),
            new ColorMapOverride(new A.MasterColorMapping()));
        slidePart.Slide.Save();
        slideIdList.Append(new SlideId { Id = slideId++, RelationshipId = relId });
    }

    AddSlide(
        "Avalonia PPTX 預覽",
        "這是示範簡報第一頁。\n\n• 開啟 .pptx 檔案\n• 左側縮圖瀏覽\n• 鍵盤 ← → 切換投影片\n• Ctrl + / - 縮放",
        "2F6FED");

    AddSlide(
        "功能重點",
        "解析引擎以 DocumentFormat.OpenXml 讀取投影片，\n再用 Avalonia 控制項重繪：\n\n1. 文字與段落對齊\n2. 矩形 / 圓角 / 橢圓 / 三角形\n3. 內嵌圖片\n4. 主題色與填色",
        "0F9D58");

    AddSlide(
        "開始使用",
        "dotnet run --project PptxAvalonia.csproj\n\n或：\n  bin/Release/net8.0/PptxAvalonia.exe Samples/demo.pptx",
        "E37400",
        tree =>
        {
            tree.Append(new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 10, Name = "Circle" },
                    new P.NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = 9_800_000, Y = 4_800_000 },
                        new A.Extents { Cx = 1_600_000, Cy = 1_600_000 }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Ellipse },
                    new A.SolidFill(new A.RgbColorModelHex { Val = "FFF3E0" }),
                    new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "E37400" })) { Width = 25400 }),
                new P.TextBody(
                    new A.BodyProperties { Anchor = A.TextAnchoringTypeValues.Center },
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.ParagraphProperties { Alignment = A.TextAlignmentTypeValues.Center },
                        new A.Run(
                            new A.RunProperties { FontSize = 1800, Bold = true },
                            new A.Text("OK")),
                        new A.EndParagraphRunProperties()))));
        });

    presentationPart.Presentation.SlideIdList = slideIdList;
    presentationPart.Presentation.SlideMasterIdList = new SlideMasterIdList(
        new SlideMasterId { Id = 2147483648U, RelationshipId = "rId1" });
    presentationPart.Presentation.SlideSize = new SlideSize { Cx = 12_192_000, Cy = 6_858_000 };
    presentationPart.Presentation.NotesSize = new NotesSize { Cx = 6_858_000, Cy = 9_144_000 };
    presentationPart.Presentation.Save();
}

Console.WriteLine("Created: " + Path.GetFullPath(outPath));
return;

static ShapeTree CreateEmptyTree() => new(
    new P.NonVisualGroupShapeProperties(
        new P.NonVisualDrawingProperties { Id = 1, Name = "" },
        new P.NonVisualGroupShapeDrawingProperties(),
        new ApplicationNonVisualDrawingProperties()),
    new GroupShapeProperties(new A.TransformGroup()));

static P.Shape MakeRect(uint id, long x, long y, long cx, long cy, string fillHex, string? lineHex)
{
    var spPr = new P.ShapeProperties(
        new A.Transform2D(new A.Offset { X = x, Y = y }, new A.Extents { Cx = cx, Cy = cy }),
        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
        new A.SolidFill(new A.RgbColorModelHex { Val = fillHex }));

    if (lineHex != null)
        spPr.Append(new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = lineHex })) { Width = 12700 });
    else
        spPr.Append(new A.Outline(new A.NoFill()));

    return new P.Shape(
        new P.NonVisualShapeProperties(
            new P.NonVisualDrawingProperties { Id = id, Name = "Shape" + id },
            new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
            new ApplicationNonVisualDrawingProperties()),
        spPr,
        new P.TextBody(
            new A.BodyProperties(),
            new A.ListStyle(),
            new A.Paragraph(new A.EndParagraphRunProperties { Language = "zh-TW" })));
}

static P.Shape MakeTextShape(
    uint id, long x, long y, long cx, long cy,
    string text, int fontPt, bool bold, string colorHex,
    A.TextAlignmentTypeValues align)
{
    // Support multi-line body via \n
    var paragraphs = text.Split('\n').Select(line =>
        new A.Paragraph(
            new A.ParagraphProperties { Alignment = align },
            new A.Run(
                new A.RunProperties(
                    new A.SolidFill(new A.RgbColorModelHex { Val = colorHex }),
                    new A.LatinFont { Typeface = "Segoe UI" },
                    new A.EastAsianFont { Typeface = "Microsoft JhengHei" })
                {
                    FontSize = fontPt * 100,
                    Bold = bold,
                    Dirty = false
                },
                new A.Text(line.Length == 0 ? " " : line)),
            new A.EndParagraphRunProperties { Language = "zh-TW" })).ToArray();

    var body = new P.TextBody(
        new A.BodyProperties { Anchor = A.TextAnchoringTypeValues.Top },
        new A.ListStyle());
    foreach (var p in paragraphs)
        body.Append(p);

    return new P.Shape(
        new P.NonVisualShapeProperties(
            new P.NonVisualDrawingProperties { Id = id, Name = "Text" + id },
            new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
            new ApplicationNonVisualDrawingProperties()),
        new P.ShapeProperties(
            new A.Transform2D(new A.Offset { X = x, Y = y }, new A.Extents { Cx = cx, Cy = cy }),
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
            new A.NoFill(),
            new A.Outline(new A.NoFill())),
        body);
}
