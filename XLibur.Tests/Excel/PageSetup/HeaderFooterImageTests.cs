using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using XLibur.Excel;
using XLibur.Excel.IO;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.PageSetup;

public class HeaderFooterImageTests
{
    private static readonly string[] ShapeIdsLH = ["LH"];
    private static readonly string[] ShapeIdsCH = ["CH"];
    private static readonly string[] ShapeIdsRH = ["RH"];
    private static readonly string[] ShapeIdsLF = ["LF"];

    private static string _pngPath = null!;
    private static string _jpegPath = null!;
    private static string _tempDir = null!;

    [Before(HookType.Class)]
    public static async Task Setup()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "XLibur_HFImageTests_" + Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(_tempDir);

        _pngPath = Path.Combine(_tempDir, "test.png");
        _jpegPath = Path.Combine(_tempDir, "test.jpg");

        await ExtractResource("XLibur.Tests.Resource.Images.SampleImagePng.png", _pngPath);
        await ExtractResource("XLibur.Tests.Resource.Images.SampleImageExif.jpg", _jpegPath);
    }

    [After(HookType.Class)]
    public static void Cleanup()
    {
        try { Directory.Delete(_tempDir, true); }
        catch { /* best effort cleanup */ }
    }

    [Test]
    public async Task AddImage_LeftHeader_EmitsGraphicMarkerAndVml()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.PageSetup.Header.Left.AddImage(_pngPath, XLHFOccurrence.AllPages);

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        ms.Position = 0;
        await AssertPackageContainsHFImage(ms, expectedShapeIds: ShapeIdsLH, expectedHeaderText: "&L&G");
    }

    [Test]
    public async Task AddImage_CenterHeader_Works()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.PageSetup.Header.Center.AddImage(_jpegPath, XLHFOccurrence.AllPages);

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        ms.Position = 0;
        await AssertPackageContainsHFImage(ms, expectedShapeIds: ShapeIdsCH, expectedHeaderText: "&C&G");
    }

    [Test]
    public async Task AddImage_RightHeader_Works()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.PageSetup.Header.Right.AddImage(_pngPath, XLHFOccurrence.AllPages);

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        ms.Position = 0;
        await AssertPackageContainsHFImage(ms, expectedShapeIds: ShapeIdsRH, expectedHeaderText: "&R&G");
    }

    [Test]
    public async Task AddImage_LeftFooter_Works()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.PageSetup.Footer.Left.AddImage(_pngPath, XLHFOccurrence.AllPages);

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        ms.Position = 0;
        await AssertPackageContainsHFImage(ms, expectedShapeIds: ShapeIdsLF, expectedFooterText: "&L&G");
    }

    [Test]
    public async Task AddImage_MixedTextAndImage_PreservesOrdering()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.PageSetup.Header.Left.AddText("Before:", XLHFOccurrence.AllPages);
        ws.PageSetup.Header.Left.AddImage(_pngPath, XLHFOccurrence.AllPages);
        ws.PageSetup.Header.Left.AddText(":After", XLHFOccurrence.AllPages);

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        ms.Position = 0;
        await AssertPackageContainsHFImage(ms, expectedShapeIds: ShapeIdsLH, expectedHeaderText: "&LBefore:&G:After");
    }

    [Test]
    public async Task AddImage_MultipleSections_HeaderAndFooter()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.PageSetup.Header.Left.AddImage(_pngPath, XLHFOccurrence.AllPages);
        ws.PageSetup.Header.Center.AddText("Title", XLHFOccurrence.AllPages);
        ws.PageSetup.Header.Right.AddImage(_jpegPath, XLHFOccurrence.AllPages);
        ws.PageSetup.Footer.Center.AddImage(_pngPath, XLHFOccurrence.AllPages);

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        ms.Position = 0;
        var (vmlXml, headerText, footerText, contentTypes, mediaFiles) = ExtractPackageInfo(ms);

        await Assert.That(headerText).IsEqualTo("&L&G&CTitle&R&G");
        await Assert.That(footerText).IsEqualTo("&C&G");

        // VML should have shapes LH, RH, CF
        var shapeIds = GetVmlShapeIds(vmlXml);
        await Assert.That(shapeIds).Contains("LH");
        await Assert.That(shapeIds).Contains("RH");
        await Assert.That(shapeIds).Contains("CF");

        // Should have image media files
        await Assert.That(mediaFiles.Count).IsGreaterThanOrEqualTo(2);
    }

    [Test]
    public async Task AddImage_InvalidPath_ThrowsFileNotFound()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        await Assert.That(() =>
            ws.PageSetup.Header.Left.AddImage(@"C:\nonexistent\image.png")).Throws<FileNotFoundException>();
    }

    [Test]
    public async Task AddImage_NullOrEmptyPath_ThrowsArgumentException()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        await Assert.That(() =>
            ws.PageSetup.Header.Left.AddImage("")).Throws<ArgumentException>();

        await Assert.That(() =>
            ws.PageSetup.Header.Left.AddImage("   ")).Throws<ArgumentException>();
    }

    [Test]
    public async Task AddImage_UnsupportedFormat_ThrowsArgumentException()
    {
        var svgPath = Path.Combine(_tempDir, "test.svg");
        File.WriteAllText(svgPath, "<svg></svg>");

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        await Assert.That(() =>
            ws.PageSetup.Header.Left.AddImage(svgPath)).Throws<ArgumentException>();
    }

    [Test]
    public async Task AddImage_ScalingPreservesAspectRatio_DoesNotUpscale()
    {
        // SampleImagePng.png is 252x152 at 96dpi => 2.625" x 1.583"
        // Max: 2.5" x 0.6"
        // Scale factor by width: 2.5/2.625 = 0.952
        // Scale factor by height: 0.6/1.583 = 0.379
        // Use min(0.952, 0.379) = 0.379
        // Result: 2.625*0.379 = ~0.995" = ~71.6pt wide, 0.6" = 43.2pt tall

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.PageSetup.Header.Left.AddImage(_pngPath, XLHFOccurrence.AllPages);

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        ms.Position = 0;
        var (vmlXml, _, _, _, _) = ExtractPackageInfo(ms);

        // Parse shape style to check dimensions
        var shapes = GetVmlShapes(vmlXml);
        var shape = shapes.First(s => s.id == "LH");

        // The width should be less than or equal to 2.5" = 180pt
        await Assert.That(shape.widthPt).IsLessThanOrEqualTo(180.1);
        // The height should be less than or equal to 0.6" = 43.2pt
        await Assert.That(shape.heightPt).IsLessThanOrEqualTo(43.3);
        // Both should be > 0
        await Assert.That(shape.widthPt).IsGreaterThan(0);
        await Assert.That(shape.heightPt).IsGreaterThan(0);
    }

    [Test]
    public async Task AddImage_PackageHasCorrectContentTypes()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.PageSetup.Header.Left.AddImage(_pngPath, XLHFOccurrence.AllPages);

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        ms.Position = 0;
        var (_, _, _, contentTypes, _) = ExtractPackageInfo(ms);

        // Should have vml and png content types
        await Assert.That(contentTypes).Contains("application/vnd.openxmlformats-officedocument.vmlDrawing");
        await Assert.That(contentTypes).Contains("image/png");
    }

    [Test]
    public async Task AddImage_VmlHasImageRelationships()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.PageSetup.Header.Left.AddImage(_pngPath, XLHFOccurrence.AllPages);

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        ms.Position = 0;
        using var archive = new ZipArchive(ms, ZipArchiveMode.Read);

        // Find VML rels file
        var vmlRelsEntry = archive.Entries.FirstOrDefault(e =>
            e.FullName.Contains("drawings/_rels/") && e.FullName.EndsWith(".rels"));
        await Assert.That(vmlRelsEntry).IsNotNull().Because("VML rels file should exist");

        using var relsStream = vmlRelsEntry!.Open();
        var relsXml = XDocument.Load(relsStream);
        var ns = XNamespace.Get("http://schemas.openxmlformats.org/package/2006/relationships");
        var imageRels = relsXml.Root!.Elements(ns + "Relationship")
            .Where(r => r.Attribute("Type")?.Value.Contains("image") == true)
            .ToList();

        await Assert.That(imageRels.Count).IsGreaterThanOrEqualTo(1).Because("VML should have image relationships");
    }

    [Test]
    public async Task AddImage_SheetHasLegacyDrawingHFElement()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.PageSetup.Header.Left.AddImage(_pngPath, XLHFOccurrence.AllPages);

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        ms.Position = 0;
        using var archive = new ZipArchive(ms, ZipArchiveMode.Read);
        var sheetEntry = archive.Entries.First(e => e.FullName.Contains("worksheets/sheet"));
        using var sheetStream = sheetEntry.Open();
        var sheetXml = XDocument.Load(sheetStream);
        var ssNs = XNamespace.Get(OpenXmlConst.Main2006SsNs);
        var legacyDrawingHF = sheetXml.Root!.Element(ssNs + "legacyDrawingHF");
        await Assert.That(legacyDrawingHF).IsNotNull().Because("Sheet should contain <legacyDrawingHF> element");
    }

    [Test]
    public async Task AddImage_Png_CanSaveAndReopen()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.PageSetup.Header.Left.AddImage(_pngPath, XLHFOccurrence.AllPages);
        ws.Cell("A1").Value = "Test";

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        // Verify we can reopen the file without errors
        ms.Position = 0;
        using var wb2 = new XLWorkbook(ms);
        var ws2 = wb2.Worksheets.First();
        await Assert.That(ws2.Cell("A1").Value.GetText()).IsEqualTo("Test");
    }

    [Test]
    public async Task AddImage_Jpeg_CanSaveAndReopen()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.PageSetup.Header.Center.AddImage(_jpegPath, XLHFOccurrence.AllPages);
        ws.Cell("A1").Value = "Test";

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        ms.Position = 0;
        using var wb2 = new XLWorkbook(ms);
        var ws2 = wb2.Worksheets.First();
        await Assert.That(ws2.Cell("A1").Value.GetText()).IsEqualTo("Test");
    }

    [Test]
    public async Task AddImage_NoImages_NoLegacyDrawingHFElement()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.PageSetup.Header.Left.AddText("Just text");
        ws.Cell("A1").Value = "Test";

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        ms.Position = 0;
        using var archive = new ZipArchive(ms, ZipArchiveMode.Read);
        var sheetEntry = archive.Entries.First(e => e.FullName.Contains("worksheets/sheet"));
        using var sheetStream = sheetEntry.Open();
        var sheetXml = XDocument.Load(sheetStream);
        var ssNs = XNamespace.Get(OpenXmlConst.Main2006SsNs);
        var legacyDrawingHF = sheetXml.Root!.Element(ssNs + "legacyDrawingHF");
        await Assert.That(legacyDrawingHF).IsNull().Because("Sheet should NOT contain <legacyDrawingHF> when no images");
    }

    [Test]
    public async Task AddImage_WithComments_BothVmlPartsCoexist()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").GetComment().AddText("A comment");
        ws.PageSetup.Header.Left.AddImage(_pngPath, XLHFOccurrence.AllPages);

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        ms.Position = 0;
        using var archive = new ZipArchive(ms, ZipArchiveMode.Read);

        // Should have at least 2 VML drawing parts (one for comments, one for HF images)
        var vmlParts = archive.Entries.Where(e =>
            e.FullName.Contains("drawings/") && e.FullName.EndsWith(".vml")).ToList();
        await Assert.That(vmlParts.Count).IsGreaterThanOrEqualTo(2).Because("Should have separate VML parts for comments and HF images");

        // Should have both legacyDrawing and legacyDrawingHF
        var sheetEntry = archive.Entries.First(e => e.FullName.Contains("worksheets/sheet"));
        using var sheetStream = sheetEntry.Open();
        var sheetXml = XDocument.Load(sheetStream);
        var ssNs = XNamespace.Get(OpenXmlConst.Main2006SsNs);

        await Assert.That(sheetXml.Root!.Element(ssNs + "legacyDrawing")).IsNotNull();
        await Assert.That(sheetXml.Root.Element(ssNs + "legacyDrawingHF")).IsNotNull();
    }

    #region Helpers

    private static async Task ExtractResource(string resourceName, string filePath)
    {
        using var stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName);
        await Assert.That(stream).IsNotNull().Because($"Resource {resourceName} not found");
        using var fs = File.Create(filePath);
        stream!.CopyTo(fs);
    }

    private static async Task AssertPackageContainsHFImage(
        Stream packageStream,
        string[] expectedShapeIds,
        string expectedHeaderText = "",
        string expectedFooterText = "")
    {
        var (vmlXml, headerText, footerText, _, _) = ExtractPackageInfo(packageStream);

        if (expectedHeaderText.Length > 0)
            await Assert.That(headerText).IsEqualTo(expectedHeaderText);

        if (expectedFooterText.Length > 0)
            await Assert.That(footerText).IsEqualTo(expectedFooterText);

        var shapeIds = GetVmlShapeIds(vmlXml);
        foreach (var expected in expectedShapeIds)
            await Assert.That(shapeIds).Contains(expected).Because($"VML should contain shape with id '{expected}'");
    }

    private static (string vmlXml, string headerText, string footerText, string[] contentTypes, string[] mediaFiles)
        ExtractPackageInfo(Stream packageStream)
    {
        packageStream.Position = 0;
        using var archive = new ZipArchive(packageStream, ZipArchiveMode.Read, leaveOpen: true);

        // Get sheet XML for header/footer text
        var sheetEntry = archive.Entries.First(e => e.FullName.Contains("worksheets/sheet"));
        string headerText, footerText;
        using (var sheetStream = sheetEntry.Open())
        {
            var sheetXml = XDocument.Load(sheetStream);
            var ssNs = XNamespace.Get(OpenXmlConst.Main2006SsNs);
            var hf = sheetXml.Root!.Element(ssNs + "headerFooter");
            headerText = hf?.Element(ssNs + "oddHeader")?.Value ?? "";
            footerText = hf?.Element(ssNs + "oddFooter")?.Value ?? "";
        }

        // Get VML content (look for the HF VML part, not the comments one)
        var vmlXml = "";
        var vmlEntries = archive.Entries.Where(e =>
            e.FullName.Contains("drawings/") && e.FullName.EndsWith(".vml")).ToList();
        foreach (var vmlEntry in vmlEntries)
        {
            using var vmlStream = vmlEntry.Open();
            using var reader = new StreamReader(vmlStream);
            var content = reader.ReadToEnd();
            // The HF VML part will have shape ids like LH, CH, RH, LF, CF, RF
            if (content.Contains("\"LH\"") || content.Contains("\"CH\"") || content.Contains("\"RH\"") ||
                content.Contains("\"LF\"") || content.Contains("\"CF\"") || content.Contains("\"RF\""))
            {
                vmlXml = content;
                break;
            }
        }

        // Get content types
        var ctEntry = archive.Entries.First(e => e.FullName == "[Content_Types].xml");
        string[] contentTypes;
        using (var ctStream = ctEntry.Open())
        {
            var ctXml = XDocument.Load(ctStream);
            contentTypes = ctXml.Root!.Elements()
                .Select(e => e.Attribute("ContentType")?.Value ?? "")
                .Where(v => !string.IsNullOrEmpty(v))
                .ToArray();
        }

        var mediaFiles = archive.Entries
            .Where(e => e.FullName.Contains("media/"))
            .Select(e => e.FullName)
            .ToArray();

        return (vmlXml, headerText, footerText, contentTypes, mediaFiles);
    }

    private static string[] GetVmlShapeIds(string vmlXml)
    {
        if (string.IsNullOrEmpty(vmlXml))
            return Array.Empty<string>();

        var doc = XDocument.Parse(vmlXml);
        var vNs = XNamespace.Get("urn:schemas-microsoft-com:vml");
        return doc.Descendants(vNs + "shape")
            .Select(s => s.Attribute("id")?.Value ?? "")
            .Where(id => !string.IsNullOrEmpty(id) && !id.StartsWith('_'))
            .ToArray();
    }

    private static (string id, double widthPt, double heightPt)[] GetVmlShapes(string vmlXml)
    {
        if (string.IsNullOrEmpty(vmlXml))
            return Array.Empty<(string, double, double)>();

        var doc = XDocument.Parse(vmlXml);
        var vNs = XNamespace.Get("urn:schemas-microsoft-com:vml");
        return doc.Descendants(vNs + "shape")
            .Select(s =>
            {
                var id = s.Attribute("id")?.Value ?? "";
                var style = s.Attribute("style")?.Value ?? "";
                var width = ParseStyleValue(style, "width");
                var height = ParseStyleValue(style, "height");
                return (id, width, height);
            })
            .Where(s => !string.IsNullOrEmpty(s.id) && !s.id.StartsWith('_'))
            .ToArray();
    }

    private static double ParseStyleValue(string style, string property)
    {
        var parts = style.Split(';');
        foreach (var part in parts)
        {
            var trimmed = part.Trim();
            if (trimmed.StartsWith(property + ":"))
            {
                var value = trimmed[(property.Length + 1)..].Trim();
                if (value.EndsWith("pt"))
                    return double.Parse(value[..^2], System.Globalization.CultureInfo.InvariantCulture);
            }
        }
        return 0;
    }

    #endregion
}
