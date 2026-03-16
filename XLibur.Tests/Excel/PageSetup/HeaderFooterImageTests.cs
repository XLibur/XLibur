using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using NUnit.Framework;
using XLibur.Excel;

namespace XLibur.Tests.Excel.PageSetup;

[TestFixture]
public class HeaderFooterImageTests
{
    private static readonly string[] ShapeIdsLH = ["LH"];
    private static readonly string[] ShapeIdsCH = ["CH"];
    private static readonly string[] ShapeIdsRH = ["RH"];
    private static readonly string[] ShapeIdsLF = ["LF"];

    private string _pngPath = null!;
    private string _jpegPath = null!;
    private string _tempDir = null!;

    [OneTimeSetUp]
    public void Setup()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "XLibur_HFImageTests_" + Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(_tempDir);

        _pngPath = Path.Combine(_tempDir, "test.png");
        _jpegPath = Path.Combine(_tempDir, "test.jpg");

        ExtractResource("XLibur.Tests.Resource.Images.SampleImagePng.png", _pngPath);
        ExtractResource("XLibur.Tests.Resource.Images.SampleImageExif.jpg", _jpegPath);
    }

    [OneTimeTearDown]
    public void Cleanup()
    {
        try { Directory.Delete(_tempDir, true); }
        catch { /* best effort cleanup */ }
    }

    [Test]
    public void AddImage_LeftHeader_EmitsGraphicMarkerAndVml()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.PageSetup.Header.Left.AddImage(_pngPath, XLHFOccurrence.AllPages);

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        ms.Position = 0;
        AssertPackageContainsHFImage(ms, expectedShapeIds: ShapeIdsLH, expectedHeaderText: "&L&G");
    }

    [Test]
    public void AddImage_CenterHeader_Works()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.PageSetup.Header.Center.AddImage(_jpegPath, XLHFOccurrence.AllPages);

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        ms.Position = 0;
        AssertPackageContainsHFImage(ms, expectedShapeIds: ShapeIdsCH, expectedHeaderText: "&C&G");
    }

    [Test]
    public void AddImage_RightHeader_Works()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.PageSetup.Header.Right.AddImage(_pngPath, XLHFOccurrence.AllPages);

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        ms.Position = 0;
        AssertPackageContainsHFImage(ms, expectedShapeIds: ShapeIdsRH, expectedHeaderText: "&R&G");
    }

    [Test]
    public void AddImage_LeftFooter_Works()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.PageSetup.Footer.Left.AddImage(_pngPath, XLHFOccurrence.AllPages);

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        ms.Position = 0;
        AssertPackageContainsHFImage(ms, expectedShapeIds: ShapeIdsLF, expectedFooterText: "&L&G");
    }

    [Test]
    public void AddImage_MixedTextAndImage_PreservesOrdering()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.PageSetup.Header.Left.AddText("Before:", XLHFOccurrence.AllPages);
        ws.PageSetup.Header.Left.AddImage(_pngPath, XLHFOccurrence.AllPages);
        ws.PageSetup.Header.Left.AddText(":After", XLHFOccurrence.AllPages);

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        ms.Position = 0;
        AssertPackageContainsHFImage(ms, expectedShapeIds: ShapeIdsLH, expectedHeaderText: "&LBefore:&G:After");
    }

    [Test]
    public void AddImage_MultipleSections_HeaderAndFooter()
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

        Assert.That(headerText, Is.EqualTo("&L&G&CTitle&R&G"));
        Assert.That(footerText, Is.EqualTo("&C&G"));

        // VML should have shapes LH, RH, CF
        var shapeIds = GetVmlShapeIds(vmlXml);
        Assert.That(shapeIds, Does.Contain("LH"));
        Assert.That(shapeIds, Does.Contain("RH"));
        Assert.That(shapeIds, Does.Contain("CF"));

        // Should have image media files
        Assert.That(mediaFiles.Count, Is.GreaterThanOrEqualTo(2));
    }

    [Test]
    public void AddImage_InvalidPath_ThrowsFileNotFound()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        Assert.Throws<FileNotFoundException>(() =>
            ws.PageSetup.Header.Left.AddImage(@"C:\nonexistent\image.png"));
    }

    [Test]
    public void AddImage_NullOrEmptyPath_ThrowsArgumentException()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        Assert.Throws<ArgumentException>(() =>
            ws.PageSetup.Header.Left.AddImage(""));

        Assert.Throws<ArgumentException>(() =>
            ws.PageSetup.Header.Left.AddImage("   "));
    }

    [Test]
    public void AddImage_UnsupportedFormat_ThrowsArgumentException()
    {
        var svgPath = Path.Combine(_tempDir, "test.svg");
        File.WriteAllText(svgPath, "<svg></svg>");

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        Assert.Throws<ArgumentException>(() =>
            ws.PageSetup.Header.Left.AddImage(svgPath));
    }

    [Test]
    public void AddImage_ScalingPreservesAspectRatio_DoesNotUpscale()
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
        Assert.That(shape.widthPt, Is.LessThanOrEqualTo(180.1));
        // The height should be less than or equal to 0.6" = 43.2pt
        Assert.That(shape.heightPt, Is.LessThanOrEqualTo(43.3));
        // Both should be > 0
        Assert.That(shape.widthPt, Is.GreaterThan(0));
        Assert.That(shape.heightPt, Is.GreaterThan(0));
    }

    [Test]
    public void AddImage_PackageHasCorrectContentTypes()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.PageSetup.Header.Left.AddImage(_pngPath, XLHFOccurrence.AllPages);

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        ms.Position = 0;
        var (_, _, _, contentTypes, _) = ExtractPackageInfo(ms);

        // Should have vml and png content types
        Assert.That(contentTypes, Does.Contain("application/vnd.openxmlformats-officedocument.vmlDrawing"));
        Assert.That(contentTypes, Does.Contain("image/png"));
    }

    [Test]
    public void AddImage_VmlHasImageRelationships()
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
        Assert.That(vmlRelsEntry, Is.Not.Null, "VML rels file should exist");

        using var relsStream = vmlRelsEntry!.Open();
        var relsXml = XDocument.Load(relsStream);
        var ns = XNamespace.Get("http://schemas.openxmlformats.org/package/2006/relationships");
        var imageRels = relsXml.Root!.Elements(ns + "Relationship")
            .Where(r => r.Attribute("Type")?.Value?.Contains("image") == true)
            .ToList();

        Assert.That(imageRels, Has.Count.GreaterThanOrEqualTo(1), "VML should have image relationships");
    }

    [Test]
    public void AddImage_SheetHasLegacyDrawingHFElement()
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
        var ssNs = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        var legacyDrawingHF = sheetXml.Root!.Element(ssNs + "legacyDrawingHF");
        Assert.That(legacyDrawingHF, Is.Not.Null, "Sheet should contain <legacyDrawingHF> element");
    }

    [Test]
    public void AddImage_Png_CanSaveAndReopen()
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
        Assert.That(ws2.Cell("A1").Value.GetText(), Is.EqualTo("Test"));
    }

    [Test]
    public void AddImage_Jpeg_CanSaveAndReopen()
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
        Assert.That(ws2.Cell("A1").Value.GetText(), Is.EqualTo("Test"));
    }

    [Test]
    public void AddImage_NoImages_NoLegacyDrawingHFElement()
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
        var ssNs = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        var legacyDrawingHF = sheetXml.Root!.Element(ssNs + "legacyDrawingHF");
        Assert.That(legacyDrawingHF, Is.Null, "Sheet should NOT contain <legacyDrawingHF> when no images");
    }

    [Test]
    public void AddImage_WithComments_BothVmlPartsCoexist()
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
        Assert.That(vmlParts.Count, Is.GreaterThanOrEqualTo(2),
            "Should have separate VML parts for comments and HF images");

        // Should have both legacyDrawing and legacyDrawingHF
        var sheetEntry = archive.Entries.First(e => e.FullName.Contains("worksheets/sheet"));
        using var sheetStream = sheetEntry.Open();
        var sheetXml = XDocument.Load(sheetStream);
        var ssNs = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");

        Assert.That(sheetXml.Root!.Element(ssNs + "legacyDrawing"), Is.Not.Null);
        Assert.That(sheetXml.Root!.Element(ssNs + "legacyDrawingHF"), Is.Not.Null);
    }

    #region Helpers

    private static void ExtractResource(string resourceName, string filePath)
    {
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName);
        Assert.That(stream, Is.Not.Null, $"Resource {resourceName} not found");
        using var fs = File.Create(filePath);
        stream!.CopyTo(fs);
    }

    private static void AssertPackageContainsHFImage(
        Stream packageStream,
        string[] expectedShapeIds,
        string expectedHeaderText = "",
        string expectedFooterText = "")
    {
        var (vmlXml, headerText, footerText, _, _) = ExtractPackageInfo(packageStream);

        if (expectedHeaderText.Length > 0)
            Assert.That(headerText, Is.EqualTo(expectedHeaderText));

        if (expectedFooterText.Length > 0)
            Assert.That(footerText, Is.EqualTo(expectedFooterText));

        var shapeIds = GetVmlShapeIds(vmlXml);
        foreach (var expected in expectedShapeIds)
            Assert.That(shapeIds, Does.Contain(expected), $"VML should contain shape with id '{expected}'");
    }

    private static (string vmlXml, string headerText, string footerText, string[] contentTypes, string[] mediaFiles)
        ExtractPackageInfo(Stream packageStream)
    {
        packageStream.Position = 0;
        using var archive = new ZipArchive(packageStream, ZipArchiveMode.Read, leaveOpen: true);

        // Get sheet XML for header/footer text
        var sheetEntry = archive.Entries.First(e => e.FullName.Contains("worksheets/sheet"));
        string headerText = "", footerText = "";
        using (var sheetStream = sheetEntry.Open())
        {
            var sheetXml = XDocument.Load(sheetStream);
            var ssNs = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
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
            .Where(s => !string.IsNullOrEmpty(s.id) && !s.id.StartsWith("_"))
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
