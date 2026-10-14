using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using XLibur.Excel.Drawings;

namespace XLibur.Tests.Excel.Cells;

/// <summary>
/// The rich data parts of an in-cell image are written as Excel writes them (#739).
/// </summary>
/// <remarks>
/// Excel refused to open any workbook with an in-cell image that XLibur wrote. Bisecting the
/// package against one Excel wrote found two faults, each of them fatal on its own: the rich value
/// structure part was in the <c>richdata2</c> namespace instead of <c>richdata</c>, and named its
/// image key <c>_rvRel:LocalImageFileUri</c> instead of <c>_rvRel:LocalImageIdentifier</c>. With
/// both fixed, Excel opened the file but showed <c>#FIELD!</c> in the cell, because the
/// relationship to the rich value relationship part had the 2017 type, which Excel does not read.
/// <c>ExcelCellImage.xlsx</c> was made by Excel 365 (<c>Range.InsertPictureInCell</c>, image in
/// B2, alt text "Sample image").
/// </remarks>
public class CellImageExcelFormatTests
{
    private const string RichDataNs = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";
    private const string RichValueRelRelType = "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel";
    private const string RichValueRelContentType = "application/vnd.ms-excel.richvaluerel+xml";
    private const string RichValueStructureRelType = "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure";
    private const string RichValueRelType = "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue";
    private const string ExcelFile = @"Other\Cells\ExcelCellImage.xlsx";

    [Test]
    public async Task Rich_value_rel_part_has_the_relationship_and_content_type_Excel_writes()
    {
        using var saved = SaveWithImage();
        using var document = SpreadsheetDocument.Open(saved, false);

        var parts = RichValueRelParts(document);

        await Assert.That(parts.Count).IsEqualTo(1);
        await Assert.That(parts[0].ContentType).IsEqualTo(RichValueRelContentType);
    }

    [Test]
    public async Task Rich_value_structure_is_in_the_richdata_namespace_with_the_keys_Excel_writes()
    {
        using var saved = SaveWithImage();
        using var document = SpreadsheetDocument.Open(saved, false);

        var root = ReadStructure(document);

        await Assert.That(root.Name).IsEqualTo(XName.Get("rvStructures", RichDataNs));
        var structure = root.Elements(XName.Get("s", RichDataNs)).Single();
        await Assert.That(structure.Attribute("t")?.Value).IsEqualTo("_localImage");
        var keys = structure.Elements(XName.Get("k", RichDataNs)).Select(k => k.Attribute("n")?.Value).ToArray();
        await Assert.That(string.Join(",", keys)).IsEqualTo("_rvRel:LocalImageIdentifier,CalcOrigin,Text");
    }

    [Test]
    public async Task Save_and_reload_keeps_the_image_and_its_alt_text()
    {
        using var saved = SaveWithImage();
        using var wb = new XLWorkbook(saved);

        var image = ((XLCell)wb.Worksheet(1).Cell("B2")).CellImage;

        await Assert.That(image).IsNotNull();
        await Assert.That(image!.AltText).IsEqualTo("red pixel");
    }

    [Test]
    public async Task Load_reads_the_image_and_alt_text_Excel_wrote()
    {
        using var wb = TestHelper.LoadWorkbook(ExcelFile);

        var image = ((XLCell)wb.Worksheet(1).Cell("B2")).CellImage;

        await Assert.That(image).IsNotNull();
        await Assert.That(image!.AltText).IsEqualTo("Sample image");
        await Assert.That(wb.InCellImages.Count).IsEqualTo(1);
    }

    [Test]
    public async Task Excel_file_round_trip_keeps_the_image_in_the_parts_Excel_reads()
    {
        using var source = TestHelper.LoadWorkbook(ExcelFile);
        source.Worksheet(1).Cell("D4").Value = "edited";
        using var saved = new MemoryStream();
        source.SaveAs(saved);

        saved.Position = 0;
        using (var document = SpreadsheetDocument.Open(saved, false))
        {
            await Assert.That(RichValueRelParts(document).Count).IsEqualTo(1);
            await Assert.That(ReadStructure(document).Name.NamespaceName).IsEqualTo(RichDataNs);
            await Assert.That(MetadataChildren(document)).IsEqualTo("metadataTypes,futureMetadata,valueMetadata");
        }

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var image = ((XLCell)reloaded.Worksheet(1).Cell("B2")).CellImage;
        await Assert.That(image).IsNotNull();
        await Assert.That(image!.AltText).IsEqualTo("Sample image");
    }

    [Test]
    public async Task Metadata_part_keeps_schema_order_with_dynamic_arrays_and_images()
    {
        // The dynamic array writer adds futureMetadata and cellMetadata first; the image writer
        // then appended its futureMetadata after the cellMetadata.
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(2)");
            using var png = new MemoryStream(XLCellImageTests.CreateTestPng());
            ws.Cell("B2").SetCellImage(png, XLPictureFormat.Png, "red pixel");
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var document = SpreadsheetDocument.Open(ms, false);

        await Assert.That(MetadataChildren(document))
            .IsEqualTo("metadataTypes,futureMetadata,futureMetadata,cellMetadata,valueMetadata");
    }

    [Test]
    public async Task Loaded_image_workbook_saved_again_keeps_metadata_in_schema_order()
    {
        using var first = SaveWithImage();
        using var wb = new XLWorkbook(first);
        wb.Worksheet(1).Cell("D4").Value = 7;
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        saved.Position = 0;
        using var document = SpreadsheetDocument.Open(saved, false);

        await Assert.That(MetadataChildren(document)).IsEqualTo("metadataTypes,futureMetadata,valueMetadata");
    }

    [Test]
    public async Task Removing_the_last_image_of_a_loaded_workbook_drops_the_cells_value_metadata_index()
    {
        // Cell C3 had vm="2". With its image removed it kept vm="2" while the save wrote one
        // value metadata record, and Excel refused the file.
        using var first = SaveWithTwoImages();
        using var wb = new XLWorkbook(first);
        wb.Worksheet(1).Cell("C3").RemoveCellImage();
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        saved.Position = 0;
        using (var document = SpreadsheetDocument.Open(saved, false))
        {
            var cells = SheetCells(document);
            await Assert.That(cells.TryGetValue("C3", out var c3) ? c3.Attribute("vm") : null).IsNull();
            await Assert.That(cells["B2"].Attribute("vm")?.Value).IsEqualTo("1");
        }

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var ws = reloaded.Worksheet(1);
        await Assert.That(ws.Cell("C3").HasCellImage).IsFalse();
        await Assert.That(ws.Cell("C3").Value.IsBlank).IsTrue();
        await Assert.That(((XLCell)ws.Cell("B2")).CellImage?.AltText).IsEqualTo("first");
    }

    [Test]
    public async Task Removing_the_first_image_of_a_loaded_workbook_does_not_give_its_cell_the_next_image()
    {
        // Cell B2 had vm="1". With its image removed it kept vm="1", which the save gave to the
        // image left in C3, so B2 showed C3's image.
        using var first = SaveWithTwoImages();
        using var wb = new XLWorkbook(first);
        wb.Worksheet(1).Cell("B2").RemoveCellImage();
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var ws = reloaded.Worksheet(1);
        await Assert.That(ws.Cell("B2").HasCellImage).IsFalse();
        await Assert.That(ws.Cell("B2").Value.IsBlank).IsTrue();
        await Assert.That(((XLCell)ws.Cell("C3")).CellImage?.AltText).IsEqualTo("second");
    }

    [Test]
    public async Task Load_reads_a_file_XLibur_wrote_before_the_fix_and_saves_it_in_Excels_form()
    {
        // XLiburLegacyCellImage.xlsx was saved by XLibur before #739: richValueRel related by
        // the 2017 type, the structure in richdata2 with key _rvRel:LocalImageFileUri.
        using var wb = TestHelper.LoadWorkbook(@"Other\Cells\XLiburLegacyCellImage.xlsx");
        var image = ((XLCell)wb.Worksheet("S").Cell("B2")).CellImage;
        await Assert.That(image).IsNotNull();
        await Assert.That(image!.AltText).IsEqualTo("alt text one");

        using var saved = new MemoryStream();
        wb.SaveAs(saved);
        saved.Position = 0;
        using var document = SpreadsheetDocument.Open(saved, false);

        await Assert.That(RichValueRelParts(document).Count).IsEqualTo(1);
        var legacy = document.WorkbookPart!.Parts.Count(p =>
            p.OpenXmlPart.RelationshipType == "http://schemas.microsoft.com/office/2017/06/relationships/richValueRel");
        await Assert.That(legacy).IsEqualTo(0);
        await Assert.That(ReadStructure(document).Name.NamespaceName).IsEqualTo(RichDataNs);
    }

    [Test]
    public async Task Load_reads_images_of_every_local_image_structure_by_key_name()
    {
        // Excel writes one structure per set of keys: here B2's image has no alt text and C3's
        // has, so their rich values use different structures.
        using var source = SaveWithTwoImages();
        RewriteRichData(source,
            """<s t="_localImage"><k n="_rvRel:LocalImageIdentifier" t="i"/><k n="CalcOrigin" t="i"/></s>""" +
            """<s t="_localImage"><k n="CalcOrigin" t="i"/><k n="Text" t="s"/><k n="_rvRel:LocalImageIdentifier" t="i"/></s>""",
            """<rv s="0"><v>0</v><v>5</v></rv><rv s="1"><v>5</v><v>second</v><v>1</v></rv>""");

        using var wb = new XLWorkbook(source);
        var ws = wb.Worksheet(1);

        var b2 = ((XLCell)ws.Cell("B2")).CellImage;
        var c3 = ((XLCell)ws.Cell("C3")).CellImage;
        await Assert.That(b2).IsNotNull();
        await Assert.That(b2!.AltText).IsEqualTo(string.Empty);
        await Assert.That(c3).IsNotNull();
        await Assert.That(c3!.AltText).IsEqualTo("second");
        await Assert.That(c3.WorkbookImageIndex).IsNotEqualTo(b2.WorkbookImageIndex);
    }

    [Test]
    public async Task A_rich_value_that_is_not_a_local_image_leaves_the_rich_data_as_the_file_had_it()
    {
        // The save writes the rich data parts from the images it knows. Had it taken B2's image,
        // it would have dropped C3's rich value, and C3's vm would have named nothing.
        using var source = SaveWithTwoImages();
        const string RichValues = """<rv s="0"><v>0</v><v>5</v><v>first</v></rv><rv s="1"><v>other</v></rv>""";
        RewriteRichData(source,
            """<s t="_localImage"><k n="_rvRel:LocalImageIdentifier" t="i"/><k n="CalcOrigin" t="i"/><k n="Text" t="s"/></s>""" +
            """<s t="_someOtherType"><k n="Name" t="s"/></s>""",
            RichValues);

        using var wb = new XLWorkbook(source);
        await Assert.That(wb.Worksheet(1).Cell("B2").HasCellImage).IsFalse();
        await Assert.That(wb.InCellImages.Count).IsEqualTo(0);
        wb.Worksheet(1).Cell("D4").Value = 7;
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        saved.Position = 0;
        using var document = SpreadsheetDocument.Open(saved, false);
        var cells = SheetCells(document);
        await Assert.That(cells["B2"].Attribute("vm")?.Value).IsEqualTo("1");
        await Assert.That(cells["C3"].Attribute("vm")?.Value).IsEqualTo("2");
        await Assert.That(ReadRichValues(document)).Contains(RichValues);
    }

    /// <summary>
    /// Replaces the structures and rich values of a package written by <see cref="SaveWithTwoImages"/>.
    /// </summary>
    private static void RewriteRichData(MemoryStream package, string structures, string richValues)
    {
        package.Position = 0;
        using (var document = SpreadsheetDocument.Open(package, true))
        {
            var structureCount = structures.Split("<s ").Length - 1;
            var richValueCount = richValues.Split("<rv ").Length - 1;
            Feed(PartByRelType(document, RichValueStructureRelType),
                $"""<rvStructures xmlns="{RichDataNs}" count="{structureCount}">{structures}</rvStructures>""");
            Feed(PartByRelType(document, RichValueRelType),
                $"""<rvData xmlns="{RichDataNs}" count="{richValueCount}">{richValues}</rvData>""");
        }

        package.Position = 0;

        static void Feed(OpenXmlPart part, string xml)
        {
            using var content = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xml));
            part.FeedData(content);
        }
    }

    private static string ReadRichValues(SpreadsheetDocument document)
    {
        using var stream = PartByRelType(document, RichValueRelType).GetStream(FileMode.Open, FileAccess.Read);
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    private static OpenXmlPart PartByRelType(SpreadsheetDocument document, string relType) =>
        document.WorkbookPart!.Parts.Single(p => p.OpenXmlPart.RelationshipType == relType).OpenXmlPart;

    private static MemoryStream SaveWithTwoImages()
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            using (var png = new MemoryStream(XLCellImageTests.CreateTestPng()))
                ws.Cell("B2").SetCellImage(png, XLPictureFormat.Png, "first");

            var other = XLCellImageTests.CreateTestPng().Concat(new byte[] { 0 }).ToArray();
            using (var png = new MemoryStream(other))
                ws.Cell("C3").SetCellImage(png, XLPictureFormat.Png, "second");

            wb.SaveAs(ms);
        }

        ms.Position = 0;
        return ms;
    }

    private static Dictionary<string, XElement> SheetCells(SpreadsheetDocument document)
    {
        var sheetPart = document.WorkbookPart!.WorksheetParts.Single();
        using var stream = sheetPart.GetStream(FileMode.Open, FileAccess.Read);
        return XDocument.Load(stream).Descendants()
            .Where(e => e.Name.LocalName == "c")
            .ToDictionary(e => e.Attribute("r")!.Value);
    }

    private static string MetadataChildren(SpreadsheetDocument document)
    {
        using var stream = document.WorkbookPart!.CellMetadataPart!.GetStream(FileMode.Open, FileAccess.Read);
        return string.Join(",", XDocument.Load(stream).Root!.Elements().Select(e => e.Name.LocalName));
    }

    private static MemoryStream SaveWithImage()
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            using var png = new MemoryStream(XLCellImageTests.CreateTestPng());
            wb.AddWorksheet().Cell("B2").SetCellImage(png, XLPictureFormat.Png, "red pixel");
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        return ms;
    }

    private static List<OpenXmlPart> RichValueRelParts(SpreadsheetDocument document)
    {
        return document.WorkbookPart!.Parts
            .Where(p => p.OpenXmlPart.RelationshipType == RichValueRelRelType)
            .Select(p => p.OpenXmlPart)
            .ToList();
    }

    private static XElement ReadStructure(SpreadsheetDocument document)
    {
        var part = document.WorkbookPart!.Parts
            .Single(p => p.OpenXmlPart.RelationshipType == RichValueStructureRelType)
            .OpenXmlPart;
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        return XDocument.Load(stream).Root!;
    }
}
