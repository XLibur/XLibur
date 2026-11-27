using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using XLibur.Excel.Drawings;
using XLibur.Excel.IO;
using XLibur.Tests.Excel.Cells;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// The error values newer than Excel 2007, such as <c>#SPILL!</c> and <c>#CALC!</c>, are not error values
/// of the file format. Excel writes such a cell as <c>#VALUE!</c>, with a vm index whose rich value is an
/// <c>_error</c> that names the error by its <c>errorType</c>, and refuses to open a file whose cell holds
/// <c>#SPILL!</c> or <c>#CALC!</c> itself. Checked over COM on 2026-11-01 with Excel 365, which also
/// opens an <c>_error</c> structure with no key but <c>errorType</c> and shows the error it names.
/// </summary>
public class ErrorValueSavingTests
{
    [Test]
    public async Task A_newer_error_is_saved_as_VALUE_with_a_rich_value_that_names_it()
    {
        using var output = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = XLError.SpillRange;
            ws.Cell("A2").Value = XLError.Calc;
            ws.Cell("A3").Value = XLError.NoValueAvailable;
            ws.Cell("A4").Value = XLError.Calc;
            wb.SaveAs(output);
        }

        var xml = output.Sheet1Xml();
        var richValues = RichValuePartXml(output);

        await Assert.That(xml).DoesNotContain("#SPILL!");
        await Assert.That(xml).DoesNotContain("#CALC!");
        await Assert.That(xml).Contains("<x:c r=\"A1\" t=\"e\" vm=\"1\"><x:v>#VALUE!</x:v></x:c>");
        await Assert.That(xml).Contains("<x:c r=\"A2\" t=\"e\" vm=\"2\"><x:v>#VALUE!</x:v></x:c>");
        await Assert.That(xml).Contains("<x:c r=\"A3\" t=\"e\"><x:v>#N/A</x:v></x:c>");
        await Assert.That(xml).Contains("<x:c r=\"A4\" t=\"e\" vm=\"2\"><x:v>#VALUE!</x:v></x:c>");
        await Assert.That(richValues).Contains("<v>8</v>");
        await Assert.That(richValues).Contains("<v>13</v>");
    }

    private static string RichValuePartXml(Stream package)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        var part = document.WorkbookPart!.Parts
            .Single(p => p.OpenXmlPart.RelationshipType == RichDataWriter.RichValueRelType).OpenXmlPart;
        using var reader = new StreamReader(part.GetStream());
        return reader.ReadToEnd();
    }

    [Test]
    public async Task The_result_of_a_formula_is_saved_so_and_loads_as_the_error_it_was()
    {
        using var output = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            ws.Cell("B1").FormulaA1 = "LAMBDA(x,x*2)";
            ws.Cell("C1").Value = 1;
            ws.Cell("C3").SetDynamicFormulaA1("SEQUENCE(3)");
            ws.Cell("C4").Value = "blocks";
            wb.RecalculateAllFormulas();
            wb.SaveAs(output);
        }

        using var reloaded = new XLWorkbook(output);
        var sheet = reloaded.Worksheet(1);

        await Assert.That(sheet.Cell("B1").CachedValue).IsEqualTo(XLError.Calc);
        await Assert.That(sheet.Cell("C3").CachedValue).IsEqualTo(XLError.SpillRange);
    }

    [Test]
    public async Task A_newer_error_and_an_in_cell_image_are_saved_together()
    {
        using var output = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = XLError.Calc;
            using var png = new MemoryStream(XLCellImageTests.CreateTestPng());
            ws.Cell("B1").SetCellImage(png, XLPictureFormat.Png, "red pixel");
            wb.SaveAs(output);
        }

        using var reloaded = new XLWorkbook(output);
        var sheet = reloaded.Worksheet(1);

        await Assert.That(sheet.Cell("A1").Value).IsEqualTo(XLError.Calc);
        await Assert.That(((XLCell)sheet.Cell("B1")!).CellImage).IsNotNull();
    }

    /// <summary>
    /// A cube function's cell names an <c>XLMDX</c> record of the value metadata, which the save does not
    /// model. A rich value's record before it must not move it: the save rewrites the rich values'
    /// records where they are, and adds new ones after the others.
    /// </summary>
    [Test]
    public async Task A_cube_cell_keeps_its_value_metadata_when_a_rich_value_comes_before_it()
    {
        using var source = new MemoryStream();
        using (var wb = TestHelper.LoadWorkbook(@"Other\Cubes\CubeFromRange-Input.xlsx"))
        {
            wb.Worksheet("Data").Cell("Z1").Value = XLError.Calc;
            wb.SaveAs(source);
        }

        // The cube's 14 records come first, then the error's. Put the error's first, as a file may.
        var metadata = XDocument.Parse(source.ReadPart(MetadataPart));
        var valueMetadata = metadata.Root!.Elements().Single(e => e.Name.LocalName == "valueMetadata");
        var errorRecord = valueMetadata.Elements().Last();
        errorRecord.Remove();
        valueMetadata.AddFirst(errorRecord);
        source.RewritePart(MetadataPart, _ => metadata.ToString(System.Xml.Linq.SaveOptions.DisableFormatting));
        source.RewritePart("xl/worksheets/sheet1.xml", xml => xml.Replace("vm=\"15\"", "vm=\"1\""));
        source.RewritePart("xl/worksheets/sheet2.xml",
            xml => Regex.Replace(xml, "vm=\"(\\d+)\"", m => $"vm=\"{int.Parse(m.Groups[1].Value) + 1}\""));
        var cubeIndexes = CubeCellIndexes(source.ReadPart("xl/worksheets/sheet2.xml"));

        using var output = new MemoryStream();
        using (var wb = new XLWorkbook(source))
        {
            await Assert.That(wb.Worksheet("Data").Cell("Z1").Value).IsEqualTo(XLError.Calc);
            wb.Worksheet("Data").Cell("Z2").Value = XLError.SpillRange;
            wb.SaveAs(output);
        }

        var saved = XDocument.Parse(output.ReadPart(MetadataPart));
        await Assert.That(CubeCellIndexes(output.ReadPart("xl/worksheets/sheet2.xml"))).IsEquivalentTo(cubeIndexes);
        await Assert.That(cubeIndexes.Select(i => TypeOfValueMetadata(saved, i)).Distinct()).IsEquivalentTo(["XLMDX"]);

        using var reloaded = new XLWorkbook(output);
        await Assert.That(reloaded.Worksheet("Data").Cell("Z1").Value).IsEqualTo(XLError.Calc);
        await Assert.That(reloaded.Worksheet("Data").Cell("Z2").Value).IsEqualTo(XLError.SpillRange);
    }

    /// <summary>
    /// A workbook whose rich data XLibur does not know keeps it as loaded, and each cell its vm index.
    /// A new error's rich value is added after the ones the file has, which keep their indexes.
    /// </summary>
    [Test]
    public async Task A_newer_error_is_added_to_rich_data_the_workbook_keeps_as_loaded()
    {
        using var source = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            using var png = new MemoryStream(XLCellImageTests.CreateTestPng());
            wb.AddWorksheet().Cell("B1").SetCellImage(png, XLPictureFormat.Png);
            wb.SaveAs(source);
        }

        // A structure XLibur does not read: the rich data is kept as the file has it.
        var structurePart = RichDataPartName(source, RichDataWriter.RichValueStructureRelType);
        source.RewritePart(structurePart, xml => xml.Replace("_localImage", "_webImage"));

        using var output = new MemoryStream();
        using (var wb = new XLWorkbook(source))
        {
            wb.Worksheet(1).Cell("A1").Value = XLError.Calc;
            wb.SaveAs(output);
        }

        var xml = output.Sheet1Xml();
        await Assert.That(xml).Contains("<x:c r=\"B1\" vm=\"1\">");
        await Assert.That(output.ReadPart(RichDataPartName(output, RichDataWriter.RichValueStructureRelType)))
            .Contains("_webImage");

        using var reloaded = new XLWorkbook(output);
        await Assert.That(reloaded.Worksheet(1).Cell("A1").Value).IsEqualTo(XLError.Calc);
    }

    private const string LambdaFormulas = @"Other\ExcelAuthored\LambdaFormulas.xlsx";

    /// <summary>The cells of <see cref="LambdaFormulas"/> whose error has a rich value, which Excel 365 wrote.</summary>
    private static readonly string[] LambdaErrorCells = ["D1", "D29", "D33", "D65", "D145"];

    /// <summary>
    /// Excel 365 writes an error of a LAMBDA with a rich value that says more than its type: a
    /// <c>subType</c>, and for some an <c>argument</c>. The save kept a rich value only for an error newer
    /// than Excel 2007, and wrote it with its <c>errorType</c> alone, so a <c>#VALUE!</c> lost its record
    /// and its vm index, and a <c>#CALC!</c> its <c>subType</c> (#954).
    /// </summary>
    [Test]
    [Arguments(false)]
    [Arguments(true)]
    public async Task A_loaded_errors_rich_value_is_saved_as_the_file_had_it(bool edited)
    {
        using var original = TestHelper.OpenResource(LambdaFormulas);
        using var output = new MemoryStream();
        using (var wb = new XLWorkbook(original))
        {
            if (edited)
                wb.Worksheet(1).Cell("Z1").Value = 1;
            wb.SaveAs(output);
        }

        foreach (var cell in LambdaErrorCells)
            await Assert.That(ErrorRichValueOf(output, cell)).IsEqualTo(ErrorRichValueOf(original, cell)).Because(cell);
    }

    [Test]
    public async Task A_second_save_keeps_a_loaded_errors_rich_value()
    {
        using var original = TestHelper.OpenResource(LambdaFormulas);
        using var wb = new XLWorkbook(original);
        using var first = new MemoryStream();
        wb.SaveAs(first);
        wb.Worksheet(1).Cell("Z1").Value = 1;
        using var second = new MemoryStream();
        wb.SaveAs(second);

        foreach (var cell in LambdaErrorCells)
            await Assert.That(ErrorRichValueOf(second, cell)).IsEqualTo(ErrorRichValueOf(original, cell)).Because(cell);
    }

    [Test]
    public async Task A_cell_that_no_longer_holds_its_loaded_error_does_not_keep_its_rich_value()
    {
        using var original = TestHelper.OpenResource(LambdaFormulas);
        using var output = new MemoryStream();
        using (var wb = new XLWorkbook(original))
        {
            var ws = wb.Worksheet(1);
            ws.Cell("D29").Value = XLError.DivisionByZero;
            ws.Cell("D33").Value = XLError.Calc;
            wb.SaveAs(output);
        }

        await Assert.That(ErrorRichValueOf(output, "D29")).IsNull();
        await Assert.That(ErrorRichValueOf(output, "D33")).IsEqualTo("_error errorType=13");
        await Assert.That(ErrorRichValueOf(output, "D65")).IsEqualTo(ErrorRichValueOf(original, "D65"));
    }

    /// <summary>
    /// The rich value the vm index of <paramref name="cellReference"/> on the first sheet names, as its
    /// structure's type and each key with its value, or null when the cell has no rich value.
    /// </summary>
    private static string? ErrorRichValueOf(Stream package, string cellReference)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        var workbookPart = document.WorkbookPart!;
        var cell = Load(workbookPart.WorksheetParts.First(p => workbookPart.GetIdOfPart(p) ==
                       workbookPart.Workbook!.Sheets!.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().First().Id!.Value))
            .Descendants().Single(e => e.Name.LocalName == "c" && (string?)e.Attribute("r") == cellReference);
        if ((string?)cell.Attribute("vm") is not { } vm)
            return null;

        var metadata = Load(workbookPart.CellMetadataPart!);
        var types = Children(metadata, "metadataTypes").Select(t => (string?)t.Attribute("name")).ToList();
        var record = Children(metadata, "valueMetadata").ElementAt(int.Parse(vm) - 1).Elements().Single();
        if (types[(int)record.Attribute("t")! - 1] != "XLRICHVALUE")
            return null;

        var block = metadata.Elements().Single(e => e.Name.LocalName == "futureMetadata" && (string?)e.Attribute("name") == "XLRICHVALUE")
            .Elements().ElementAt((int)record.Attribute("v")!);
        var richValueIndex = (int)block.Descendants().Single(e => e.Name.LocalName == "rvb").Attribute("i")!;
        var richValue = Load(PartByRelType(workbookPart, RichDataWriter.RichValueRelType)).Elements().ElementAt(richValueIndex);
        var structure = Load(PartByRelType(workbookPart, RichDataWriter.RichValueStructureRelType)).Elements()
            .ElementAt((int)richValue.Attribute("s")!);
        var keys = structure.Elements().Select(k => (string)k.Attribute("n")!);
        var values = richValue.Elements().Select(v => v.Value);
        return string.Join(" ", [(string)structure.Attribute("t")!, .. keys.Zip(values, (k, v) => $"{k}={v}")]);

        static XElement Load(OpenXmlPart part)
        {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            return XDocument.Load(stream).Root!;
        }

        static IEnumerable<XElement> Children(XElement metadata, string name) =>
            metadata.Elements().Single(e => e.Name.LocalName == name).Elements();

        static OpenXmlPart PartByRelType(WorkbookPart workbookPart, string relationshipType) =>
            workbookPart.Parts.Single(p => p.OpenXmlPart.RelationshipType == relationshipType).OpenXmlPart;
    }

    private static string RichDataPartName(Stream package, string relationshipType)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        return document.WorkbookPart!.Parts.Single(p => p.OpenXmlPart.RelationshipType == relationshipType)
            .OpenXmlPart.Uri.OriginalString.TrimStart('/');
    }

    private const string MetadataPart = "xl/metadata.xml";

    private static int[] CubeCellIndexes(string sheetXml) =>
        Regex.Matches(sheetXml, "vm=\"(\\d+)\"").Select(m => int.Parse(m.Groups[1].Value)).Order().ToArray();

    /// <summary>The name of the metadata type of the value metadata record at the 1-based <paramref name="index"/>.</summary>
    private static string TypeOfValueMetadata(XDocument metadata, int index)
    {
        var root = metadata.Root!;
        var types = root.Elements().Single(e => e.Name.LocalName == "metadataTypes").Elements().ToList();
        var record = root.Elements().Single(e => e.Name.LocalName == "valueMetadata").Elements().ElementAt(index - 1)
            .Elements().Single();
        return (string)types[(int)record.Attribute("t")! - 1].Attribute("name")!;
    }

    /// <summary>
    /// A second save copies a sheet it does not change as the first save wrote it (#719), with the vm
    /// indexes the first save gave its cells, so those must name the same rich values again.
    /// </summary>
    [Test]
    public async Task A_sheet_a_second_save_leaves_as_it_was_keeps_its_errors()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Errors").Cell("A1").Value = XLError.Calc;
        var other = wb.AddWorksheet("Other");
        using var first = new MemoryStream();
        wb.SaveAs(first);

        // The new image takes the first rich value, ahead of the error's.
        using var png = new MemoryStream(XLCellImageTests.CreateTestPng());
        other.Cell("A1").SetCellImage(png, XLPictureFormat.Png);
        using var second = new MemoryStream();
        wb.SaveAs(second);

        using var reloaded = new XLWorkbook(second);

        await Assert.That(reloaded.Worksheet("Errors").Cell("A1").Value).IsEqualTo(XLError.Calc);
        await Assert.That(((XLCell)reloaded.Worksheet("Other").Cell("A1")!).CellImage).IsNotNull();
    }
}
