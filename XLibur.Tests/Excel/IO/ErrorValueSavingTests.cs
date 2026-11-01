using System.IO;
using System.Linq;
using System.Threading.Tasks;
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
