using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// A blank cell is written when its style differs from the one a missing cell would get on load,
/// which is its column's when the row has no style of its own (#906).
/// </summary>
public class BlankStyledCellWriteTests
{
    [Test]
    public async Task A_blank_cell_with_the_sheet_style_in_a_styled_column_keeps_its_style()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Column("B").Style.NumberFormat.Format = "yyyy/mm/dd";
        ws.Cell("B2").Style.NumberFormat.NumberFormatId = 0;

        using var reloaded = SaveAndReload(wb);

        var reloadedSheet = reloaded.Worksheet("Sheet1");
        await Assert.That(reloadedSheet.Cell("B2").Style.NumberFormat.Format).IsEqualTo("");
        await Assert.That(reloadedSheet.Cell("B2").Style.NumberFormat.NumberFormatId).IsEqualTo(0);
        await Assert.That(reloadedSheet.Cell("B3").Style.NumberFormat.Format).IsEqualTo("yyyy/mm/dd");
    }

    [Test]
    public async Task A_blank_cell_with_the_sheet_style_in_a_styled_row_keeps_its_style()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Row(2).Style.Font.Bold = true;
        ws.Cell("B2").Style.Font.Bold = false;

        using var reloaded = SaveAndReload(wb);

        var reloadedSheet = reloaded.Worksheet("Sheet1");
        await Assert.That(reloadedSheet.Cell("B2").Style.Font.Bold).IsFalse();
        await Assert.That(reloadedSheet.Cell("C2").Style.Font.Bold).IsTrue();
    }

    /// <summary>
    /// Cells of files from the test corpus whose own style differs from their column's. Only an
    /// edited sheet is written from the model, so another cell of the sheet is edited first.
    /// </summary>
    [Test]
    [Arguments(@"TryToLoad\TemplateWithTableSourcePivotTables.xlsx", "K2")]
    [Arguments(@"TryToLoad\LoadImageWithoutTransform2D.xlsx", "A426")]
    [Arguments(@"TryToLoad\LO\xlsm\tdf111974.xlsm", "D194")]
    [Arguments(@"TryToLoad\LO\xlsm\tdf111974.xlsm", "F194")]
    public async Task A_blank_cell_of_an_edited_sheet_keeps_the_style_the_file_gave_it(string resource, string address)
    {
        // Not disposed before the save: a loaded workbook saves by copying the stream it was loaded from.
        using var input = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(resource));
        using var wb = new XLWorkbook(input);

        var expected = new Dictionary<string, XLStyleValue>();
        foreach (var ws in wb.Worksheets)
        {
            var sheet = (XLWorksheet)ws;
            var cell = (XLCell)ws.Cell(address);
            if (!cell.IsEmpty() || sheet.Internals.CellsCollection.StyleSlice[cell.SheetPoint] is null)
                continue;

            if (cell.StyleValue.Equals(sheet.GetInheritedStyleValue(cell.SheetPoint.Row, cell.SheetPoint.Column)))
                continue;

            expected.Add(ws.Name, cell.StyleValue);
            ws.Cell("A1").Value = "Edited";
            await Assert.That(sheet.CellsChangedSinceLoad()).IsTrue().Because(ws.Name);
        }

        await Assert.That(expected).IsNotEmpty();

        using var reloaded = SaveAndReload(wb);

        foreach (var (sheetName, style) in expected)
        {
            var cell = (XLCell)reloaded.Worksheet(sheetName).Cell(address);
            await Assert.That(cell.StyleValue).IsEqualTo(style).Because(sheetName);
        }
    }

    private static XLWorkbook SaveAndReload(XLWorkbook wb)
    {
        var saved = new MemoryStream();
        wb.SaveAs(saved);
        saved.Position = 0;
        return new XLWorkbook(saved);
    }
}
