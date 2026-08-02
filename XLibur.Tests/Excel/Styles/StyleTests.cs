using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Styles;

public class StyleTests
{
    [Test]
    public async Task EmptyCellWithQuotePrefixNotTreatedAsEmpty()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Empty cell with quote prefix:");
            var cell = ws.FirstCell().CellRight() as XLCell;

            await Assert.That(cell.IsEmpty()).IsTrue();
            cell.Value = String.Empty;
            cell.Style.IncludeQuotePrefix = true;

            await Assert.That(cell.IsEmpty()).IsTrue();
            await Assert.That(cell.IsEmpty(XLCellsUsedOptions.All)).IsFalse();

            wb.SaveAs(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();
            var cell = (XLCell)ws.Cell("B1");
            await Assert.That(cell.MemorySstId).IsEqualTo(1);

            await Assert.That(cell.IsEmpty()).IsTrue();
            await Assert.That(cell.IsEmpty(XLCellsUsedOptions.All)).IsFalse();
        }
    }

    [Test]
    [Arguments("A1", DisplayName = "First cell")]
    [Arguments("A2", DisplayName = "Cell from initialized row")]
    [Arguments("B1", DisplayName = "Cell from initialized column")]
    [Arguments("D4", DisplayName = "Initialized cell")]
    [Arguments("F6", DisplayName = "Non-initialized cell")]
    public async Task CellTakesWorksheetStyle(string cellAddress)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Column(2);
        ws.Row(2);
        ws.Cell("D4").Value = "Non empty";
        ws.Style.Font.SetFontName("Arial");
        ws.Style.Font.SetFontSize(9);

        var cell = ws.Cell(cellAddress);
        await Assert.That(cell.Style.Font.FontName).IsEqualTo("Arial");
        await Assert.That(cell.Style.Font.FontSize).IsEqualTo(9);
    }

    [Test]
    [MethodDataSource(nameof(StylizedEntities))]
    public async Task WorksheetStyleAffectsAllNestedEntities(string entity, Func<IXLWorksheet, IXLStyle> getEntityStyle)
    {
        _ = entity; // Present so each generated case is identifiable; see StylizedEntities.

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        ws.Style.Font.FontSize = 8;

        var style = getEntityStyle(ws);

        await Assert.That(style.Font.FontSize).IsEqualTo(8);
    }

    // https://github.com/XLibur/XLibur/issues/1813
    [Test]
    public async Task RowColors()
    {
        await TestHelper.CreateAndCompare(() =>
        {
            var wb = new XLWorkbook();
            {
                var ws = wb.Worksheets.Add("Row Settings 1");
                ws.Style.Fill.BackgroundColor = XLColor.Green;

                var row1 = ws.Row(2);
                row1.Style.Fill.BackgroundColor = XLColor.Red;
                row1.Height = 30;

                var row2 = ws.Row(4);
                row2.Style.Fill.BackgroundColor = XLColor.DarkOrange;
                row2.Height = 3;
            }

            {
                var ws = wb.Worksheets.Add("Row Settings 2");
                ws.Style.Fill.BackgroundColor = XLColor.Red;

                var row1 = ws.Row(2);
                row1.Style.Fill.BackgroundColor = XLColor.Red;

                var row2 = ws.Row(4);
                row2.Style.Fill.BackgroundColor = XLColor.DarkOrange;
                row2.Height = 3;
            }

            {
                var ws = wb.Worksheets.Add("Row Settings 3");
                ws.Style.Fill.BackgroundColor = XLColor.Red;

                var row1 = ws.Row(2);
                row1.Style.Fill.BackgroundColor = XLColor.Red;
                row1.Height = 30;

                var row2 = ws.Row(4);
                row2.Style.Fill.BackgroundColor = XLColor.DarkOrange;
                row2.Height = 3;
            }

            return wb;
        }, @"Other\StyleReferenceFiles\RowColors\output.xlsx");
    }

    [Test]
    public async Task Style_for_cells_without_explicitly_set_style_uses_combination_of_row_and_columns_styles()
    {
        // If a style for a cell hasn't been explicitly set (e.g. though `cell.Style.Font
        // .SetBold(true)`), it is not yet instantiated to save memory and the actual value
        // is determined by the column style and row style. Generally speaking, the axis that
        // had its value set explicitly has a precedence, but because we can't detect that with
        // current structure, use difference from worksheet as an indication of explicitly set
        // value instead.
        // If row and column style components differ, the cells at the cross are pinged, thus test
        // sets different components for each axis.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var rowStyle = ws.Row(4).Style
            .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
            .Fill.SetBackgroundColor(XLColor.Blue)
            .SetIncludeQuotePrefix()
            .Protection.SetLocked(true);

        var colStyle = ws.Column(2).Style
            .Border.SetBottomBorder(XLBorderStyleValues.Double)
            .Font.SetFontName("Arial")
            .NumberFormat.SetNumberFormatId((int)XLPredefinedFormat.Number.Precision2);

        var crossCellStyle = ws.Cell(4, 2).Style;
        await Assert.That(crossCellStyle.Alignment.Horizontal).IsEqualTo(XLAlignmentHorizontalValues.Center);
        await Assert.That(crossCellStyle.Border.BottomBorder).IsEqualTo(XLBorderStyleValues.Double);
        await Assert.That(crossCellStyle.Fill.BackgroundColor).IsEqualTo(XLColor.Blue);
        await Assert.That(crossCellStyle.IncludeQuotePrefix).IsTrue();
        await Assert.That(crossCellStyle.NumberFormat.NumberFormatId).IsEqualTo((int)XLPredefinedFormat.Number.Precision2);
        await Assert.That(crossCellStyle.Protection.Locked).IsTrue();

        var rowCellStyle = ws.Cell(4, 3).Style;
        await Assert.That(rowCellStyle).IsEqualTo(rowStyle);

        var colCellStyle = ws.Cell(5, 2).Style;
        await Assert.That(colCellStyle).IsEqualTo(colStyle);
    }

    // NUnit's TestCaseData.SetName(...) gave each case a readable name. TUnit derives case
    // names from the arguments, and a Func<> argument renders as nothing useful, so the
    // description travels as an explicit first argument instead.
    public static IEnumerable<(string Entity, Func<IXLWorksheet, IXLStyle> GetStyle)> StylizedEntities()
    {
        yield return ("Worksheet", ws => ws.Style);

        yield return ("Columns()", ws => ws.Columns().Style);
        yield return ("Columns(1, 3)", ws => ws.Columns(1, 3).Style);
        yield return ("Columns(\"B:F\")", ws => ws.Columns("B:F").Style);
        yield return ("Columns(\"B\", \"F\")", ws => ws.Columns("B", "F").Style);
        yield return ("Column(5)", ws => ws.Column(5).Style);
        yield return ("Column(\"D\")", ws => ws.Column("D").Style);

        yield return ("Rows()", ws => ws.Rows().Style);
        yield return ("Rows(1, 3)", ws => ws.Rows(1, 3).Style);
        yield return ("Rows(\"1:3\")", ws => ws.Rows("1:3").Style);
        yield return ("Row(5)", ws => ws.Row(5).Style);

        yield return ("Cells()", ws => ws.Cells().Style);
        yield return ("Cells(\"B2, D4\")", ws => ws.Cells("B2,D4").Style);
        yield return ("Cell(\"F6\")", ws => ws.Cell("F6").Style);
        yield return ("Cell(2, 3)", ws => ws.Cell(2, 3).Style);

        yield return ("Ranges(\"F6:H9,I8:K10\")", ws => ws.Ranges("F6:H9,I8:K10").Style);
        yield return ("Range(\"G8:H10\")", ws => ws.Range("G8:H10").Style);
        yield return ("Range(\"G8:H10\").Column(1)", ws => ws.Range("G8:H10").Column(1).Style);
        yield return ("Range(\"G8:H10\").Row(2)", ws => ws.Range("G8:H10").Row(2).Style);
    }
}
