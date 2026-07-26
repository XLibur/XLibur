using System;
using System.Linq;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Ranges;

public class InsertingRangesTests
{
    [Test]
    public async Task InsertingColumnsPreservesFormatting()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet");
        var column1 = ws.Column(1);
        column1.Style.Fill.SetBackgroundColor(XLColor.FrenchLilac);
        column1.Cell(2).Style.Fill.SetBackgroundColor(XLColor.Fulvous);
        var column2 = ws.Column(2);
        column2.Style.Fill.SetBackgroundColor(XLColor.Xanadu);
        column2.Cell(2).Style.Fill.SetBackgroundColor(XLColor.MacaroniAndCheese);

        column1.InsertColumnsAfter(1);
        column1.InsertColumnsBefore(1);
        column2.InsertColumnsBefore(1);

        await Assert.That(ws.Column(1).Style.Fill.BackgroundColor).IsEqualTo(ws.Style.Fill.BackgroundColor);
        await Assert.That(ws.Column(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FrenchLilac);
        await Assert.That(ws.Column(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FrenchLilac);
        await Assert.That(ws.Column(4).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FrenchLilac);
        await Assert.That(ws.Column(5).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Xanadu);

        await Assert.That(ws.Cell(2, 1).Style.Fill.BackgroundColor).IsEqualTo(ws.Style.Fill.BackgroundColor);
        await Assert.That(ws.Cell(2, 2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Fulvous);
        await Assert.That(ws.Cell(2, 3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Fulvous);
        await Assert.That(ws.Cell(2, 4).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Fulvous);
        await Assert.That(ws.Cell(2, 5).Style.Fill.BackgroundColor).IsEqualTo(XLColor.MacaroniAndCheese);
    }

    [Test]
    public async Task InsertingRowsAbove()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet");

        ws.Cell("B3").SetValue("X")
            .CellBelow().SetValue("B");

        var r = ws.Range("B4").InsertRowsAbove(1).First();
        r.Cell(1).SetValue("A");

        await Assert.That(ws.Cell("B3").GetText()).IsEqualTo("X");
        await Assert.That(ws.Cell("B4").GetText()).IsEqualTo("A");
        await Assert.That(ws.Cell("B5").GetText()).IsEqualTo("B");
    }

    [Test]
    public async Task InsertingRowsPreservesFormatting()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet");
        var row1 = ws.Row(1);
        row1.Style.Fill.SetBackgroundColor(XLColor.FrenchLilac);
        row1.Cell(2).Style.Fill.SetBackgroundColor(XLColor.Fulvous);
        var row2 = ws.Row(2);
        row2.Style.Fill.SetBackgroundColor(XLColor.Xanadu);
        row2.Cell(2).Style.Fill.SetBackgroundColor(XLColor.MacaroniAndCheese);

        row1.InsertRowsBelow(1);
        row1.InsertRowsAbove(1);
        row2.InsertRowsAbove(1);

        await Assert.That(ws.Row(1).Style.Fill.BackgroundColor).IsEqualTo(ws.Style.Fill.BackgroundColor);
        await Assert.That(ws.Row(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FrenchLilac);
        await Assert.That(ws.Row(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FrenchLilac);
        await Assert.That(ws.Row(4).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FrenchLilac);
        await Assert.That(ws.Row(5).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Xanadu);

        await Assert.That(ws.Cell(1, 2).Style.Fill.BackgroundColor).IsEqualTo(ws.Style.Fill.BackgroundColor);
        await Assert.That(ws.Cell(2, 2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Fulvous);
        await Assert.That(ws.Cell(3, 2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Fulvous);
        await Assert.That(ws.Cell(4, 2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Fulvous);
        await Assert.That(ws.Cell(5, 2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.MacaroniAndCheese);
    }

    [Test]
    public async Task InsertingRowsPreservesComments()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        ws.Cell("A1").SetValue("Insert Below");
        ws.Cell("A2").SetValue("Already existing cell");
        ws.Cell("A3").SetValue("Cell with comment").GetComment().AddText("Comment here");

        ws.Row(1).InsertRowsBelow(2);
        await Assert.That(ws.Cell("A5").GetComment().Text).IsEqualTo("Comment here");
    }

    [Test]
    public async Task InsertingColumnsPreservesComments()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        ws.Cell("A1").SetValue("Insert to the right");
        ws.Cell("B1").SetValue("Already existing cell");
        ws.Cell("C1").SetValue("Cell with comment").GetComment().AddText("Comment here");

        ws.Column(1).InsertColumnsAfter(2);
        await Assert.That(ws.Cell("E1").GetComment().Text).IsEqualTo("Comment here");
    }

    [Test]
    [Arguments("C4:F7", "C4:F7", 2, "E4:H7")] // Coincide, shift right
    [Arguments("C4:F7", "C4:F7", -2, "C4:D7")] // Coincide, shift left
    [Arguments("D5:E6", "C4:F7", 2, "F5:G6")] // Inside, shift right
    [Arguments("D5:E6", "C4:F7", -2, "C5:C6")] // Inside, shift left
    [Arguments("B4:G7", "C4:F7", 2, "B4:I7")] // Includes, shift right
    [Arguments("B4:G7", "C4:F7", -2, "B4:E7")] // Includes, shift left
    [Arguments("B4:E7", "C4:F7", 2, "B4:G7")] // Intersects at left, shift right
    [Arguments("B4:E7", "C4:F7", -2, "B4:C7")] // Intersects at left, shift left
    [Arguments("D4:G7", "C4:F7", 2, "F4:I7")] // Intersects at right, shift right
    [Arguments("D4:G7", "C4:F7", -2, "C4:E7")] // Intersects at right, shift left
    [Arguments("A5:B6", "C4:F7", 2, "A5:B6")] // No intersection, at left, shift right
    [Arguments("A5:B6", "C4:F7", -1, "A5:B6")] // No intersection, at left, shift left
    [Arguments("H5:I6", "C4:F7", 2, "J5:K6")] // No intersection, at right, shift right
    [Arguments("H5:I6", "C4:F7", -2, "F5:G6")] // No intersection, at right, shift left
    [Arguments("C8:F11", "C4:F7", 2, "C8:F11")] // Different rows
    [Arguments("B1:B8", "A1:C4", 1, "B1:B8")]  // More rows, shift right
    [Arguments("B1:B8", "A1:C4", -1, "B1:B8")]  // More rows, shift left
    public async Task ShiftColumnsValid(string thisRangeAddress, string shiftedRangeAddress, int shiftedColumns, string expectedRange)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        var thisRange = ws.Range(thisRangeAddress) as XLRange;
        var shiftedRange = ws.Range(shiftedRangeAddress) as XLRange;

        thisRange!.WorksheetRangeShiftedColumns(shiftedRange!, shiftedColumns);

        await Assert.That(thisRange.RangeAddress.IsValid).IsTrue();
        await Assert.That(thisRange.RangeAddress.ToString()).IsEqualTo(expectedRange);
    }

    [Test]
    [Arguments("B1:B4", "A1:C4", -2)] // Shift left too much
    public async Task ShiftColumnsInvalid(string thisRangeAddress, string shiftedRangeAddress, int shiftedColumns)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        var thisRange = ws.Range(thisRangeAddress) as XLRange;
        var shiftedRange = ws.Range(shiftedRangeAddress) as XLRange;

        thisRange!.WorksheetRangeShiftedColumns(shiftedRange!, shiftedColumns);

        await Assert.That(thisRange.RangeAddress.IsValid).IsFalse();
    }

    [Test]
    [Arguments("C4:F7", "C4:F7", 2, "C6:F9")]   // Coincide, shift down
    [Arguments("C4:F7", "C4:F7", -2, "C4:F5")]   // Coincide, shift up
    [Arguments("D5:E6", "C4:F7", 2, "D7:E8")]   // Inside, shift down
    [Arguments("D5:E6", "C4:F7", -2, "D4:E4")]   // Inside, shift up
    [Arguments("C3:F8", "C4:F7", 2, "C3:F10")]  // Includes, shift down
    [Arguments("C3:F8", "C4:F7", -2, "C3:F6")]   // Includes, shift up
    [Arguments("C3:F6", "C4:F7", 2, "C3:F8")]   // Intersects at top, shift down
    [Arguments("C2:F6", "C4:F7", -3, "C2:F3")]   // Intersects at top, shift up to the sheet boundary
    [Arguments("C3:F6", "C4:F7", -2, "C3:F4")]   // Intersects at top, shift up
    [Arguments("C5:F8", "C4:F7", 2, "C7:F10")]  // Intersects at bottom, shift down
    [Arguments("C5:F8", "C4:F7", -2, "C4:F6")]   // Intersects at bottom, shift up
    [Arguments("C1:F3", "C4:F7", 2, "C1:F3")]   // No intersection, at top, shift down
    [Arguments("C1:F3", "C4:F7", -2, "C1:F3")]   // No intersection, at top, shift up
    [Arguments("C8:F10", "C4:F7", 2, "C10:F12")] // No intersection, at bottom, shift down
    [Arguments("C8:F10", "C4:F7", -2, "C6:F8")]   // No intersection, at bottom, shift up
    [Arguments("G4:J7", "C4:F7", 2, "G4:J7")]   // Different columns
    [Arguments("A2:D2", "A1:C4", 1, "A2:D2")]   // More columns, shift down
    [Arguments("A2:D2", "A1:C4", -1, "A2:D2")]   // More columns, shift up
    public async Task ShiftRowsValid(string thisRangeAddress, string shiftedRangeAddress, int shiftedRows, string expectedRange)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        var thisRange = ws.Range(thisRangeAddress) as XLRange;
        var shiftedRange = ws.Range(shiftedRangeAddress) as XLRange;

        thisRange!.WorksheetRangeShiftedRows(shiftedRange!, shiftedRows);

        await Assert.That(thisRange.RangeAddress.IsValid).IsTrue();
        await Assert.That(thisRange.RangeAddress.ToString()).IsEqualTo(expectedRange);
    }

    [Test]
    [Arguments("A2:C2", "A1:C4", -2)] // Shift up too much
    public async Task ShiftRowsInvalid(string thisRangeAddress, string shiftedRangeAddress, int shiftedRows)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        var thisRange = ws.Range(thisRangeAddress) as XLRange;
        var shiftedRange = ws.Range(shiftedRangeAddress) as XLRange;

        thisRange!.WorksheetRangeShiftedRows(shiftedRange!, shiftedRows);

        await Assert.That(thisRange.RangeAddress.IsValid).IsFalse();
    }

    [Test]
    public async Task InsertZeroColumnsFails()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet1");
        var range = ws.FirstCell().AsRange();
        await Assert.That(() => range.InsertColumnsAfter(0)).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => range.InsertColumnsBefore(0)).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task InsertNegativeNumberOfColumnsFails()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet1");
        var range = ws.FirstCell().AsRange();
        await Assert.That(() => range.InsertColumnsAfter(-1)).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => range.InsertColumnsBefore(-1)).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task InsertTooLargeNumberOfColumnsFails()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet1");
        var range = ws.FirstCell().AsRange();
        await Assert.That(() => range.InsertColumnsAfter(16385)).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => range.InsertColumnsBefore(16385)).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task InsertZeroRowsFails()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet1");
        var range = ws.FirstCell().AsRange();
        await Assert.That(() => range.InsertRowsAbove(0)).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => range.InsertRowsBelow(0)).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task InsertNegativeNumberOfRowsFails()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet1");
        var range = ws.FirstCell().AsRange();
        await Assert.That(() => range.InsertRowsAbove(-1)).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => range.InsertRowsBelow(-1)).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task InsertTooLargeNumberOrRowsFails()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet1");
        var range = ws.FirstCell().AsRange();
        await Assert.That(() => range.InsertRowsAbove(1048577)).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => range.InsertRowsBelow(1048577)).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task MergedRangesConsistencyWhenInsertingRows()
    {
        // https://github.com/XLibur/XLibur/issues/1013
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        //create a merged row
        ws.Cell("A1").Value = "Merged Row(1) of Range (A1:F1)";
        ws.Range("A1:F1").Row(1).Merge();

        var row = ws.FirstRow();

        // Add some lines and copy a format & merging
        for (var r = 1; r <= 10; r++)
        {
            row.InsertRowsBelow(1);         // insert a row below row 1, as a row 2
            row.CopyTo(row.RowBelow());     // copy format and merging from row 1 to row 2

            var duplicates = ws.MergedRanges
                .GroupBy(s => s.ToString())
                .Where(g => g.Count() > 1)
                .Select(y => new { Element = y.Key, Counter = y.Count() })
                .ToList();

            await Assert.That(duplicates.Count).IsEqualTo(0);
        }
    }
}
