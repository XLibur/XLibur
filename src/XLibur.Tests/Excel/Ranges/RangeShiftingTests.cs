using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Ranges;

public class RangeShiftingTests
{
    [Test]
    public async Task CellsContentShiftedAfterColumnDeleted()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        SetContent(ws.Cell("D4"));

        ws.Column("C").Delete();

        await AssertContent(ws.Cell("C4"), "D4");
    }

    [Test]
    public async Task CellsContentShiftedAfterRowDeleted()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        SetContent(ws.Cell("D4"));

        ws.Row(3).Delete();

        await AssertContent(ws.Cell("D3"), "D4");
    }

    [Test]
    public async Task CellsContentShiftedAfterColumnInserted()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        SetContent(ws.Cell("D4"));

        ws.Column("C").InsertColumnsBefore(1);

        await AssertContent(ws.Cell("E4"), "D4");
    }

    [Test]
    public async Task CellsContentShiftedAfterRowInserted()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        SetContent(ws.Cell("D4"));

        ws.Row(3).InsertRowsAbove(1);

        await AssertContent(ws.Cell("D5"), "D4");
    }

    [Test]
    public async Task CellsContentShiftAfterRangeDeleted()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        SetContent(ws.Cell("D4"));
        SetContent(ws.Cell("F8"));

        ws.Range("B2:C5").Delete(XLShiftDeletedCells.ShiftCellsLeft);
        ws.Range("E5:F7").Delete(XLShiftDeletedCells.ShiftCellsUp);

        await AssertContent(ws.Cell("B4"), "D4");
        await AssertContent(ws.Cell("F5"), "F8");
    }

    [Test]
    [Arguments("A5:F5")]
    [Arguments("A5:F6")]
    public async Task RangesBelowStayMergedAfterRangeDeleted(string deletedRangeAddress)
    {
        //There is an edge case when a merged range of same size as the deleted range got unmerged (see #2358)
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var deletedRange = ws.Range(deletedRangeAddress);
        var rangeHeight = deletedRange.LastRow()!.RowNumber() - deletedRange.FirstRow()!.RowNumber() + 1;
        var mergedRange = ws.Range(
            deletedRange.LastRow()!.RowNumber() + 1,
            deletedRange.FirstColumn()!.ColumnNumber(),
            deletedRange.LastRow()!.RowNumber() + rangeHeight,
            deletedRange.LastColumn()!.ColumnNumber()
        );
        mergedRange.Merge();

        deletedRange.Delete(XLShiftDeletedCells.ShiftCellsUp);

        await Assert.That(mergedRange.IsMerged()).IsTrue();
        await Assert.That(mergedRange.RangeAddress.ToString()).IsEqualTo(deletedRangeAddress);
    }

    [Test]
    [Arguments("A5:A8")]
    [Arguments("A5:B8")]
    public async Task RangesToTheRightStayMergedAfterRangeDeleted(string deletedRangeAddress)
    {
        //There is an edge case when a merged range of same size as the deleted range got unmerged (see #2358)
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var deletedRange = ws.Range(deletedRangeAddress);
        var rangeWidth = deletedRange.LastColumn()!.ColumnNumber() - deletedRange.FirstColumn()!.ColumnNumber() + 1;
        var mergedRange = ws.Range(
            deletedRange.FirstRow()!.RowNumber(),
            deletedRange.LastColumn()!.ColumnNumber() + 1,
            deletedRange.LastRow()!.RowNumber(),
            deletedRange.LastColumn()!.ColumnNumber() + rangeWidth
        );
        mergedRange.Merge();

        deletedRange.Delete(XLShiftDeletedCells.ShiftCellsLeft);

        await Assert.That(mergedRange.IsMerged()).IsTrue();
        await Assert.That(mergedRange.RangeAddress.ToString()).IsEqualTo(deletedRangeAddress);
    }

    private static void SetContent(IXLCell cell)
    {
        cell.FormulaA1 = $"\"Formula \" & \"{cell.Address}\"";
        cell.Style.Fill.SetBackgroundColor(XLColor.Green);
        cell.CreateComment().AddText("Some comment " + cell.Address);
    }

    private static async Task AssertContent(IXLCell cell, string originalAddress)
    {
        await Assert.That(cell.FormulaA1).IsEqualTo($"\"Formula \" & \"{originalAddress}\"");
        await Assert.That(cell.Style.Fill.BackgroundColor).IsEqualTo(XLColor.Green);
        await Assert.That(cell.HasComment).IsTrue();
        await Assert.That(cell.GetComment().Text).IsEqualTo($"Some comment {originalAddress}");
    }
}
