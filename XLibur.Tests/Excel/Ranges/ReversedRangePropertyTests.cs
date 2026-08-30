using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Ranges;

/// <summary>
/// The centrepiece test for spec 36: for a rectangle and each of its four corner orderings,
/// every geometry consumer must agree. Where <see cref="ReversedRangeGeometryTests"/> pins one
/// symptom per defect, this drives the same rectangle through all of them at once, for several
/// rectangle shapes, so a regression in any single consumer fails here even if no named test
/// happens to cover that shape.
/// </summary>
public class ReversedRangePropertyTests
{
    /// <param name="topRow">Top row of the rectangle in its forward (normalised) form.</param>
    /// <param name="leftColumn">Left column of the rectangle in its forward form.</param>
    /// <param name="bottomRow">Bottom row of the rectangle in its forward form.</param>
    /// <param name="rightColumn">Right column of the rectangle in its forward form.</param>
    [Test]
    [Arguments(2, 2, 5, 5)] // square
    [Arguments(1, 1, 1, 1)] // single cell
    [Arguments(3, 1, 4, 10)] // wide, short
    [Arguments(1, 3, 10, 4)] // tall, narrow
    public async Task AllCornerOrdersAgreeAcrossEveryGeometryConsumer(int topRow, int leftColumn, int bottomRow,
        int rightColumn)
    {
        var expectedWidth = rightColumn - leftColumn + 1;
        var expectedHeight = bottomRow - topRow + 1;
        var expectedCellCount = expectedWidth * expectedHeight;
        var forwardAddress = $"{ColumnLetter(leftColumn)}{topRow}:{ColumnLetter(rightColumn)}{bottomRow}";

        // The four ways a user can name the same rectangle's corners.
        var cornerOrders = new (int R1, int C1, int R2, int C2)[]
        {
            (topRow, leftColumn, bottomRow, rightColumn), // forward
            (bottomRow, leftColumn, topRow, rightColumn), // rows reversed
            (topRow, rightColumn, bottomRow, leftColumn), // columns reversed
            (bottomRow, rightColumn, topRow, leftColumn), // both reversed
        };

        string[]? expectedCellAddresses = null;

        foreach (var (r1, c1, r2, c2) in cornerOrders)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            var range = ws.Range($"{ColumnLetter(c1)}{r1}:{ColumnLetter(c2)}{r2}");

            // Cell count, and the address's own spans (unaffected by the bug this spec fixes,
            // but pinned here as a baseline every consumer below is checked against).
            await Assert.That(range.Cells().Count()).IsEqualTo(expectedCellCount);
            await Assert.That(range.RangeAddress.RowSpan).IsEqualTo(expectedHeight);
            await Assert.That(range.RangeAddress.ColumnSpan).IsEqualTo(expectedWidth);

            // Row and column counts (defect 4).
            await Assert.That(range.RowCount()).IsEqualTo(expectedHeight);
            await Assert.That(range.ColumnCount()).IsEqualTo(expectedWidth);

            // The set of cells enumeration reaches, independent of corner order.
            var cellAddresses = range.Cells().Select(c => c.Address.ToString()!).OrderBy(a => a).ToArray();
            expectedCellAddresses ??= cellAddresses;
            await Assert.That(cellAddresses).IsEquivalentTo(expectedCellAddresses);

            // Style application reaches the same cells as enumeration (defect 3).
            range.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            foreach (var address in expectedCellAddresses)
                await Assert.That(ws.Cell(address).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);

            // Consolidation includes the range, at its forward address (defect 5).
            var ranges = new XLRanges { range };
            var consolidated = ranges.Consolidate().ToList();
            await Assert.That(consolidated.Count).IsEqualTo(1);
            await Assert.That(consolidated[0].RangeAddress.ToString()).IsEqualTo(forwardAddress);

            // A table over the range has one field per column (defect 4's table consequence).
            // A fresh workbook: creating a table here and then saving the same workbook below
            // (for the data-validation check) hits an unrelated, pre-existing defect - XLTable's
            // relative Range(int,int,int,int) anchors to RangeAddress.FirstAddress directly
            // rather than the normalised top-left, so DataRange throws on save for a reversed
            // source range. Out of this spec's scope (Table isn't in its consumer list); see the
            // final report. ColumnCount() itself - what field count is derived from - is already
            // covered by RowCountAndColumnCountOnReversedRangeReturnPositiveMagnitudes.
            {
                var tableWb = new XLWorkbook();
                var tableWs = tableWb.Worksheets.Add("Sheet1");
                var tableRange = tableWs.Range($"{ColumnLetter(c1)}{r1}:{ColumnLetter(c2)}{r2}");
                var table = tableRange.CreateTable();
                await Assert.That(table.Fields.Count()).IsEqualTo(expectedWidth);
            }

            // A data validation on the range survives a save/reload at its forward address
            // (defect 2), in its own workbook so an unrelated table (above) cannot affect it.
            {
                var dvWb = new XLWorkbook();
                var dvWs = dvWb.Worksheets.Add("Sheet1");
                var dvRange = dvWs.Range($"{ColumnLetter(c1)}{r1}:{ColumnLetter(c2)}{r2}");
                dvRange.CreateDataValidation().WholeNumber.Between(0, 100);
                using var ms = new MemoryStream();
                dvWb.SaveAs(ms);
                using var dvWb2 = new XLWorkbook(ms);
                var reloadedValidations = dvWb2.Worksheet("Sheet1").DataValidations.ToList();
                await Assert.That(reloadedValidations.Count).IsEqualTo(1);
                var reloadedRanges = reloadedValidations[0].Ranges.ToList();
                await Assert.That(reloadedRanges.Count).IsEqualTo(1);
                await Assert.That(reloadedRanges[0].RangeAddress.ToString()).IsEqualTo(forwardAddress);
            }
        }
    }

    private static string ColumnLetter(int column) => XLHelper.GetColumnLetterFromNumber(column);
}
