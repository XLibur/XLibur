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

            // A table over the range has one field per column (defect 4's table consequence),
            // and now saves successfully with the forward address as its own ref (follow-up
            // finding: XLTable.DataRange/relative Range(int,int,int,int) anchored to
            // RangeAddress.FirstAddress directly, which threw on save for a reversed source
            // range; TablePartWriter separately wrote the table's own ref unnormalised). A fresh
            // workbook per corner order, since a table changes what the sheet's used range looks
            // like for the other checks in this loop.
            {
                var tableWb = new XLWorkbook();
                var tableWs = tableWb.Worksheets.Add("Sheet1");
                var tableRange = tableWs.Range($"{ColumnLetter(c1)}{r1}:{ColumnLetter(c2)}{r2}");
                var table = tableRange.CreateTable();
                await Assert.That(table.Fields.Count()).IsEqualTo(expectedWidth);

                using var tableMs = new MemoryStream();
                await Assert.That(() => tableWb.SaveAs(tableMs)).ThrowsNothing();

                // A 1x1 table is a degenerate case XLibur auto-expands with a data row
                // regardless of corner order (not corner-order-sensitive, so not interesting to
                // pin here) - skip the exact-ref comparison for it.
                if (expectedCellCount > 1)
                {
                    using var tableWb2 = new XLWorkbook(tableMs);
                    var reloadedTable = tableWb2.Worksheet("Sheet1").Table(0);
                    await Assert.That(reloadedTable.RangeAddress.ToString()).IsEqualTo(forwardAddress);
                }
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

            // Merging the range makes every cell in it report merged, and the merged range it
            // reports back has the rectangle's own dimensions (user story 8: merge behaviour is
            // identical regardless of corner order). A 1x1 rectangle is a degenerate case -
            // Merge() on a single cell is a no-op that never reaches the range index - so it is
            // skipped here rather than asserted on.
            if (expectedCellCount > 1)
            {
                var mergeWb = new XLWorkbook();
                var mergeWs = mergeWb.Worksheets.Add("Sheet1");
                var mergeRange = mergeWs.Range($"{ColumnLetter(c1)}{r1}:{ColumnLetter(c2)}{r2}");
                mergeRange.Merge();

                foreach (var address in expectedCellAddresses)
                {
                    var cell = mergeWs.Cell(address);
                    await Assert.That(cell.IsMerged()).IsTrue();
                    var merged = cell.MergedRange();
                    await Assert.That(merged).IsNotNull();
                    await Assert.That(merged!.RowCount()).IsEqualTo(expectedHeight);
                    await Assert.That(merged.ColumnCount()).IsEqualTo(expectedWidth);
                }
            }

            // Index intersection (spec user story 10 and its flat-list/point-containment
            // follow-ups): an overlap query finds the range both below and at/above the range
            // index's 20-item QuadTree promotion threshold. The range is added via AddRange to
            // an already-registered validation rather than CreateDataValidation directly: a
            // validation's *first* range is indexed from its already-normalised Ranges
            // projection, which would make this check pass even with every fix in this area
            // reverted (see ReversedRangeGeometryTests.DataValidationIndexFindsReversedRangeBeforePromotion).
            // Twenty filler data validations, well away from the rectangle, force promotion
            // without touching its own coverage.
            {
                var indexWb = new XLWorkbook();
                var indexWs = indexWb.Worksheets.Add("Sheet1");
                var indexDv = indexWs.Cell(1000, 200).CreateDataValidation();
                indexDv.WholeNumber.Between(0, 100);
                indexDv.AddRange(indexWs.Range($"{ColumnLetter(c1)}{r1}:{ColumnLetter(c2)}{r2}"));

                var probeAddress = indexWs.Range($"{ColumnLetter(leftColumn)}{topRow}").RangeAddress;
                var foundBeforePromotion = indexWs.DataValidations.GetAllInRange(probeAddress).ToList();
                await Assert.That(foundBeforePromotion.Count).IsEqualTo(1);

                for (var i = 1; i <= 20; i++)
                    indexWs.Cell(i, 300).CreateDataValidation().WholeNumber.Between(0, 100);

                var foundAfterPromotion = indexWs.DataValidations.GetAllInRange(probeAddress).ToList();
                await Assert.That(foundAfterPromotion.Count).IsEqualTo(1);
            }
        }
    }

    private static string ColumnLetter(int column) => XLHelper.GetColumnLetterFromNumber(column);
}
