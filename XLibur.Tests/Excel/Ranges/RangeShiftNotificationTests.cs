using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Ranges;

/// <summary>
/// Characterisation tests for range-shift notification: how many passes a structural edit runs, and
/// the order live ranges are visited in.
/// <para>
/// Both are load-bearing. A pass visits every live range the shift can reach, so a batch edit that
/// notified per row would be quadratic. And the visit order is what stops two live ranges colliding
/// on one repository key: when a range shifts onto an address another live range still occupies, the
/// repository's add fails silently and the moved range detaches — it keeps its address but stops
/// receiving later shifts. Processing furthest-first vacates the target key before it is needed.
/// </para>
/// <para>
/// These pin the behaviour of the existing repository scan so spec 05's index-driven replacement can
/// be shown to preserve it.
/// </para>
/// </summary>
public class RangeShiftNotificationTests
{
    [Test]
    public async Task InsertRowsAboveRunsOneNotificationPassRegardlessOfCount()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet();
        ws.Cell("A1")!.Value = "anchor";
        var before = ws.RangeShiftPasses;

        ws.Row(3).InsertRowsAbove(1000);

        await Assert.That(ws.RangeShiftPasses - before).IsEqualTo(1);
    }

    [Test]
    public async Task InsertRowsBelowRunsOneNotificationPassRegardlessOfCount()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet();
        ws.Cell("A1")!.Value = "anchor";
        var before = ws.RangeShiftPasses;

        ws.Row(3).InsertRowsBelow(1000);

        await Assert.That(ws.RangeShiftPasses - before).IsEqualTo(1);
    }

    [Test]
    public async Task InsertColumnsBeforeRunsOneNotificationPassRegardlessOfCount()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet();
        ws.Cell("A1")!.Value = "anchor";
        var before = ws.RangeShiftPasses;

        ws.Column(3).InsertColumnsBefore(500);

        await Assert.That(ws.RangeShiftPasses - before).IsEqualTo(1);
    }

    [Test]
    public async Task InsertColumnsAfterRunsOneNotificationPassRegardlessOfCount()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet();
        ws.Cell("A1")!.Value = "anchor";
        var before = ws.RangeShiftPasses;

        ws.Column(3).InsertColumnsAfter(500);

        await Assert.That(ws.RangeShiftPasses - before).IsEqualTo(1);
    }

    [Test]
    public async Task RangeDeleteRunsOneNotificationPass()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet();
        ws.Cell("A1")!.Value = "anchor";
        var before = ws.RangeShiftPasses;

        ws.Range("B2:C500")!.Delete(XLShiftDeletedCells.ShiftCellsUp);

        await Assert.That(ws.RangeShiftPasses - before).IsEqualTo(1);
    }

    [Test]
    public async Task InsertingRowsShiftsEveryLiveRangeBelowTheInsertion()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var above = ws.Range("A1:A5");
        var straddling = ws.Range("A4:A12");
        var below = ws.Range("A6:A10");
        var alsoBelow = ws.Range("A8:A12");

        ws.Row(6).InsertRowsAbove(2);

        await Assert.That(above.RangeAddress.ToString()).IsEqualTo("A1:A5");
        await Assert.That(straddling.RangeAddress.ToString()).IsEqualTo("A4:A14");
        await Assert.That(below.RangeAddress.ToString()).IsEqualTo("A8:A12");
        await Assert.That(alsoBelow.RangeAddress.ToString()).IsEqualTo("A10:A14");
    }

    /// <summary>
    /// The aliasing case the visit order exists for: <c>A6:A10</c> shifts onto <c>A8:A12</c>, which is
    /// another live range's address. If the mover went first the repository would already hold that key
    /// and the move would be dropped, detaching the range — visible only on the *next* shift, which
    /// would leave it behind. Asserting a second insert still moves both is what catches that.
    /// </summary>
    [Test]
    public async Task RangeShiftedOntoAnotherLiveRangeAddressStaysRegistered()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var mover = ws.Range("A6:A10");
        var occupant = ws.Range("A8:A12");

        ws.Row(6).InsertRowsAbove(2);

        await Assert.That(mover.RangeAddress.ToString()).IsEqualTo("A8:A12");
        await Assert.That(occupant.RangeAddress.ToString()).IsEqualTo("A10:A14");

        ws.Row(6).InsertRowsAbove(2);

        await Assert.That(mover.RangeAddress.ToString()).IsEqualTo("A10:A14");
        await Assert.That(occupant.RangeAddress.ToString()).IsEqualTo("A12:A16");
    }

    [Test]
    public async Task ColumnRangeShiftedOntoAnotherLiveRangeAddressStaysRegistered()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var mover = ws.Range("F1:J1");
        var occupant = ws.Range("H1:L1");

        ws.Column(6).InsertColumnsBefore(2);

        await Assert.That(mover.RangeAddress.ToString()).IsEqualTo("H1:L1");
        await Assert.That(occupant.RangeAddress.ToString()).IsEqualTo("J1:N1");

        ws.Column(6).InsertColumnsBefore(2);

        await Assert.That(mover.RangeAddress.ToString()).IsEqualTo("J1:N1");
        await Assert.That(occupant.RangeAddress.ToString()).IsEqualTo("L1:P1");
    }

    /// <summary>
    /// Deletion visits in the opposite order — nearest-first — for the same reason: ranges collapse
    /// upward onto the addresses of ranges above them.
    /// </summary>
    [Test]
    public async Task RangeShiftedUpOntoAnotherLiveRangeAddressStaysRegistered()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var upper = ws.Range("A10:A14");
        var lower = ws.Range("A12:A16");

        ws.Range("A6:A7").Delete(XLShiftDeletedCells.ShiftCellsUp);

        await Assert.That(upper.RangeAddress.ToString()).IsEqualTo("A8:A12");
        await Assert.That(lower.RangeAddress.ToString()).IsEqualTo("A10:A14");

        ws.Range("A6:A7").Delete(XLShiftDeletedCells.ShiftCellsUp);

        await Assert.That(upper.RangeAddress.ToString()).IsEqualTo("A6:A10");
        await Assert.That(lower.RangeAddress.ToString()).IsEqualTo("A8:A12");
    }

    [Test]
    public async Task EntireColumnRangesAreNotShiftedByRowInsertion()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var entireColumn = ws.Column(2).AsRange();

        ws.Row(3).InsertRowsAbove(5);

        await Assert.That(entireColumn.RangeAddress.FirstAddress.RowNumber).IsEqualTo(1);
        await Assert.That(entireColumn.RangeAddress.LastAddress.RowNumber).IsEqualTo(XLHelper.MaxRowNumber);
    }

    /// <summary>
    /// A row insertion scoped to columns C:E does not move a range that only partly overlaps it
    /// horizontally: B1:D5 starts left of the inserted block, so the insert cuts across it rather than
    /// pushing it down, and Excel leaves it where it is.
    /// <para>
    /// This exercises the cross-axis containment condition in <c>CollectRangesShiftedByRows</c>, but
    /// note what it can and cannot catch. The filter only decides which ranges get *visited*;
    /// <c>XLRangeShiftHelper.ShiftRows</c> re-checks the same condition and returns the address
    /// unchanged, so a filter that is too loose costs time and nothing else — deleting the containment
    /// check outright leaves every test here passing. A filter that is too *tight* strands a range that
    /// should have moved, which is a correctness bug, and that direction does fail these tests.
    /// </para>
    /// </summary>
    [Test]
    public async Task PartiallyOverlappingRangeIsNotShiftedByScopedRowInsertion()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var straddling = ws.Range("B1:D5");
        var contained = ws.Range("C1:E5");

        ws.Range("C3:E3").InsertRowsAbove(1);

        await Assert.That(straddling.RangeAddress.ToString()).IsEqualTo("B1:D5");
        await Assert.That(contained.RangeAddress.ToString()).IsEqualTo("C1:E6");
    }

    /// <summary>
    /// Column twin of <see cref="PartiallyOverlappingRangeIsNotShiftedByScopedRowInsertion"/>: A1:B4
    /// reaches above the rows the insertion is scoped to, so it is cut across rather than moved.
    /// </summary>
    [Test]
    public async Task PartiallyOverlappingRangeIsNotShiftedByScopedColumnInsertion()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var straddling = ws.Range("A1:B4");
        var contained = ws.Range("A3:B5");

        ws.Range("C3:C5").InsertColumnsBefore(1);

        await Assert.That(straddling.RangeAddress.ToString()).IsEqualTo("A1:B4");
        await Assert.That(contained.RangeAddress.ToString()).IsEqualTo("A3:B5");
    }

    [Test]
    public async Task EntireRowRangesAreNotShiftedByColumnInsertion()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var entireRow = ws.Row(2).AsRange();

        ws.Column(3).InsertColumnsBefore(5);

        await Assert.That(entireRow.RangeAddress.FirstAddress.ColumnNumber).IsEqualTo(1);
        await Assert.That(entireRow.RangeAddress.LastAddress.ColumnNumber).IsEqualTo(XLHelper.MaxColumnNumber);
    }
}
