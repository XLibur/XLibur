using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Ranges;

/// <summary>
/// Pins the pre-filter that lets the shift pass skip a formula without parsing it: a formula whose
/// furthest shiftable reference stops above (or left of) the shifted region cannot be rewritten by
/// that shift, so <c>XLCellFormula.MaxShiftableRow</c>/<c>MaxShiftableColumn</c> is consulted before
/// <c>XLCellFormulaShifter</c> is asked to do the work.
/// <para>
/// Note which direction of error each test can catch. A bound that is too <em>large</em> costs a parse
/// and nothing else — the shifter re-derives the real answer and leaves the formula alone — so no test
/// here would notice one. A bound that is too <em>small</em> strands a formula that should have moved,
/// which is silent data corruption, and that is what every test below is aimed at. Deleting the filter
/// outright leaves them all green; tightening it wrongly does not.
/// </para>
/// </summary>
public class FormulaShiftFilterTests
{
    [Test]
    public async Task ReferenceBelowDeletionIsShifted()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("D1").FormulaA1 = "SUM(A100:C100)";

        ws.Row(50).Delete();

        await Assert.That(ws.Cell("D1").FormulaA1).IsEqualTo("SUM(A99:C99)");
    }

    [Test]
    public async Task ReferenceAboveDeletionIsUntouched()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("D1").FormulaA1 = "SUM(A10:C10)";

        ws.Row(50).Delete();

        await Assert.That(ws.Cell("D1").FormulaA1).IsEqualTo("SUM(A10:C10)");
    }

    /// <summary>
    /// The boundary the filter compares against. A reference ending exactly on the first deleted row is
    /// inside the deletion, not above it, so it must still be rewritten — an off-by-one that used
    /// <c>&lt;=</c> instead of <c>&lt;</c> would strand this.
    /// </summary>
    [Test]
    public async Task ReferenceEndingOnTheFirstDeletedRowIsShifted()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("D1").FormulaA1 = "SUM(A5:A50)";

        ws.Row(50).Delete();

        await Assert.That(ws.Cell("D1").FormulaA1).IsEqualTo("SUM(A5:A49)");
    }

    /// <summary>
    /// The case the cached bound is most likely to get wrong: after the first shift the formula is a new
    /// instance whose bound was carried over rather than re-derived, and the second shift reads that
    /// carried value. A carry that under-estimates would leave the third rewrite undone.
    /// </summary>
    [Test]
    public async Task RepeatedDeletionsKeepShiftingTheSameFormula()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("D1").FormulaA1 = "SUM(A100:C100)";

        ws.Row(10).Delete();
        ws.Row(10).Delete();
        ws.Row(10).Delete();

        await Assert.That(ws.Cell("D1").FormulaA1).IsEqualTo("SUM(A97:C97)");
    }

    /// <summary>
    /// Repeated deletions walking <em>up</em> toward the reference. Each delete leaves the reference
    /// closer to the deletion point, so a carried bound that drifted low would stop the rewrite partway.
    /// </summary>
    [Test]
    public async Task DeletionsApproachingTheReferenceKeepShiftingIt()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("D1").FormulaA1 = "A20";

        for (var row = 15; row >= 10; row--)
            ws.Row(row).Delete();

        await Assert.That(ws.Cell("D1").FormulaA1).IsEqualTo("A14");
    }

    /// <summary>
    /// A reference on another sheet. The bound deliberately ignores which sheet a reference names, so a
    /// formula parked on Sheet2 still shifts when Sheet1 is edited beneath it.
    /// </summary>
    [Test]
    public async Task CrossSheetReferenceIsShiftedWhenTheOtherSheetIsEdited()
    {
        using var wb = new XLWorkbook();
        var source = wb.AddWorksheet("Source");
        var consumer = wb.AddWorksheet("Consumer");
        consumer.Cell("A1").FormulaA1 = "Source!A100";

        source.Row(50).Delete();

        await Assert.That(consumer.Cell("A1").FormulaA1).IsEqualTo("Source!A99");
    }

    [Test]
    public async Task AbsoluteReferenceBelowDeletionIsShifted()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("D1").FormulaA1 = "$A$100";

        ws.Row(50).Delete();

        await Assert.That(ws.Cell("D1").FormulaA1).IsEqualTo("$A$99");
    }

    [Test]
    public async Task DeletingTheReferencedRowsProducesRefError()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("D1").FormulaA1 = "A100";

        ws.Range("A100:A100").Delete(XLShiftDeletedCells.ShiftCellsUp);

        await Assert.That(ws.Cell("D1").FormulaA1).IsEqualTo("#REF!");
    }

    /// <summary>
    /// A whole-column reference names no rows, so its row extent is the whole sheet and no row shift can
    /// ever filter it out. Excel leaves such a reference alone, which is also what the shifter decides —
    /// the point here is that the filter does not reach that conclusion on its own by short-circuiting.
    /// </summary>
    [Test]
    public async Task WholeColumnReferenceSurvivesRowDeletion()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("D1").FormulaA1 = "SUM(B:B)";

        ws.Row(50).Delete();

        await Assert.That(ws.Cell("D1").FormulaA1).IsEqualTo("SUM(B:B)");
    }

    /// <summary>
    /// A formula whose text contains no reference at all still has to shift as a <em>cell</em>, and the
    /// filter must not interfere with that — it only skips rewriting the text.
    /// </summary>
    [Test]
    public async Task ReferencelessFormulaMovesWithItsCell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("D100").FormulaA1 = "1+2";

        ws.Row(50).Delete();

        await Assert.That(ws.Cell("D99").FormulaA1).IsEqualTo("1+2");
        await Assert.That(ws.Cell("D100").FormulaA1).IsEqualTo(string.Empty);
    }

    /// <summary>
    /// An array formula with no references still owns a range that has to be relocated by the shift, so
    /// array formulas opt out of the filter entirely rather than being skipped on an empty bound.
    /// </summary>
    [Test]
    public async Task ArrayFormulaRangeRelocatesWhenTextHasNoReferences()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("D100:D102").FormulaArrayA1 = "1+2";

        ws.Row(50).Delete();

        await Assert.That(ws.Cell("D99").HasArrayFormula).IsTrue();
        await Assert.That(ws.Cell("D101").HasArrayFormula).IsTrue();
        await Assert.That(ws.Cell("D102").HasArrayFormula).IsFalse();
    }

    [Test]
    public async Task ColumnReferenceRightOfDeletionIsShifted()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaA1 = "SUM(J5:L5)";

        ws.Column(5).Delete();

        await Assert.That(ws.Cell("A1").FormulaA1).IsEqualTo("SUM(I5:K5)");
    }

    [Test]
    public async Task ColumnReferenceLeftOfDeletionIsUntouched()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaA1 = "SUM(B5:C5)";

        ws.Column(20).Delete();

        await Assert.That(ws.Cell("A1").FormulaA1).IsEqualTo("SUM(B5:C5)");
    }

    [Test]
    public async Task RepeatedColumnDeletionsKeepShiftingTheSameFormula()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaA1 = "J5";

        ws.Column(5).Delete();
        ws.Column(5).Delete();

        await Assert.That(ws.Cell("A1").FormulaA1).IsEqualTo("H5");
    }

    /// <summary>
    /// Inserting rows uses the same filter with the sign reversed.
    /// </summary>
    [Test]
    public async Task ReferenceBelowInsertionIsShifted()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("D1").FormulaA1 = "SUM(A100:C100)";

        ws.Row(50).InsertRowsAbove(3);

        await Assert.That(ws.Cell("D1").FormulaA1).IsEqualTo("SUM(A103:C103)");
    }

    [Test]
    public async Task ReferenceAboveInsertionIsUntouched()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("D1").FormulaA1 = "SUM(A10:C10)";

        ws.Row(50).InsertRowsAbove(3);

        await Assert.That(ws.Cell("D1").FormulaA1).IsEqualTo("SUM(A10:C10)");
    }

    /// <summary>
    /// A sheet rename rewrites the formula text, which drops the cached bound. If it did not, a formula
    /// renamed and then shifted would be measured against a bound derived from text it no longer has.
    /// </summary>
    [Test]
    public async Task ShiftAfterSheetRenameStillRewritesTheReference()
    {
        using var wb = new XLWorkbook();
        var source = wb.AddWorksheet("Source");
        var consumer = wb.AddWorksheet("Consumer");
        consumer.Cell("A1").FormulaA1 = "Source!A100";

        source.Name = "Renamed";
        source.Row(50).Delete();

        await Assert.That(consumer.Cell("A1").FormulaA1).IsEqualTo("Renamed!A99");
    }
}
