using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Ranges;

/// <summary>
/// Covers deleting a set of whole rows in one call, where the set need not be contiguous.
/// <para>
/// The batched path re-points every formula once against the whole set of deleted rows instead of once
/// per row, so the thing worth pinning is that it lands in the same place the row-at-a-time path does.
/// Several tests assert exactly that equivalence rather than a hand-computed answer, because the
/// row-at-a-time path is the definition of correct here and a hand-computed expectation would just be a
/// second chance to get the same arithmetic wrong.
/// </para>
/// </summary>
public class BatchRowDeleteTests
{
    [Test]
    public async Task ScatteredDeleteMatchesRowAtATimeDelete()
    {
        using var perRow = Build(60, out var a);
        using var batched = Build(60, out var b);
        var targets = Enumerable.Range(1, 20).Select(i => i * 3).Reverse().ToList();

        foreach (var row in targets)
            a.Row(row).Delete();

        b.Rows(Spec(targets)).Delete();

        for (var row = 1; row <= 60; row++)
        {
            await Assert.That(b.Cell(row, 4).FormulaA1).IsEqualTo(a.Cell(row, 4).FormulaA1);
            await Assert.That(b.Cell(row, 1).GetString()).IsEqualTo(a.Cell(row, 1).GetString());
        }
    }

    [Test]
    public async Task ContiguousDeleteMatchesRowAtATimeDelete()
    {
        using var perRow = Build(40, out var a);
        using var batched = Build(40, out var b);
        var targets = new[] { 12, 11, 10, 9, 8 };

        foreach (var row in targets)
            a.Row(row).Delete();

        b.Rows("8:12").Delete();

        for (var row = 1; row <= 40; row++)
            await Assert.That(b.Cell(row, 4).FormulaA1).IsEqualTo(a.Cell(row, 4).FormulaA1);
    }

    /// <summary>
    /// Runs and singletons in the same call. The set is split into contiguous runs and removed furthest
    /// down first, so a mis-ordered split shows up as rows going missing from the wrong places.
    /// </summary>
    [Test]
    public async Task MixedRunsAndSingletonsMatchRowAtATimeDelete()
    {
        using var perRow = Build(40, out var a);
        using var batched = Build(40, out var b);
        var targets = new[] { 30, 21, 20, 19, 12, 5, 4 };

        foreach (var row in targets)
            a.Row(row).Delete();

        b.Rows(Spec(targets)).Delete();

        for (var row = 1; row <= 40; row++)
        {
            await Assert.That(b.Cell(row, 1).GetString()).IsEqualTo(a.Cell(row, 1).GetString());
            await Assert.That(b.Cell(row, 4).FormulaA1).IsEqualTo(a.Cell(row, 4).FormulaA1);
        }
    }

    [Test]
    public async Task ScatteredDeleteShiftsRowsUpByTheCountRemovedAboveThem()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        for (var row = 1; row <= 10; row++)
            ws.Cell(row, 1).Value = row;

        ws.Rows("2:2,5:5,8:8").Delete();

        // 1,3,4,6,7,9,10 survive and close up in order.
        var expected = new[] { 1, 3, 4, 6, 7, 9, 10 };
        for (var i = 0; i < expected.Length; i++)
            await Assert.That(ws.Cell(i + 1, 1).GetValue<int>()).IsEqualTo(expected[i]);

        await Assert.That(ws.Cell(8, 1).IsEmpty()).IsTrue();
    }

    /// <summary>
    /// A reference spanning the deleted rows loses exactly the rows deleted inside it, and the rows
    /// deleted above it move the whole thing up.
    /// </summary>
    [Test]
    public async Task ReferenceSpanningDeletedRowsShrinksByTheRowsRemovedInsideIt()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("E1").FormulaA1 = "SUM(A10:A20)";

        ws.Rows("2:2,12:12,15:15").Delete();

        // One row deleted above the reference moves it up by one; two deleted inside shorten it.
        await Assert.That(ws.Cell("E1").FormulaA1).IsEqualTo("SUM(A9:A17)");
    }

    /// <summary>
    /// Only the reference is destroyed, not the call around it, so the surviving text is
    /// <c>SUM(#REF!)</c> — which is also what deleting the rows one at a time produces.
    /// </summary>
    [Test]
    public async Task ReferenceEntirelyInsideTheDeletedSetBecomesRefError()
    {
        using var batched = new XLWorkbook();
        var b = batched.AddWorksheet();
        b.Cell("E1").FormulaA1 = "SUM(A10:A12)";
        b.Rows("10:12").Delete();

        using var perRow = new XLWorkbook();
        var a = perRow.AddWorksheet();
        a.Cell("E1").FormulaA1 = "SUM(A10:A12)";
        a.Row(12).Delete();
        a.Row(11).Delete();
        a.Row(10).Delete();

        await Assert.That(b.Cell("E1").FormulaA1).IsEqualTo("SUM(#REF!)");
        await Assert.That(b.Cell("E1").FormulaA1).IsEqualTo(a.Cell("E1").FormulaA1);
    }

    [Test]
    public async Task ReferenceAboveEveryDeletedRowIsUntouched()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("E1").FormulaA1 = "SUM(A2:A5)";

        ws.Rows("10:10,20:20,30:30").Delete();

        await Assert.That(ws.Cell("E1").FormulaA1).IsEqualTo("SUM(A2:A5)");
    }

    [Test]
    public async Task FormulaOnAnotherSheetIsRepointedByTheBatchedDelete()
    {
        using var wb = new XLWorkbook();
        var source = wb.AddWorksheet("Source");
        var consumer = wb.AddWorksheet("Consumer");
        consumer.Cell("A1").FormulaA1 = "SUM(Source!A10:A20)";

        source.Rows("2:2,12:12,15:15").Delete();

        await Assert.That(consumer.Cell("A1").FormulaA1).IsEqualTo("SUM(Source!A9:A17)");
    }

    /// <summary>
    /// An array formula carries a stored range that only the per-cell shift relocates, so a workbook
    /// containing one takes the row-at-a-time path. The batched call must still produce the same answer.
    /// </summary>
    [Test]
    public async Task ArrayFormulaWorkbookFallsBackAndStillShifts()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("F20:F22").FormulaArrayA1 = "1+2";
        ws.Cell("E1").FormulaA1 = "SUM(A10:A20)";

        ws.Rows("2:2,12:12").Delete();

        await Assert.That(ws.Cell("E1").FormulaA1).IsEqualTo("SUM(A9:A18)");
        await Assert.That(ws.Cell("F18").HasArrayFormula).IsTrue();
        await Assert.That(ws.Cell("F20").HasArrayFormula).IsTrue();
        await Assert.That(ws.Cell("F21").HasArrayFormula).IsFalse();
    }

    /// <summary>
    /// The row map collapses duplicates and sorts, so a caller handing over rows in any order — or the
    /// same row twice — gets the same result as a clean ascending set.
    /// </summary>
    [Test]
    public async Task UnsortedAndDuplicatedRowsDeleteEachRowOnce()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        for (var row = 1; row <= 10; row++)
            ws.Cell(row, 1).Value = row;

        ws.Rows("8:8,2:2,8:8,5:5,2:2").Delete();

        var expected = new[] { 1, 3, 4, 6, 7, 9, 10 };
        for (var i = 0; i < expected.Length; i++)
            await Assert.That(ws.Cell(i + 1, 1).GetValue<int>()).IsEqualTo(expected[i]);
    }

    [Test]
    public async Task RowPropertiesFollowTheSurvivingRows()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        for (var row = 1; row <= 10; row++)
            ws.Cell(row, 1).Value = row;

        ws.Row(7).Height = 33;

        ws.Rows("2:2,5:5").Delete();

        // Row 7 loses two rows above it and becomes row 5.
        await Assert.That(ws.Cell(5, 1).GetValue<int>()).IsEqualTo(7);
        await Assert.That(ws.Row(5).Height).IsEqualTo(33);
    }

    [Test]
    public async Task DefinedNameShiftsAcrossAScatteredDelete()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Range("A10:A20").AddToNamed("Block", XLScope.Workbook);

        ws.Rows("2:2,12:12,15:15").Delete();

        await Assert.That(wb.DefinedName("Block").RefersTo).IsEqualTo("Data!$A$9:$A$17");
    }

    [Test]
    public async Task MergedRangeShiftsAcrossAScatteredDelete()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("A10:B11").Merge();

        ws.Rows("2:2,5:5").Delete();

        await Assert.That(ws.MergedRanges.Single().RangeAddress.ToString()).IsEqualTo("A8:B9");
    }

    [Test]
    public async Task LiveRangeShiftsAcrossAScatteredDelete()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var tracked = ws.Range("A10:A20");

        ws.Rows("2:2,12:12,15:15").Delete();

        await Assert.That(tracked.RangeAddress.ToString()).IsEqualTo("A9:A17");
    }

    /// <summary>
    /// The values a batched delete leaves behind must still recalculate, i.e. the formulas it rewrote
    /// are marked dirty rather than keeping the cached result of the pre-delete layout.
    /// </summary>
    [Test]
    public async Task FormulasRecalculateAfterABatchedDelete()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        for (var row = 1; row <= 10; row++)
            ws.Cell(row, 1).Value = row * 100;

        ws.Cell("C1").FormulaA1 = "A8";
        await Assert.That(ws.Cell("C1").GetValue<int>()).IsEqualTo(800);

        ws.Rows("2:2,4:4").Delete();

        // A8 became A6, which now holds what was in row 8 -> 800 still, but via the moved cell.
        await Assert.That(ws.Cell("C1").FormulaA1).IsEqualTo("A6");
        await Assert.That(ws.Cell("C1").GetValue<int>()).IsEqualTo(800);
    }

    private static string Spec(System.Collections.Generic.IEnumerable<int> rows)
        => string.Join(",", rows.Select(r => $"{r}:{r}"));

    private static XLWorkbook Build(int rows, out IXLWorksheet ws)
    {
        var wb = new XLWorkbook();
        ws = wb.AddWorksheet("Sheet1");
        for (var row = 1; row <= rows; row++)
        {
            ws.Cell(row, 1).Value = $"r{row}";
            ws.Cell(row, 4).FormulaA1 = $"SUM(A{row}:C{row})";
        }

        return wb;
    }
}
