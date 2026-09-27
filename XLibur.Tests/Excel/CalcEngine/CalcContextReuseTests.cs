using System;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.CalcEngine.Exceptions;
using XLibur.Excel.Coordinates;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// Every top-level evaluation of a normal formula needs a <c>CalcContext</c>. The engine keeps one
/// and reuses it instead of allocating 105 bytes per formula (#686). Nothing one evaluation sets on
/// it may reach the next, and an evaluation nested inside another must not take the context the
/// outer one is still using.
/// </summary>
public class CalcContextReuseTests
{
    private const int Calls = 1_000;

    /// <summary>
    /// Evaluating <c>1+2</c> in a cell measured 0 bytes with the context reused, in Debug and
    /// Release on net8.0 and net10.0, and 104 with a new context each time. The ceiling sits
    /// between the two.
    /// </summary>
    private const long CeilingBytesPerCall = 52;

    /// <remarks>
    /// Counts the test's own thread only: <c>GC.GetTotalAllocatedBytes</c> also counts what other
    /// threads allocate during the loop, which <c>--coverage</c> makes happen (#697).
    /// </remarks>
    [Test]
    public async Task EvaluatingAFormulaInACell_ReusesTheContext()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("Sheet1");
        var engine = wb.CalcEngine;
        var point = new Point(2, 3);

        await Assert.That(engine.EvaluateFormula("1+2", wb, ws, point).GetNumber()).IsEqualTo(3.0);

        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        var before = GC.GetAllocatedBytesForCurrentThread();
        for (var i = 0; i < Calls; i++)
            engine.EvaluateFormula("1+2", wb, ws, point);
        var perCall = (GC.GetAllocatedBytesForCurrentThread() - before) / Calls;
        await Assert.That(perCall).IsLessThan(CeilingBytesPerCall);
    }

    /// <summary>
    /// The outer evaluation holds the context while <c>A1</c>, a dirty formula, is evaluated inside
    /// it. If the inner one took the same context, it would leave its own cell, A1, behind, and the
    /// outer <c>ROW()</c>, read after <c>A1</c>, would answer 1 instead of 7. <c>A1*1</c> reads the
    /// value of A1 before <c>ROW()</c> runs; <c>A1+ROW()</c> would not, because a reference operand
    /// is read only when its operator combines the two.
    /// </summary>
    [Test]
    public async Task ANestedEvaluationGetsAContextOfItsOwn()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaA1 = "ROW()*10";

        // A fresh engine has no spare context yet, so the outer and inner evaluations would each
        // build one and could never share. One evaluation first leaves a spare behind.
        await Assert.That(ws.Evaluate("1")).IsEqualTo(1);

        await Assert.That(ws.Evaluate("A1*1+ROW()", "C7")).IsEqualTo(17);
    }

    /// <summary>
    /// A cell's evaluation knows its cell. The next evaluation, which has none, must not find it.
    /// </summary>
    [Test]
    public async Task TheCellOfOneEvaluationDoesNotReachTheNext()
    {
        await AssertSameAsFresh(
            ws => ws.Cell("C5").FormulaA1 = "ROW()",
            ws => _ = ws.Cell("C5").Value,
            ws => ws.Evaluate("ROW()"));
    }

    /// <summary>
    /// A cell's evaluation intersects a range operand with its row (D38). The next evaluation,
    /// which has no cell, keeps the range whole. The operator is at the top of the formula: inside
    /// a function's argument the intersection is off whatever the context says.
    /// </summary>
    [Test]
    public async Task IntersectionOfOneEvaluationDoesNotReachTheNext()
    {
        await AssertSameAsFresh(
            ws =>
            {
                ws.Cell("A1").Value = 15;
                ws.Cell("A2").Value = 10;
                ws.Cell("B1").Value = 5;
                ws.Cell("C2").FormulaA1 = "A1:A2-B1";
            },
            ws => _ = ws.Cell("C2").Value,
            ws => ws.Evaluate("A1:A2-B1"));
    }

    /// <summary>
    /// A recursive evaluation remembers each dirty cell it evaluated, so that it does not evaluate
    /// one twice. The next evaluation must evaluate it again, or it reads the old value.
    /// </summary>
    [Test]
    public async Task CellValuesOneEvaluationRemembersDoNotReachTheNext()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B1").Value = 1;
        ws.Cell("A1").FormulaA1 = "B1";

        await Assert.That(ws.Evaluate("A1*2")).IsEqualTo(2);

        ws.Cell("B1").Value = 5;
        await Assert.That(ws.Evaluate("A1*2")).IsEqualTo(10);
    }

    /// <summary>
    /// A recursive evaluation, such as <c>Evaluate</c>, calculates a dirty cell it reads on the
    /// spot. An evaluation on the calculation chain must not: it signals the chain to calculate
    /// the cell first, which is how the chain keeps its order and finds circular references.
    /// </summary>
    [Test]
    public async Task TheRecursiveChoiceOfOneEvaluationDoesNotReachTheNext()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet();
        sheet.Cell("A1").FormulaA1 = "1+1";
        var ws = (XLWorksheet)sheet;
        var engine = wb.CalcEngine;
        var point = new Point(5, 5);

        await Assert.That(engine.EvaluateFormula("1", wb, ws, point, recursive: true).GetNumber()).IsEqualTo(1.0);

        await Assert.That(() => { engine.EvaluateFormula("A1*1", wb, ws, point, recursive: false); })
            .Throws<GettingDataException>();
    }

    /// <summary>
    /// A sheet-only recalculation reads other sheets as they stand. A later read must calculate
    /// a dirty cell on another sheet first, not keep treating that sheet as left alone. If it
    /// did, the calculation chain would wait for a cell it never calculates, and never end.
    /// </summary>
    [Test]
    public async Task TheSheetOfASheetRecalculationDoesNotReachTheNext()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("S1");
        var ws2 = wb.AddWorksheet("S2");
        ws2.Cell("B1").Value = 1;
        ws2.Cell("A1").FormulaA1 = "B1";
        ws1.Cell("A1").FormulaA1 = "S2!A1*2";

        // Bounded, so that a hang fails this test instead of stalling the suite.
        var value = await Task.Run(() =>
        {
            ws1.RecalculateAllFormulas();
            ws2.Cell("B1").Value = 3;
            return ws1.Cell("A1").Value;
        }).WaitAsync(TimeSpan.FromSeconds(10));

        await Assert.That(value).IsEqualTo(6);
    }

    /// <summary>
    /// Runs <paramref name="setUp"/> and <paramref name="second"/> on a fresh workbook, and on one
    /// where <paramref name="first"/> ran in between, and expects the same outcome of
    /// <paramref name="second"/>: the same value, or an exception of the same type.
    /// </summary>
    private static async Task AssertSameAsFresh(Action<IXLWorksheet> setUp, Action<IXLWorksheet> first,
        Func<IXLWorksheet, XLCellValue> second)
    {
        string fresh;
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            setUp(ws);
            fresh = Outcome(() => second(ws));
        }

        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            setUp(ws);
            first(ws);
            await Assert.That(Outcome(() => second(ws))).IsEqualTo(fresh);
        }
    }

    private static string Outcome(Func<XLCellValue> evaluate)
    {
        try
        {
            return evaluate().ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        catch (Exception e)
        {
            return e.GetType().FullName!;
        }
    }
}
