using System;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// An operator or a range-accepting function applied to a large reference must compute its elements
/// as they are read, not fill a <c>ScalarValue[height, width]</c> up front (D38).
/// </summary>
/// <remarks>
/// <para>
/// These assert on <b>allocation</b> rather than on elapsed time. The investigation measured the
/// eager form at a steady ~24 MB per million cells, so the quantity is large, linear in the operand
/// and stable — while this machine has roughly 40% run-to-run timing variance, which no threshold
/// survives. A regression here is an order of magnitude, not a few percent.
/// </para>
/// <para>
/// The operands are deliberately four columns wide — 4,194,304 cells, which the investigation
/// measured at 96.0 MB eager. That is far above the ceiling below and far below anything that would
/// exhaust a CI runner. The defect's own reproducer is 566 columns and about 13.6 GB; it lives in
/// the fuzz corpus as <c>XLibur.Fuzz/corpus/formula/operand-implicit-intersection-whole-column</c>
/// rather than here, because a unit test that allocates 13.6 GB before failing is worse than no
/// unit test.
/// </para>
/// <para>
/// Both formulas are evaluated through <c>IXLWorksheet.Evaluate(expression)</c> with no formula
/// address, which is what the fuzz target does. That path cannot use implicit intersection — there
/// is no cell to intersect against — so laziness is the only thing keeping the array unbuilt, and
/// these tests cannot pass by accident through the D38 correctness fix.
/// </para>
/// </remarks>
[SetCulture("en-US")]
public class LazyArrayEvaluationTests
{
    /// <summary>
    /// Four columns is 96.0 MB if the array is materialised. Anything under this is "did not
    /// materialise"; the measured figure for both tests is a few kilobytes.
    /// </summary>
    private const long MaterialisationCeilingBytes = 32L * 1024 * 1024;

    private static IXLWorksheet NewSheet(out XLWorkbook wb)
    {
        wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = 42;
        ws.Cell("B1").Value = 1;
        ws.Cell("B2").Value = 2;
        ws.Cell("B3").Value = 3;
        return ws;
    }

    /// <summary>
    /// Allocation attributable to <paramref name="evaluate"/>, with the expression parsed and cached
    /// first so the measurement covers evaluation rather than the one-off parse.
    /// </summary>
    private static long AllocatedByEvaluating(Func<object> evaluate)
    {
        _ = evaluate();

        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        var before = GC.GetTotalAllocatedBytes(precise: true);
        _ = evaluate();
        return GC.GetTotalAllocatedBytes(precise: true) - before;
    }

    /// <summary>
    /// The operator path: <c>Array.Apply</c>, reached through <c>AnyValue.BinaryOperation</c>.
    /// </summary>
    [Test]
    public async Task Operator_OverAWholeColumnRange_DoesNotMaterialiseTheArray()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            // The result keeps one element, B1, and discards 4,194,303 others.
            await Assert.That(ws.Evaluate("A1+B:E")).IsEqualTo(43.0);

            var allocated = AllocatedByEvaluating(() => ws.Evaluate("A1+B:E"));
            await Assert.That(allocated).IsLessThan(MaterialisationCeilingBytes);
        }
    }

    /// <summary>
    /// The function path: <c>Text.TArray</c>, which filled an array the size of its argument.
    /// </summary>
    /// <remarks>
    /// This is the fuzzer's own input, narrowed from <c>QU:B</c> (458 columns) to <c>B:E</c>. The
    /// full-width form is base64 <c>VChCMTpDMS9WK0FNL1UvK1FVOkIlKzEp</c> and took 59,757 ms and
    /// 11,088 MB. It was found only after <c>Array.Apply</c> was made lazy, because until then the
    /// target could not get past its own seed — the same defect in a second place, which is why the
    /// shape rather than the one call site is worth pinning.
    /// </remarks>
    [Test]
    public async Task RangeAcceptingFunction_OverAWholeColumnRange_DoesNotMaterialiseTheArray()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            const string formula = "T(B1:C1/V+AM/U/+B:E%+1)";

            // V and AM are undefined names, so the answer is #NAME? however it is computed. The
            // point is the cost of reaching it, not the value.
            await Assert.That(ws.Evaluate(formula)).IsEqualTo(XLError.NameNotRecognized);

            var allocated = AllocatedByEvaluating(() => ws.Evaluate(formula));
            await Assert.That(allocated).IsLessThan(MaterialisationCeilingBytes);
        }
    }

    /// <summary>
    /// Computing on access must not change an answer, including when the same element is read more
    /// than once.
    /// </summary>
    /// <remarks>
    /// This is the trade the lazy form makes: a consumer reading an element twice computes it twice.
    /// <c>SUMPRODUCT((B1:B3+0)*(B1:B3+0))</c> reads the same lazy array on both sides of the
    /// multiply, and the spill reads every element of one. Cheap element access is what makes
    /// recomputation acceptable; giving a different answer the second time would not be.
    /// </remarks>
    [Test]
    public async Task LazyElements_AreStableAcrossRepeatedReads()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            // 1*1 + 2*2 + 3*3, with each operand array read twice.
            await Assert.That(ws.Evaluate("SUMPRODUCT((B1:B3+0)*(B1:B3+0))")).IsEqualTo(14.0);

            // Every element of an operator's result, read once each through a spill.
            ws.Cell("D1").SetDynamicFormulaA1("A1+B1:B3");
            await Assert.That(ws.Cell("D1").Value).IsEqualTo(43.0);
            await Assert.That(ws.Cell("D2").Value).IsEqualTo(44.0);
            await Assert.That(ws.Cell("D3").Value).IsEqualTo(45.0);
        }
    }
}
