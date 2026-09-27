using System;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.Coordinates;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// SUM, COUNT, AVERAGE, MIN, MAX and the other tally functions read each reference argument's
/// values. They read them through an iterator, one heap object per reference per evaluation, the
/// largest single allocation left in a <c>SUM(D1:H1)</c> once its tree was parsed (#686).
/// </summary>
/// <remarks>
/// This asserts on <b>allocation</b>, which is exact and repeatable, not on elapsed time.
/// </remarks>
public class TallyAllocationTests
{
    private const int Calls = 1_000;

    /// <summary>
    /// Evaluating either formula in a cell measured 105 bytes with the struct enumerator, in Debug
    /// and Release, on net8.0 and net10.0: only the context, as for <c>1+2</c>. Through the iterator
    /// it measured 273 in Release and 361 in Debug. The ceiling sits between the two.
    /// </summary>
    private const long CeilingBytesPerCall = 190;

    [Test]
    [Arguments("SUM(A1:E1)", 15d)]
    [Arguments("AVERAGEA(A1:E1)", 3d)]
    public async Task TallyingAReference_AllocatesNoIterator(string formula, double expected)
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("Sheet1");
        for (var column = 1; column <= 5; column++)
            ws.Cell(1, column).Value = column;

        var engine = wb.CalcEngine;
        var point = new Point(3, 1);

        await Assert.That(engine.EvaluateFormula(formula, wb, ws, point).GetNumber()).IsEqualTo(expected);

        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        var before = GC.GetAllocatedBytesForCurrentThread();
        for (var i = 0; i < Calls; i++)
            engine.EvaluateFormula(formula, wb, ws, point);
        var perCall = (GC.GetAllocatedBytesForCurrentThread() - before) / Calls;
        await Assert.That(perCall).IsLessThan(CeilingBytesPerCall);
    }
}
