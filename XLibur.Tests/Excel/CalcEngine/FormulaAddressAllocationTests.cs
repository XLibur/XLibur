using System;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.Coordinates;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// Every evaluation of a cell's formula is told which cell it is in, and most formulas never ask.
/// The cell was passed as <see cref="IXLAddress"/>, and <c>XLAddress</c> is a struct, so each
/// evaluation boxed it: 40 bytes per formula, 2 MB over the 50,000 formulas of the benchmark that
/// found it (#686).
/// </summary>
/// <remarks>
/// This asserts on <b>allocation</b>, which is exact and repeatable, not on elapsed time.
/// </remarks>
public class FormulaAddressAllocationTests
{
    private const int Calls = 1_000;

    /// <summary>
    /// Evaluating <c>1+2</c> in a cell measured 105 bytes with the point unboxed and 145 with the
    /// address boxed, in Debug and Release alike. The ceiling sits between the two.
    /// </summary>
    private const long CeilingBytesPerCall = 125;

    [Test]
    public async Task EvaluatingAFormulaInACell_DoesNotBoxTheAddress()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("Sheet1");
        var engine = wb.CalcEngine;
        var point = new Point(2, 3);

        await Assert.That(engine.EvaluateFormula("1+2", wb, ws, point).GetNumber()).IsEqualTo(3.0);

        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        var before = GC.GetTotalAllocatedBytes(precise: true);
        for (var i = 0; i < Calls; i++)
            engine.EvaluateFormula("1+2", wb, ws, point);
        var perCall = (GC.GetTotalAllocatedBytes(precise: true) - before) / Calls;
        await Assert.That(perCall).IsLessThan(CeilingBytesPerCall);
    }
}
