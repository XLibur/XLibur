using System;
using System.Globalization;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// The sparse walk behind every read of a reference's values runs once per formula that takes a
/// range, so its fixed cost per call is paid tens of thousands of times on an ordinary sheet (#680).
/// </summary>
/// <remarks>
/// This asserts on <b>allocation</b>, which is exact and repeatable, not on elapsed time. A
/// <c>yield</c> method that wraps another <c>yield</c> method costs one more heap object on every
/// call: layering the walk three iterators deep added about 250 bytes to each <c>SUM(D1:H1)</c>,
/// 12.4 MB over the 50,000 formulas of the benchmark that found it.
/// </remarks>
public class SparseWalkAllocationTests
{
    private const int Calls = 1_000;

    /// <summary>
    /// One iterator per call measured 425 bytes in Release and 513 in Debug, on net8.0 and net10.0
    /// alike; the three-deep walk measured 625 and 745. The ceiling sits between the two shapes in
    /// both configurations, so it catches the regression wherever the suite runs.
    /// </summary>
    private const long CeilingBytesPerCall = 560;

    [Test]
    public async Task ReadingAOneAreaReference_AllocatesOneIteratorPerCall()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("Sheet1");
        for (var column = 1; column <= 5; column++)
            ws.Cell(1, column).Value = column;

        var ctx = new CalcContext(wb.CalcEngine, CultureInfo.InvariantCulture, wb, ws, formulaAddress: null);
        var reference = new Reference(new XLRangeAddress(ws, "A1:E1"));

        await Assert.That(Sum(ctx, reference)).IsEqualTo(15.0);

        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        var before = GC.GetTotalAllocatedBytes(precise: true);
        for (var i = 0; i < Calls; i++)
            Sum(ctx, reference);
        var perCall = (GC.GetTotalAllocatedBytes(precise: true) - before) / Calls;
        await Assert.That(perCall).IsLessThan(CeilingBytesPerCall);
    }

    private static double Sum(CalcContext ctx, Reference reference)
    {
        var sum = 0.0;
        foreach (var value in ctx.GetNonBlankValues(reference))
            sum += value.GetNumber();
        return sum;
    }
}
