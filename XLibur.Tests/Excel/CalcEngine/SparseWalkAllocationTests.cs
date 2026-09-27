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

    /// <summary>
    /// A small area is walked point by point, with no slice enumerators (#686). Point by point it
    /// measured 172 bytes in Release and 260 in Debug on net10.0, and 169 and 257 on net8.0;
    /// through the slice enumerators it cost the same as the sparse walk above. The ceiling sits
    /// between the two, like the one above.
    /// </summary>
    private const long SmallAreaCeilingBytesPerCall = 340;

    /// <summary>
    /// An area larger than the point-by-point walk takes, so the sparse walk reads it.
    /// </summary>
    [Test]
    public async Task ReadingAOneAreaReference_AllocatesOneIteratorPerCall()
    {
        using var wb = new XLWorkbook();
        var ws = NewSheet(wb);
        var reference = new Reference(new XLRangeAddress(ws, "A1:E100"));

        var perCall = await MeasureSum(wb, ws, reference);
        await Assert.That(perCall).IsLessThan(CeilingBytesPerCall);
    }

    [Test]
    public async Task ReadingASmallArea_AllocatesNoSliceEnumerators()
    {
        using var wb = new XLWorkbook();
        var ws = NewSheet(wb);
        var reference = new Reference(new XLRangeAddress(ws, "A1:E1"));

        var perCall = await MeasureSum(wb, ws, reference);
        await Assert.That(perCall).IsLessThan(SmallAreaCeilingBytesPerCall);
    }

    private static XLWorksheet NewSheet(XLWorkbook wb)
    {
        var ws = (XLWorksheet)wb.AddWorksheet("Sheet1");
        for (var column = 1; column <= 5; column++)
            ws.Cell(1, column).Value = column;
        return ws;
    }

    private static async Task<long> MeasureSum(XLWorkbook wb, XLWorksheet ws, Reference reference)
    {
        var ctx = new CalcContext(wb.CalcEngine, CultureInfo.InvariantCulture, wb, ws, formulaPoint: null);

        await Assert.That(Sum(ctx, reference)).IsEqualTo(15.0);

        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        var before = GC.GetTotalAllocatedBytes(precise: true);
        for (var i = 0; i < Calls; i++)
            Sum(ctx, reference);
        return (GC.GetTotalAllocatedBytes(precise: true) - before) / Calls;
    }

    private static double Sum(CalcContext ctx, Reference reference)
    {
        var sum = 0.0;
        foreach (var value in ctx.GetNonBlankValues(reference))
            sum += value.GetNumber();
        return sum;
    }
}
