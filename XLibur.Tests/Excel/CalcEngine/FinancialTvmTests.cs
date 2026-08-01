using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;
// PV / NPER / PPMT / NPV / IRR / RATE — time-value-of-money functions added alongside FV/PMT/IPMT.
public class FinancialTvmTests
{
    private const double Tolerance = 1e-4;
    private const double IterativeTolerance = 1e-3;

    private static XLWorksheet NewSheet(out XLWorkbook wb)
    {
        wb = new XLWorkbook();
        return (XLWorksheet)wb.AddWorksheet("Sheet1");
    }

    [Test]
    public async Task Pv_ComputesPresentValue()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            await Assert.That((double)ws.Evaluate("PV(0, 10, 100)")).IsEqualTo(-1000d).Within(Tolerance);
            await Assert.That((double)ws.Evaluate("PV(0.05, 10, 100)")).IsEqualTo(-772.173493d).Within(Tolerance);
            // Only a future value, no periodic payment: -fv / (1 + rate).
            await Assert.That((double)ws.Evaluate("PV(0.1, 1, 0, 100)")).IsEqualTo(-90.909091d).Within(Tolerance);
        }
    }

    [Test]
    public async Task Nper_ComputesNumberOfPeriods()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            await Assert.That((double)ws.Evaluate("NPER(0, -100, 1000)")).IsEqualTo(10d).Within(Tolerance);
            await Assert.That((double)ws.Evaluate("NPER(0.05, -100, 1000)")).IsEqualTo(14.206699d).Within(Tolerance);
        }
    }

    [Test]
    public async Task Nper_ZeroRateAndZeroPayment_ReturnsNumberInvalid()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            await Assert.That(ws.Evaluate("NPER(0, 0, 1000)")).IsEqualTo(XLError.NumberInvalid);
        }
    }

    [Test]
    public async Task Ppmt_ComputesPrincipalPortion()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            // First period of a 3-period 10% loan of 1000: PMT - IPMT.
            await Assert.That((double)ws.Evaluate("PPMT(0.1, 1, 3, 1000)")).IsEqualTo(-302.114804d).Within(Tolerance);
        }
    }

    [Test]
    public async Task Ppmt_PeriodOutOfRange_ReturnsNumberInvalid()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            await Assert.That(ws.Evaluate("PPMT(0.1, 5, 3, 1000)")).IsEqualTo(XLError.NumberInvalid);
        }
    }

    [Test]
    public async Task Npv_DiscountsCashflows()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            await Assert.That((double)ws.Evaluate("NPV(0.1, 100, 100, 100)")).IsEqualTo(248.685199d).Within(Tolerance);

            ws.Cell("A1").Value = 100;
            ws.Cell("A2").Value = 100;
            ws.Cell("A3").Value = 100;
            await Assert.That((double)ws.Evaluate("NPV(0.1, A1:A3)")).IsEqualTo(248.685199d).Within(Tolerance);
        }
    }

    [Test]
    public async Task Irr_FindsInternalRateOfReturn()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = -100;
            ws.Cell("A2").Value = 110;
            await Assert.That((double)ws.Evaluate("IRR(A1:A2)")).IsEqualTo(0.1d).Within(IterativeTolerance);

            ws.Cell("B1").Value = -1000;
            ws.Cell("B2").Value = 500;
            ws.Cell("B3").Value = 500;
            ws.Cell("B4").Value = 500;
            await Assert.That((double)ws.Evaluate("IRR(B1:B4)")).IsEqualTo(0.233751d).Within(IterativeTolerance);
        }
    }

    [Test]
    public async Task Irr_TooFewValues_ReturnsNumberInvalid()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            await Assert.That(ws.Evaluate("IRR({100})")).IsEqualTo(XLError.NumberInvalid);
        }
    }

    [Test]
    public async Task Rate_SolvesForPeriodicRate()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            // 100 now, pay 110 in one period -> 10% per period.
            await Assert.That((double)ws.Evaluate("RATE(1, -110, 100)")).IsEqualTo(0.1d).Within(IterativeTolerance);
            // Inverts PMT(0.05, 10, 1000) = -129.5046.
            await Assert.That((double)ws.Evaluate("RATE(10, -129.504575, 1000)")).IsEqualTo(0.05d).Within(IterativeTolerance);
        }
    }
}
