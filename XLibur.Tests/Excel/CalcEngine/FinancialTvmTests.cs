using NUnit.Framework;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

// PV / NPER / PPMT / NPV / IRR / RATE — time-value-of-money functions added alongside FV/PMT/IPMT.
[TestFixture]
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
    public void Pv_ComputesPresentValue()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual(-1000d, (double)ws.Evaluate("PV(0, 10, 100)"), Tolerance);
            Assert.AreEqual(-772.173493d, (double)ws.Evaluate("PV(0.05, 10, 100)"), Tolerance);
            // Only a future value, no periodic payment: -fv / (1 + rate).
            Assert.AreEqual(-90.909091d, (double)ws.Evaluate("PV(0.1, 1, 0, 100)"), Tolerance);
        }
    }

    [Test]
    public void Nper_ComputesNumberOfPeriods()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual(10d, (double)ws.Evaluate("NPER(0, -100, 1000)"), Tolerance);
            Assert.AreEqual(14.206699d, (double)ws.Evaluate("NPER(0.05, -100, 1000)"), Tolerance);
        }
    }

    [Test]
    public void Nper_ZeroRateAndZeroPayment_ReturnsNumberInvalid()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual(XLError.NumberInvalid, ws.Evaluate("NPER(0, 0, 1000)"));
        }
    }

    [Test]
    public void Ppmt_ComputesPrincipalPortion()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            // First period of a 3-period 10% loan of 1000: PMT - IPMT.
            Assert.AreEqual(-302.114804d, (double)ws.Evaluate("PPMT(0.1, 1, 3, 1000)"), Tolerance);
        }
    }

    [Test]
    public void Ppmt_PeriodOutOfRange_ReturnsNumberInvalid()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual(XLError.NumberInvalid, ws.Evaluate("PPMT(0.1, 5, 3, 1000)"));
        }
    }

    [Test]
    public void Npv_DiscountsCashflows()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual(248.685199d, (double)ws.Evaluate("NPV(0.1, 100, 100, 100)"), Tolerance);

            ws.Cell("A1").Value = 100;
            ws.Cell("A2").Value = 100;
            ws.Cell("A3").Value = 100;
            Assert.AreEqual(248.685199d, (double)ws.Evaluate("NPV(0.1, A1:A3)"), Tolerance);
        }
    }

    [Test]
    public void Irr_FindsInternalRateOfReturn()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = -100;
            ws.Cell("A2").Value = 110;
            Assert.AreEqual(0.1d, (double)ws.Evaluate("IRR(A1:A2)"), IterativeTolerance);

            ws.Cell("B1").Value = -1000;
            ws.Cell("B2").Value = 500;
            ws.Cell("B3").Value = 500;
            ws.Cell("B4").Value = 500;
            Assert.AreEqual(0.233751d, (double)ws.Evaluate("IRR(B1:B4)"), IterativeTolerance);
        }
    }

    [Test]
    public void Irr_TooFewValues_ReturnsNumberInvalid()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual(XLError.NumberInvalid, ws.Evaluate("IRR({100})"));
        }
    }

    [Test]
    public void Rate_SolvesForPeriodicRate()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            // 100 now, pay 110 in one period -> 10% per period.
            Assert.AreEqual(0.1d, (double)ws.Evaluate("RATE(1, -110, 100)"), IterativeTolerance);
            // Inverts PMT(0.05, 10, 1000) = -129.5046.
            Assert.AreEqual(0.05d, (double)ws.Evaluate("RATE(10, -129.504575, 1000)"), IterativeTolerance);
        }
    }
}
