using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// ISPMT, CUMIPMT, CUMPRINC, FVSCHEDULE, MIRR, XNPV and XIRR. Expected values come from the worked
/// examples in Microsoft's per-function documentation; each comment shows the documented formula
/// applied to the arguments.
/// </summary>
public class FinancialCashFlowTests
{
    private static XLWorksheet NewSheet(out XLWorkbook wb)
    {
        wb = new XLWorkbook();
        return (XLWorksheet)wb.AddWorksheet("Sheet1");
    }

    [Test]
    // Microsoft's ISPMT example: an 8,000,000 loan over three years at 10%. The principal falls
    // linearly, so period `per` still owes (1 - per/nper) of it: pv * rate * (per/nper - 1).
    [Arguments("ISPMT(0.1/12, 1, 3*12, 8000000)", -64814.814814814818d)]
    [Arguments("ISPMT(0.1, 1, 3, 8000000)", -533333.33333333337d)]
    [Arguments("ISPMT(0.1, 3, 3, 8000000)", 0d)] // The last period owes nothing.
    public async Task IsPmt_ReferenceExamplesFromExcelDocumentation(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-6);
    }

    [Test]
    public async Task IsPmt_ZeroPeriodsIsDivisionByZero()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("ISPMT(0.1, 1, 0, 8000000)")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    // Microsoft's CUMIPMT/CUMPRINC example: a 125,000 mortgage over 30 years at 9%.
    [Arguments("CUMIPMT(0.09/12, 30*12, 125000, 13, 24, 0)", -11135.232130750162d)] // The second year's interest.
    [Arguments("CUMIPMT(0.09/12, 30*12, 125000, 1, 1, 0)", -937.5d)] // 125000 * 0.09/12.
    [Arguments("CUMPRINC(0.09/12, 30*12, 125000, 13, 24, 0)", -934.10712342088d)] // The second year's principal.
    [Arguments("CUMPRINC(0.09/12, 30*12, 125000, 1, 1, 0)", -68.278271563389d)]
    public async Task Cumulative_ReferenceExamplesFromExcelDocumentation(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-4);
    }

    [Test]
    public async Task Cumulative_SinglePeriodMatchesIpmtAndPpmt()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "CUMIPMT(0.09/12, 360, 125000, 7, 7, 0)";
            ws.Cell("A2").FormulaA1 = "IPMT(0.09/12, 7, 360, 125000)";
            ws.Cell("B1").FormulaA1 = "CUMPRINC(0.09/12, 360, 125000, 7, 7, 0)";
            ws.Cell("B2").FormulaA1 = "PPMT(0.09/12, 7, 360, 125000)";

            await Assert.That((double)ws.Cell("A1").Value).IsEqualTo((double)ws.Cell("A2").Value).Within(1e-9);
            await Assert.That((double)ws.Cell("B1").Value).IsEqualTo((double)ws.Cell("B2").Value).Within(1e-9);
        }
    }

    [Test]
    public async Task Cumulative_OverTheWholeTermSumsToTheTotalPayments()
    {
        // Every payment is either interest or principal, and the principal repaid over the full term
        // is the amount borrowed.
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "CUMIPMT(0.09/12, 360, 125000, 1, 360, 0)";
            ws.Cell("A2").FormulaA1 = "CUMPRINC(0.09/12, 360, 125000, 1, 360, 0)";
            ws.Cell("A3").FormulaA1 = "PMT(0.09/12, 360, 125000) * 360";

            await Assert.That((double)ws.Cell("A2").Value).IsEqualTo(-125000d).Within(1e-6);
            await Assert.That((double)ws.Cell("A1").Value + (double)ws.Cell("A2").Value)
                .IsEqualTo((double)ws.Cell("A3").Value).Within(1e-6);
        }
    }

    [Test]
    [Arguments("CUMIPMT(0, 360, 125000, 1, 12, 0)")] // The rate must be positive.
    [Arguments("CUMIPMT(0.0075, 0, 125000, 1, 12, 0)")] // The term must be positive.
    [Arguments("CUMIPMT(0.0075, 360, 0, 1, 12, 0)")] // The loan must be positive.
    [Arguments("CUMIPMT(0.0075, 360, 125000, 0, 12, 0)")] // Periods are one-based.
    [Arguments("CUMIPMT(0.0075, 360, 125000, 13, 12, 0)")] // The range must not run backwards.
    [Arguments("CUMIPMT(0.0075, 360, 125000, 1, 361, 0)")] // The range must not exceed the term.
    [Arguments("CUMIPMT(0.0075, 360, 125000, 1, 12, 2)")] // Type is 0 or 1.
    [Arguments("CUMPRINC(0, 360, 125000, 1, 12, 0)")]
    [Arguments("CUMPRINC(0.0075, 360, 125000, 13, 12, 0)")]
    [Arguments("CUMPRINC(0.0075, 360, 125000, 1, 12, 2)")]
    public async Task Cumulative_OutOfRangeArgumentsReturnNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    // Microsoft's FVSCHEDULE example: 1 compounded at 9%, 11% and 10% = 1.09 * 1.11 * 1.1.
    public async Task FvSchedule_CompoundsEachRateInTurn()
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr("FVSCHEDULE(1, {0.09,0.11,0.1})")).IsEqualTo(1.33089d).Within(1e-12);
    }

    [Test]
    public async Task FvSchedule_ReadsTheScheduleFromARange()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 0.09;
            ws.Cell("A2").Value = 0.11;
            ws.Cell("A3").Value = 0.1;
            ws.Cell("B1").FormulaA1 = "FVSCHEDULE(1, A1:A3)";

            await Assert.That((double)ws.Cell("B1").Value).IsEqualTo(1.33089d).Within(1e-12);
        }
    }

    [Test]
    public async Task FvSchedule_SkipsBlanksAndPropagatesErrors()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 0.09;
            // A2 deliberately left blank — a missing rate compounds by nothing.
            ws.Cell("A3").Value = 0.11;
            ws.Cell("B1").FormulaA1 = "FVSCHEDULE(1, A1:A3)";

            ws.Cell("C1").Value = 0.09;
            ws.Cell("C2").FormulaA1 = "1/0";
            ws.Cell("D1").FormulaA1 = "FVSCHEDULE(1, C1:C2)";

            await Assert.That((double)ws.Cell("B1").Value).IsEqualTo(1.09d * 1.11d).Within(1e-12);
            await Assert.That(ws.Cell("D1").Value).IsEqualTo(XLError.DivisionByZero);
        }
    }

    [Test]
    // Microsoft's MIRR example: an initial 120,000 outlay followed by five years of income,
    // financed at 10% and reinvested at 12%.
    [Arguments("MIRR(A1:A6, 0.1, 0.12)", 0.12609413036d)] // All five years.
    [Arguments("MIRR(A1:A4, 0.1, 0.12)", -0.048044655249d)] // The first three years only.
    [Arguments("MIRR(A1:A6, 0.1, 0.14)", 0.13475911059d)] // Reinvesting at 14% instead.
    public async Task Mirr_ReferenceExamplesFromExcelDocumentation(string formula, double expected)
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = -120000;
            ws.Cell("A2").Value = 39000;
            ws.Cell("A3").Value = 30000;
            ws.Cell("A4").Value = 21000;
            ws.Cell("A5").Value = 37000;
            ws.Cell("A6").Value = 46000;
            ws.Cell("C1").FormulaA1 = formula;

            await Assert.That((double)ws.Cell("C1").Value).IsEqualTo(expected).Within(1e-6);
        }
    }

    [Test]
    public async Task Mirr_MatchesIrrWhenBothRatesEqualIt()
    {
        // When money is financed and reinvested at the project's own IRR, the modified rate of
        // return collapses back to it.
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = -1000;
            ws.Cell("A2").Value = 500;
            ws.Cell("A3").Value = 400;
            ws.Cell("A4").Value = 300;
            ws.Cell("B1").FormulaA1 = "IRR(A1:A4)";
            ws.Cell("B2").FormulaA1 = "MIRR(A1:A4, B1, B1)";

            await Assert.That((double)ws.Cell("B2").Value).IsEqualTo((double)ws.Cell("B1").Value).Within(1e-6);
        }
    }

    [Test]
    public async Task Mirr_WithoutBothSignsIsDivisionByZero()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 100;
            ws.Cell("A2").Value = 200;
            ws.Cell("B1").FormulaA1 = "MIRR(A1:A2, 0.1, 0.12)";

            await Assert.That(ws.Cell("B1").Value).IsEqualTo(XLError.DivisionByZero);
        }
    }

    [Test]
    // Microsoft's XNPV/XIRR example: -10,000 on 2008-01-01 followed by four irregular receipts.
    // XNPV discounts each flow by its own actual/365 offset from the first date.
    public async Task XNpv_ReferenceExampleFromExcelDocumentation()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedIrregularSchedule(ws);
            ws.Cell("C1").FormulaA1 = "XNPV(0.09, A1:A5, B1:B5)";

            await Assert.That((double)ws.Cell("C1").Value).IsEqualTo(2086.6476020315d).Within(1e-6);
        }
    }

    [Test]
    public async Task XIrr_ReferenceExampleFromExcelDocumentation()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedIrregularSchedule(ws);
            ws.Cell("C1").FormulaA1 = "XIRR(A1:A5, B1:B5)";

            await Assert.That((double)ws.Cell("C1").Value).IsEqualTo(0.373362535d).Within(1e-7);
        }
    }

    [Test]
    public async Task XIrr_IsTheRateAtWhichXNpvIsZero()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedIrregularSchedule(ws);
            ws.Cell("C1").FormulaA1 = "XIRR(A1:A5, B1:B5)";
            ws.Cell("C2").FormulaA1 = "XNPV(C1, A1:A5, B1:B5)";

            await Assert.That((double)ws.Cell("C2").Value).IsEqualTo(0d).Within(1e-6);
        }
    }

    [Test]
    public async Task XIrr_ConvergesFromAPoorGuess()
    {
        // A guess far from the answer must still land on it — Newton falls back to bisection.
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedIrregularSchedule(ws);
            ws.Cell("C1").FormulaA1 = "XIRR(A1:A5, B1:B5, 5)";

            await Assert.That((double)ws.Cell("C1").Value).IsEqualTo(0.373362535d).Within(1e-6);
        }
    }

    [Test]
    public async Task XNpv_MatchesAnnualNpvOnExactYearBoundaries()
    {
        // Flows a whole number of 365-day years apart discount exactly like periodic ones.
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 100;
            ws.Cell("A2").Value = 100;
            ws.Cell("B1").FormulaA1 = "DATE(2007,1,1)";
            ws.Cell("B2").FormulaA1 = "DATE(2008,1,1)"; // 365 days: 2007 is not a leap year.
            ws.Cell("C1").FormulaA1 = "XNPV(0.1, A1:A2, B1:B2)";

            await Assert.That((double)ws.Cell("C1").Value).IsEqualTo(100d + 100d / 1.1d).Within(1e-9);
        }
    }

    [Test]
    public async Task XIrr_WithoutBothSignsReturnsNumberInvalid()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 100;
            ws.Cell("A2").Value = 200;
            ws.Cell("B1").FormulaA1 = "DATE(2008,1,1)";
            ws.Cell("B2").FormulaA1 = "DATE(2009,1,1)";
            ws.Cell("C1").FormulaA1 = "XIRR(A1:A2, B1:B2)";

            await Assert.That(ws.Cell("C1").Value).IsEqualTo(XLError.NumberInvalid);
        }
    }

    [Test]
    public async Task XNpv_MismatchedValueAndDateCountsReturnNumberInvalid()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedIrregularSchedule(ws);
            ws.Cell("C1").FormulaA1 = "XNPV(0.09, A1:A5, B1:B4)";

            await Assert.That(ws.Cell("C1").Value).IsEqualTo(XLError.NumberInvalid);
        }
    }

    [Test]
    public async Task XNpv_DateBeforeTheFirstOneReturnsNumberInvalid()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedIrregularSchedule(ws);
            ws.Cell("B3").FormulaA1 = "DATE(2007,1,1)"; // Earlier than the schedule's start date.
            ws.Cell("C1").FormulaA1 = "XNPV(0.09, A1:A5, B1:B5)";

            await Assert.That(ws.Cell("C1").Value).IsEqualTo(XLError.NumberInvalid);
        }
    }

    [Test]
    public async Task XNpv_PropagatesErrorsFromTheSchedule()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedIrregularSchedule(ws);
            ws.Cell("A3").FormulaA1 = "1/0";
            ws.Cell("C1").FormulaA1 = "XNPV(0.09, A1:A5, B1:B5)";

            await Assert.That(ws.Cell("C1").Value).IsEqualTo(XLError.DivisionByZero);
        }
    }

    [Test]
    public async Task RangeFunctions_AcceptCellReferencesForTheirScalarArguments()
    {
        // NPV, IRR, MIRR, XNPV and XIRR take a range for one argument and scalars for the rest, so
        // their scalar arguments arrive unreduced — a reference has to be unwrapped rather than
        // rejected as the wrong shape.
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = -1000;
            ws.Cell("A2").Value = 500;
            ws.Cell("A3").Value = 400;
            ws.Cell("A4").Value = 300;
            ws.Cell("D1").Value = 0.1;

            ws.Cell("C1").FormulaA1 = "NPV(D1, A1:A4)";
            ws.Cell("C2").FormulaA1 = "NPV(0.1, A1:A4)";
            ws.Cell("C3").FormulaA1 = "IRR(A1:A4, D1)";
            ws.Cell("C4").FormulaA1 = "MIRR(A1:A4, D1, D1)";
            ws.Cell("C5").FormulaA1 = "MIRR(A1:A4, 0.1, 0.1)";

            await Assert.That((double)ws.Cell("C1").Value).IsEqualTo((double)ws.Cell("C2").Value).Within(1e-12);
            await Assert.That((double)ws.Cell("C3").Value).IsGreaterThan(0d);
            await Assert.That((double)ws.Cell("C4").Value).IsEqualTo((double)ws.Cell("C5").Value).Within(1e-12);
        }
    }

    private static void SeedIrregularSchedule(XLWorksheet ws)
    {
        ws.Cell("A1").Value = -10000;
        ws.Cell("A2").Value = 2750;
        ws.Cell("A3").Value = 4250;
        ws.Cell("A4").Value = 3250;
        ws.Cell("A5").Value = 2750;
        ws.Cell("B1").FormulaA1 = "DATE(2008,1,1)";
        ws.Cell("B2").FormulaA1 = "DATE(2008,3,1)";
        ws.Cell("B3").FormulaA1 = "DATE(2008,10,30)";
        ws.Cell("B4").FormulaA1 = "DATE(2009,2,15)";
        ws.Cell("B5").FormulaA1 = "DATE(2009,4,1)";
    }
}
