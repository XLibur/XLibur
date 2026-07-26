using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// Rate conversion (EFFECT, NOMINAL, RRI, PDURATION), fractional dollar notation (DOLLARDE,
/// DOLLARFR) and the discount-security functions (TBILLEQ, TBILLPRICE, TBILLYIELD, DISC, INTRATE,
/// RECEIVED). Expected values come from the worked examples in Microsoft's per-function
/// documentation; each comment shows the documented formula applied to the arguments.
/// </summary>
public class FinancialSecuritiesTests
{
    [Test]
    // (1 + 0.0525/4)^4 - 1 = 1.013125^4 - 1.
    [Arguments("EFFECT(0.0525, 4)", 0.053542667370758d)]
    [Arguments("EFFECT(0.1, 1)", 0.1d)] // Compounding once a year changes nothing.
    [Arguments("EFFECT(0.1, 2)", 0.1025d)] // 1.05^2 - 1.
    public async Task Effect_CompoundsTheNominalRate(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-9);
    }

    [Test]
    // ((1 + 0.053543)^(1/4) - 1) * 4.
    [Arguments("NOMINAL(0.053543, 4)", 0.0525003198683561d)]
    [Arguments("NOMINAL(0.1, 1)", 0.1d)]
    [Arguments("NOMINAL(0.1025, 2)", 0.1d)] // The inverse of EFFECT(0.1, 2).
    public async Task Nominal_UndoesTheCompounding(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-9);
    }

    [Test]
    public async Task NominalAndEffect_RoundTrip()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = "EFFECT(0.0525, 12)";
        ws.Cell("A2").FormulaA1 = "NOMINAL(A1, 12)";

        await Assert.That((double)ws.Cell("A2").Value).IsEqualTo(0.0525d).Within(1e-12);
    }

    [Test]
    [Arguments("EFFECT(0, 4)")] // The rate must be positive.
    [Arguments("EFFECT(-0.05, 4)")]
    [Arguments("EFFECT(0.05, 0)")] // There must be at least one compounding period.
    [Arguments("NOMINAL(0, 4)")]
    [Arguments("NOMINAL(0.05, 0.5)")] // Truncated to zero periods.
    public async Task EffectAndNominal_OutOfRangeArgumentsReturnNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    // (fv/pv)^(1/nper) - 1 = 1.1^(1/96) - 1.
    [Arguments("RRI(96, 10000, 11000)", 0.000993307372823d)]
    [Arguments("RRI(1, 100, 110)", 0.1d)] // A single period is just the growth rate.
    [Arguments("RRI(2, 100, 121)", 0.1d)]
    public async Task Rri_ReturnsTheEquivalentPeriodicRate(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-9);
    }

    [Test]
    // (LN(fv) - LN(pv)) / LN(1 + rate) = LN(1.1) / LN(1.025).
    [Arguments("PDURATION(0.025, 2000, 2200)", 3.859866162622655d)] // Microsoft prints 3.859866.
    [Arguments("PDURATION(0.1, 100, 121)", 2d)] // 100 grows to 121 in exactly two 10% periods.
    public async Task PDuration_ReturnsThePeriodsNeededToReachAValue(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-9);
    }

    [Test]
    [Arguments("RRI(0, 10000, 11000)")] // Periods must be positive.
    [Arguments("RRI(96, 0, 11000)")] // Present value must be positive.
    [Arguments("RRI(96, 10000, -1)")] // Future value may not be negative.
    [Arguments("PDURATION(0, 2000, 2200)")] // The rate must be positive.
    [Arguments("PDURATION(0.025, 0, 2200)")]
    [Arguments("PDURATION(0.025, 2000, 0)")]
    public async Task RriAndPDuration_OutOfRangeArgumentsReturnNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    // The digits after the point are sixteenths (or thirty-seconds), shifted by the two decimal
    // places the denominator occupies: 1 + 0.02 * 100 / 16 = 1.125.
    [Arguments("DOLLARDE(1.02, 16)", 1.125d)]
    [Arguments("DOLLARDE(1.1, 32)", 1.3125d)]
    [Arguments("DOLLARDE(-1.02, 16)", -1.125d)] // The sign carries through both parts.
    public async Task DollarDe_ConvertsFractionalNotationToDecimal(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-12);
    }

    [Test]
    [Arguments("DOLLARFR(1.125, 16)", 1.02d)]
    [Arguments("DOLLARFR(1.3125, 32)", 1.1d)]
    [Arguments("DOLLARFR(-1.125, 16)", -1.02d)]
    public async Task DollarFr_ConvertsDecimalToFractionalNotation(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-12);
    }

    [Test]
    public async Task DollarDeAndDollarFr_AreInverses()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = "DOLLARDE(1.09, 16)";
        ws.Cell("A2").FormulaA1 = "DOLLARFR(A1, 16)";

        await Assert.That((double)ws.Cell("A1").Value).IsEqualTo(1.5625d).Within(1e-12); // 1 + 9/16.
        await Assert.That((double)ws.Cell("A2").Value).IsEqualTo(1.09d).Within(1e-12);
    }

    [Test]
    [Arguments("DOLLARDE(1.02, 0)")]
    [Arguments("DOLLARFR(1.125, 0)")]
    [Arguments("DOLLARDE(1.02, 0.5)")] // Truncated to a zero denominator.
    public async Task DollarConversions_ZeroFractionIsDivisionByZero(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    [Arguments("DOLLARDE(1.02, -1)")]
    [Arguments("DOLLARFR(1.125, -1)")]
    public async Task DollarConversions_NegativeFractionReturnsNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    // Microsoft's Treasury bill examples: settlement 2008-03-31, maturity 2008-06-01 (62 days).
    // TBILLEQ = (365 * 0.0914) / (360 - 0.0914 * 62) = 33.361 / 354.3332.
    [Arguments("TBILLEQ(DATE(2008,3,31), DATE(2008,6,1), 0.0914)", 0.0941514937d)]
    // TBILLPRICE = 100 * (1 - 0.09 * 62/360).
    [Arguments("TBILLPRICE(DATE(2008,3,31), DATE(2008,6,1), 0.09)", 98.45d)]
    // TBILLYIELD = (100 - 98.45)/98.45 * 360/62.
    [Arguments("TBILLYIELD(DATE(2008,3,31), DATE(2008,6,1), 98.45)", 0.0914169629d)]
    public async Task TBill_ReferenceExamplesFromExcelDocumentation(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-9);
    }

    [Test]
    public async Task TBillYield_ExceedsTheDiscountRateOfTheSamePrice()
    {
        // The discount is quoted on the face value but earned on the lower purchase price, so the
        // yield a buyer actually gets is the larger of the two.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = "TBILLPRICE(DATE(2008,3,31), DATE(2008,6,1), 0.09)";
        ws.Cell("A2").FormulaA1 = "TBILLYIELD(DATE(2008,3,31), DATE(2008,6,1), A1)";

        await Assert.That((double)ws.Cell("A1").Value).IsEqualTo(98.45d).Within(1e-12);
        await Assert.That((double)ws.Cell("A2").Value).IsGreaterThan(0.09d);
    }

    [Test]
    [Arguments("TBILLEQ(DATE(2008,6,1), DATE(2008,3,31), 0.0914)")] // Maturity before settlement.
    [Arguments("TBILLEQ(DATE(2008,3,31), DATE(2008,3,31), 0.0914)")] // Same day.
    [Arguments("TBILLEQ(DATE(2008,3,31), DATE(2010,6,1), 0.0914)")] // More than a year to maturity.
    [Arguments("TBILLEQ(DATE(2008,3,31), DATE(2008,6,1), 0)")] // The discount must be positive.
    [Arguments("TBILLPRICE(DATE(2008,3,31), DATE(2010,6,1), 0.09)")]
    [Arguments("TBILLPRICE(DATE(2008,3,31), DATE(2008,6,1), -0.09)")]
    [Arguments("TBILLYIELD(DATE(2008,3,31), DATE(2010,6,1), 98.45)")]
    [Arguments("TBILLYIELD(DATE(2008,3,31), DATE(2008,6,1), 0)")] // The price must be positive.
    public async Task TBill_OutOfRangeArgumentsReturnNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    // Settlement 2007-01-07, maturity 2007-06-15. Actual days 159; 30/360 days 158.
    // DISC = (100 - 97.975)/100 / yearFraction, and (100 - 97.975)/100 = 0.02025.
    [Arguments("DISC(DATE(2007,1,7), DATE(2007,6,15), 97.975, 100, 2)", 0.045849056603773585d)] // 0.02025 / (159/360).
    [Arguments("DISC(DATE(2007,1,7), DATE(2007,6,15), 97.975, 100, 0)", 0.046139240506329116d)] // 0.02025 / (158/360).
    [Arguments("DISC(DATE(2007,1,7), DATE(2007,6,15), 97.975, 100, 3)", 0.046485849056603776d)] // 0.02025 / (159/365).
    public async Task Disc_ReferenceExampleFromExcelDocumentation(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-9);
    }

    [Test]
    public async Task Disc_BasisDefaultsToUsThirtyThreeSixty()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("DISC(DATE(2007,1,7), DATE(2007,6,15), 97.975, 100)"))
            .IsEqualTo(XLWorkbook.EvaluateExpr("DISC(DATE(2007,1,7), DATE(2007,6,15), 97.975, 100, 0)"));
    }

    [Test]
    // Microsoft's INTRATE example: settlement 2008-02-15, maturity 2008-05-15, basis 2 gives a year
    // fraction of exactly 90/360. (1014420 - 1000000)/1000000 / 0.25.
    public async Task IntRate_ReferenceExampleFromExcelDocumentation()
    {
        var actual = (double)XLWorkbook.EvaluateExpr("INTRATE(DATE(2008,2,15), DATE(2008,5,15), 1000000, 1014420, 2)");
        await Assert.That(actual).IsEqualTo(0.05768d).Within(1e-12);
    }

    [Test]
    // Microsoft's RECEIVED example, same dates and basis: 1000000 / (1 - 0.0575 * 0.25), and
    // 1 - 0.014375 = 1577/1600, so the result is 1000000 * 1600/1577 = 1,014,584.65.
    public async Task Received_ReferenceExampleFromExcelDocumentation()
    {
        var actual = (double)XLWorkbook.EvaluateExpr("RECEIVED(DATE(2008,2,15), DATE(2008,5,15), 1000000, 0.0575, 2)");
        await Assert.That(actual).IsEqualTo(1_600_000_000d / 1577d).Within(1e-6);
    }

    [Test]
    public async Task IntRateAndReceived_DescribeTheSameSecurity()
    {
        // Discounting the proceeds back at the quoted discount rate has to return the investment,
        // and growing the investment at the derived interest rate has to return the proceeds.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = "RECEIVED(DATE(2008,2,15), DATE(2008,5,15), 1000000, 0.0575, 2)";
        ws.Cell("A2").FormulaA1 = "INTRATE(DATE(2008,2,15), DATE(2008,5,15), 1000000, A1, 2)";

        var received = (double)ws.Cell("A1").Value;
        var rate = (double)ws.Cell("A2").Value;

        await Assert.That(received * (1 - 0.0575 * 0.25)).IsEqualTo(1000000d).Within(1e-6);
        await Assert.That(1000000 * (1 + rate * 0.25)).IsEqualTo(received).Within(1e-6);
    }

    [Test]
    [Arguments("DISC(DATE(2007,6,15), DATE(2007,1,7), 97.975, 100, 2)")] // Maturity before settlement.
    [Arguments("DISC(DATE(2007,1,7), DATE(2007,6,15), 0, 100, 2)")] // The price must be positive.
    [Arguments("DISC(DATE(2007,1,7), DATE(2007,6,15), 97.975, 0, 2)")] // The redemption must be positive.
    [Arguments("DISC(DATE(2007,1,7), DATE(2007,6,15), 97.975, 100, 5)")] // Only bases 0..4 exist.
    [Arguments("INTRATE(DATE(2008,5,15), DATE(2008,2,15), 1000000, 1014420, 2)")]
    [Arguments("INTRATE(DATE(2008,2,15), DATE(2008,5,15), 0, 1014420, 2)")]
    [Arguments("RECEIVED(DATE(2008,5,15), DATE(2008,2,15), 1000000, 0.0575, 2)")]
    [Arguments("RECEIVED(DATE(2008,2,15), DATE(2008,5,15), 1000000, 0, 2)")] // The discount must be positive.
    [Arguments("RECEIVED(DATE(2008,2,15), DATE(2008,5,15), 1000000, 5, 2)")] // A discount that wipes out the proceeds.
    public async Task Securities_OutOfRangeArgumentsReturnNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task Securities_EvaluateAgainstWorksheetCells()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = "DATE(2008,2,15)";
        ws.Cell("A2").FormulaA1 = "DATE(2008,5,15)";
        ws.Cell("A3").Value = 1000000;
        ws.Cell("A4").Value = 1014420;
        ws.Cell("A5").Value = 2;

        ws.Cell("B1").FormulaA1 = "INTRATE(A1, A2, A3, A4, A5)";
        ws.Cell("B2").FormulaA1 = "EFFECT(B1, 4)";

        await Assert.That((double)ws.Cell("B1").Value).IsEqualTo(0.05768d).Within(1e-12);
        await Assert.That((double)ws.Cell("B2").Value).IsGreaterThan(0.05768d);
    }
}
