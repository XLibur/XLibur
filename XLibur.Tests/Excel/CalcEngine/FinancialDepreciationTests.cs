using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// SLN, SYD, DB, DDB and VDB. Expected values are the worked examples published in Microsoft's
/// per-function documentation; where a case is not covered there it is derived from the documented
/// formula and the derivation is shown in the comment.
/// </summary>
public class FinancialDepreciationTests
{
    private const double Tolerance = 1e-6;

    [Test]
    // Microsoft's SLN example: an asset costing 30,000 with a 7,500 salvage value over 10 years.
    [Arguments("SLN(30000, 7500, 10)", 2250d)]
    [Arguments("SLN(30000, 7500, 120)", 187.5d)] // The same asset depreciated monthly.
    [Arguments("SLN(1000, 1000, 5)", 0d)] // Nothing to depreciate when salvage equals cost.
    [Arguments("SLN(1000, 1200, 5)", -40d)] // A salvage value above cost appreciates.
    public async Task Sln_SpreadsTheDepreciableAmountEvenly(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(Tolerance);
    }

    [Test]
    public async Task Sln_ZeroLifeIsDivisionByZero()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("SLN(30000, 7500, 0)")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    // Microsoft's SYD example: cost 30,000, salvage 7,500, life 10.
    // (cost - salvage) * (life - per + 1) * 2 / (life * (life + 1)) = 22500 * 10 * 2 / 110.
    [Arguments("SYD(30000, 7500, 10, 1)", 4090.909090909091d)]
    [Arguments("SYD(30000, 7500, 10, 10)", 409.0909090909091d)] // 22500 * 1 * 2 / 110.
    [Arguments("SYD(30000, 7500, 10, 5)", 2454.5454545454545d)] // 22500 * 6 * 2 / 110.
    public async Task Syd_WeightsPeriodsByRemainingLife(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(Tolerance);
    }

    [Test]
    public async Task Syd_ChargesSumToTheDepreciableAmount()
    {
        var total = 0d;
        for (var period = 1; period <= 10; period++)
            total += (double)XLWorkbook.EvaluateExpr($"SYD(30000, 7500, 10, {period})");

        await Assert.That(total).IsEqualTo(22500d).Within(1e-9);
    }

    [Test]
    [Arguments("SYD(30000, 7500, 0, 1)")] // Life must be positive.
    [Arguments("SYD(30000, 7500, 10, 0)")] // Period must be at least one.
    [Arguments("SYD(30000, 7500, 10, 11)")] // Period may not run past the asset's life.
    public async Task Syd_OutOfRangeArgumentsReturnNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    // Microsoft's DB example: cost 1,000,000, salvage 100,000, life 6 years, first year 7 months.
    // The rate is ROUND(1 - (100000/1000000)^(1/6), 3) = 0.319.
    [Arguments("DB(1000000, 100000, 6, 1, 7)", 186083.33333333334d)] // 1000000 * 0.319 * 7/12.
    [Arguments("DB(1000000, 100000, 6, 2, 7)", 259639.41666666669d)] // (1000000 - 186083.33) * 0.319.
    [Arguments("DB(1000000, 100000, 6, 3, 7)", 176814.44275000002d)]
    [Arguments("DB(1000000, 100000, 6, 4, 7)", 120410.63551275003d)]
    [Arguments("DB(1000000, 100000, 6, 5, 7)", 81999.642784418779d)]
    [Arguments("DB(1000000, 100000, 6, 6, 7)", 55841.756736028462d)]
    [Arguments("DB(1000000, 100000, 6, 7, 7)", 15845.098473848071d)] // Stub period: * (12-7)/12.
    public async Task Db_ReferenceExampleFromExcelDocumentation(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-4);
    }

    [Test]
    public async Task Db_MonthDefaultsToTwelve()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("DB(1000000, 100000, 6, 1)"))
            .IsEqualTo(XLWorkbook.EvaluateExpr("DB(1000000, 100000, 6, 1, 12)"));
    }

    [Test]
    public async Task Db_ZeroCostDepreciatesNothing()
    {
        // The rate is undefined when cost is zero, so Excel short-circuits to no depreciation
        // rather than raising the 0/0 that (salvage/cost)^(1/life) would produce.
        await Assert.That((double)XLWorkbook.EvaluateExpr("DB(0, 0, 5, 1)")).IsEqualTo(0d);
    }

    [Test]
    [Arguments("DB(-1000, 100, 5, 1)")] // Negative cost.
    [Arguments("DB(1000, -100, 5, 1)")] // Negative salvage.
    [Arguments("DB(1000, 100, 0, 1)")] // Life must be positive.
    [Arguments("DB(1000, 100, 5, 0)")] // Period must be at least one.
    [Arguments("DB(1000, 100, 5, 1, 0)")] // Month must be 1..12.
    [Arguments("DB(1000, 100, 5, 1, 13)")]
    [Arguments("DB(1000, 100, 5, 6)")] // With a full first year there is no stub period.
    [Arguments("DB(1000, 100, 5, 7, 7)")] // Even a partial first year only adds one stub period.
    public async Task Db_OutOfRangeArgumentsReturnNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    // Microsoft's DDB example: cost 2,400, salvage 300, life 10 years.
    [Arguments("DDB(2400, 300, 3650, 1)", 1.3150684931506849d)] // First day. 2400 - 2400*(1-2/3650).
    [Arguments("DDB(2400, 300, 120, 1, 2)", 40d)] // First month. 2400 * 2/120.
    [Arguments("DDB(2400, 300, 10, 1, 2)", 480d)] // First year. 2400 * 0.2.
    [Arguments("DDB(2400, 300, 10, 2, 1.5)", 306d)] // 2400*0.85 - 2400*0.85^2.
    [Arguments("DDB(2400, 300, 10, 10)", 22.122547200000029d)] // Capped at the salvage value.
    public async Task Ddb_ReferenceExamplesFromExcelDocumentation(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(Tolerance);
    }

    [Test]
    public async Task Ddb_FactorDefaultsToTwo()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("DDB(2400, 300, 10, 3)"))
            .IsEqualTo(XLWorkbook.EvaluateExpr("DDB(2400, 300, 10, 3, 2)"));
    }

    [Test]
    public async Task Ddb_NeverDepreciatesBelowSalvage()
    {
        // Once the book value has reached the salvage value there is nothing left to charge.
        await Assert.That((double)XLWorkbook.EvaluateExpr("DDB(1000, 900, 10, 8)")).IsEqualTo(0d);
    }

    [Test]
    [Arguments("DDB(-2400, 300, 10, 1)")] // Negative cost.
    [Arguments("DDB(2400, -300, 10, 1)")] // Negative salvage.
    [Arguments("DDB(2400, 300, 0, 1)")] // Life must be positive.
    [Arguments("DDB(2400, 300, 10, 0)")] // Period must be positive.
    [Arguments("DDB(2400, 300, 10, 11)")] // Period may not run past the asset's life.
    [Arguments("DDB(2400, 300, 10, 1, 0)")] // Factor must be positive.
    public async Task Ddb_OutOfRangeArgumentsReturnNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    // Microsoft's VDB example: cost 2,400, salvage 300, life 10 years / 120 months / 3,650 days.
    [Arguments("VDB(2400, 300, 3650, 0, 1)", 1.3150684931506849d)] // First day, same as DDB.
    [Arguments("VDB(2400, 300, 120, 0, 1)", 40d)] // First month.
    [Arguments("VDB(2400, 300, 10, 0, 1)", 480d)] // First year.
    // 2400*(1-(59/60)^18) - 2400*(1-(59/60)^6) = 626.5256 - 230.2193.
    [Arguments("VDB(2400, 300, 120, 6, 18)", 396.30605326475d)] // Months 7 through 18; Microsoft prints 396.31.
    [Arguments("VDB(2400, 300, 10, 0, 0.875, 1.5)", 315d)] // 0.875 of the first year at factor 1.5.
    public async Task Vdb_ReferenceExamplesFromExcelDocumentation(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-4);
    }

    [Test]
    public async Task Vdb_SumsToTheSameTotalAsItsParts()
    {
        // Consecutive slices must charge the same as taking the range whole.
        var whole = (double)XLWorkbook.EvaluateExpr("VDB(2400, 300, 10, 0, 6)");
        var first = (double)XLWorkbook.EvaluateExpr("VDB(2400, 300, 10, 0, 3)");
        var second = (double)XLWorkbook.EvaluateExpr("VDB(2400, 300, 10, 3, 6)");

        await Assert.That(first + second).IsEqualTo(whole).Within(1e-9);
    }

    [Test]
    public async Task Vdb_OverTheWholeLifeChargesTheDepreciableAmount()
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr("VDB(2400, 300, 10, 0, 10)")).IsEqualTo(2100d).Within(1e-9);
        await Assert.That((double)XLWorkbook.EvaluateExpr("VDB(2400, 300, 10, 0, 10, 1)")).IsEqualTo(2100d).Within(1e-9);
    }

    [Test]
    public async Task Vdb_NoSwitchStaysOnDecliningBalance()
    {
        // At factor 1 straight-line overtakes declining balance in the third year, so suppressing
        // the switch leaves the asset short of its salvage value: 2400 - 2400*0.9^10 = 1563.1717.
        var switching = (double)XLWorkbook.EvaluateExpr("VDB(2400, 300, 10, 0, 10, 1)");
        var noSwitch = (double)XLWorkbook.EvaluateExpr("VDB(2400, 300, 10, 0, 10, 1, TRUE)");

        await Assert.That(switching).IsEqualTo(2100d).Within(1e-9);
        await Assert.That(noSwitch).IsEqualTo(1563.17174376d).Within(1e-6);
    }

    [Test]
    public async Task Vdb_WholePeriodMatchesDdbWhenDecliningBalanceStillWins()
    {
        // Over the first year declining balance is still the larger charge, so VDB and DDB agree.
        var vdb = (double)XLWorkbook.EvaluateExpr("VDB(2400, 300, 10, 0, 1)");
        var ddb = (double)XLWorkbook.EvaluateExpr("DDB(2400, 300, 10, 1)");

        await Assert.That(vdb).IsEqualTo(ddb).Within(1e-9);
    }

    [Test]
    [Arguments("VDB(2400, 300, 10, -1, 5)")] // Start period may not be negative.
    [Arguments("VDB(2400, 300, 10, 5, 4)")] // End must not precede start.
    [Arguments("VDB(2400, 300, 10, 0, 11)")] // End may not run past the asset's life.
    [Arguments("VDB(-2400, 300, 10, 0, 5)")] // Negative cost.
    [Arguments("VDB(2400, 3000, 10, 0, 5)")] // Salvage above cost.
    [Arguments("VDB(2400, 300, 10, 0, 5, 0)")] // Factor must be positive.
    public async Task Vdb_OutOfRangeArgumentsReturnNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task Depreciation_EvaluatesAgainstWorksheetCells()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = 2400;
        ws.Cell("A2").Value = 300;
        ws.Cell("A3").Value = 10;
        ws.Cell("A4").Value = 1;

        ws.Cell("B1").FormulaA1 = "SLN(A1, A2, A3)";
        ws.Cell("B2").FormulaA1 = "SYD(A1, A2, A3, A4)";
        ws.Cell("B3").FormulaA1 = "DDB(A1, A2, A3, A4)";
        ws.Cell("B4").FormulaA1 = "VDB(A1, A2, A3, 0, A4)";

        await Assert.That((double)ws.Cell("B1").Value).IsEqualTo(210d).Within(Tolerance);
        await Assert.That((double)ws.Cell("B2").Value).IsEqualTo(381.81818181818181d).Within(Tolerance);
        await Assert.That((double)ws.Cell("B3").Value).IsEqualTo(480d).Within(Tolerance);
        await Assert.That((double)ws.Cell("B4").Value).IsEqualTo(480d).Within(Tolerance);
    }
}
