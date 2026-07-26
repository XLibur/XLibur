using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// CONVERT, the Bessel functions, the error functions and the bitwise and comparison functions.
/// Expected values are the worked examples from Microsoft's per-function documentation, or exact
/// definitional identities (5,280 feet in a mile, 1,024 bytes in a kibibyte) where the docs give no
/// example.
/// </summary>
[SetCulture("en-US")]
public class EngineeringConvertTests
{
    [Test]
    [Arguments("CONVERT(1, \"lbm\", \"kg\")", 0.45359237d)] // A pound is defined as exactly 453.59237 g.
    [Arguments("CONVERT(68, \"F\", \"C\")", 20d)]
    [Arguments("CONVERT(2.5, \"ft\", \"m\")", 0.762d)]
    [Arguments("CONVERT(CONVERT(100, \"ft\", \"m\"), \"ft\", \"m\")", 9.290304d)]
    public async Task Convert_ReferenceExamplesFromExcelDocumentation(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-12);
    }

    [Test]
    [Arguments("CONVERT(1, \"mi\", \"ft\")", 5280d)]
    [Arguments("CONVERT(1, \"yd\", \"ft\")", 3d)]
    [Arguments("CONVERT(1, \"ft\", \"in\")", 12d)]
    [Arguments("CONVERT(1, \"day\", \"hr\")", 24d)]
    [Arguments("CONVERT(1, \"hr\", \"sec\")", 3600d)]
    [Arguments("CONVERT(1, \"yr\", \"day\")", 365.25d)]
    [Arguments("CONVERT(1, \"atm\", \"mmHg\")", 760d)]
    [Arguments("CONVERT(1, \"atm\", \"Pa\")", 101325d)]
    [Arguments("CONVERT(1, \"gal\", \"l\")", 3.785411784d)]
    [Arguments("CONVERT(1, \"gal\", \"qt\")", 4d)]
    [Arguments("CONVERT(1, \"gal\", \"cup\")", 16d)]
    [Arguments("CONVERT(1, \"tbs\", \"tsp\")", 3d)]
    [Arguments("CONVERT(1, \"lbm\", \"ozm\")", 16d)]
    [Arguments("CONVERT(1, \"stone\", \"lbm\")", 14d)]
    [Arguments("CONVERT(1, \"ton\", \"lbm\")", 2000d)]
    [Arguments("CONVERT(1, \"ha\", \"m2\")", 10000d)]
    [Arguments("CONVERT(1, \"m3\", \"l\")", 1000d)]
    [Arguments("CONVERT(1, \"byte\", \"bit\")", 8d)]
    [Arguments("CONVERT(1, \"kn\", \"m/s\")", 0.5144444444444445d)] // A nautical mile per hour.
    public async Task Convert_DefinitionalIdentitiesWithinAMeasure(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-9);
    }

    [Test]
    [Arguments("CONVERT(1, \"km\", \"m\")", 1000d)]
    [Arguments("CONVERT(1, \"m\", \"cm\")", 100d)]
    [Arguments("CONVERT(1, \"m\", \"mm\")", 1000d)]
    [Arguments("CONVERT(1, \"kg\", \"g\")", 1000d)]
    [Arguments("CONVERT(1, \"Mg\", \"kg\")", 1000d)] // Mega beats milli: the prefix is case sensitive.
    [Arguments("CONVERT(1, \"mg\", \"ug\")", 1000d)]
    [Arguments("CONVERT(1, \"kPa\", \"Pa\")", 1000d)]
    [Arguments("CONVERT(1, \"dam\", \"m\")", 10d)] // "da" for deka wins over "d" for deci.
    [Arguments("CONVERT(1, \"kibyte\", \"byte\")", 1024d)]
    [Arguments("CONVERT(1, \"Mibit\", \"bit\")", 1048576d)]
    [Arguments("CONVERT(1, \"kbyte\", \"byte\")", 1000d)] // The metric prefix is still a thousand.
    public async Task Convert_AppliesPrefixes(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-9);
    }

    [Test]
    [Arguments("CONVERT(0, \"C\", \"F\")", 32d)]
    [Arguments("CONVERT(100, \"C\", \"F\")", 212d)]
    [Arguments("CONVERT(0, \"C\", \"K\")", 273.15d)]
    [Arguments("CONVERT(0, \"K\", \"C\")", -273.15d)]
    [Arguments("CONVERT(0, \"C\", \"Rank\")", 491.67d)]
    [Arguments("CONVERT(100, \"C\", \"Reau\")", 80d)]
    [Arguments("CONVERT(-40, \"C\", \"F\")", -40d)] // The one point where the two scales meet.
    public async Task Convert_TemperatureIsAffineNotProportional(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-9);
    }

    [Test]
    [Arguments("CONVERT(1, \"Pica\", \"in\")", 1d / 72d)] // "Pica" is a point.
    [Arguments("CONVERT(1, \"pica\", \"in\")", 1d / 6d)] // "pica" is six to the inch.
    public async Task Convert_UnitNamesAreCaseSensitive(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-12);
    }

    [Test]
    [Arguments("CONVERT(2.5, \"ft\", \"sec\")")] // Distance is not time.
    [Arguments("CONVERT(1, \"kg\", \"m\")")]
    [Arguments("CONVERT(1, \"zzz\", \"m\")")] // Unknown unit.
    [Arguments("CONVERT(1, \"m\", \"zzz\")")]
    [Arguments("CONVERT(1, \"kft\", \"m\")")] // A foot takes no metric prefix.
    [Arguments("CONVERT(1, \"kibm\", \"m\")")] // Binary prefixes belong to the information units only.
    public async Task Convert_UnknownOrMismatchedUnitsReturnNoValueAvailable(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NoValueAvailable);
    }

    [Test]
    public async Task Convert_RoundTripsBackToTheOriginalValue()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = 12.5;
        ws.Cell("B1").FormulaA1 = "CONVERT(A1, \"mi\", \"km\")";
        ws.Cell("C1").FormulaA1 = "CONVERT(B1, \"km\", \"mi\")";

        await Assert.That((double)ws.Cell("C1").Value).IsEqualTo(12.5d).Within(1e-12);
    }

    [Test]
    // Microsoft's Bessel examples.
    [Arguments("BESSELI(1.5, 1)", 0.981666428d)]
    [Arguments("BESSELJ(1.9, 2)", 0.329925829d)]
    [Arguments("BESSELK(1.5, 1)", 0.277387804d)]
    [Arguments("BESSELY(2.5, 1)", 0.145918138d)]
    public async Task Bessel_ReferenceExamplesFromExcelDocumentation(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-7);
    }

    [Test]
    // Well-known values: J0(0) = 1, and every other kind of order-n function vanishes at the origin.
    [Arguments("BESSELJ(0, 0)", 1d)]
    [Arguments("BESSELJ(0, 1)", 0d)]
    [Arguments("BESSELJ(0, 5)", 0d)]
    [Arguments("BESSELI(0, 0)", 1d)]
    [Arguments("BESSELI(0, 1)", 0d)]
    public async Task Bessel_ValuesAtTheOrigin(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-9);
    }

    [Test]
    public async Task Bessel_SatisfiesItsRecurrenceRelation()
    {
        // J(n-1, x) + J(n+1, x) = 2n/x · J(n, x) holds for every kind, which exercises both the
        // upward and the downward recurrences at once.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = "BESSELJ(3.7, 2) + BESSELJ(3.7, 4) - 2 * 3 / 3.7 * BESSELJ(3.7, 3)";
        ws.Cell("A2").FormulaA1 = "BESSELY(3.7, 2) + BESSELY(3.7, 4) - 2 * 3 / 3.7 * BESSELY(3.7, 3)";

        await Assert.That((double)ws.Cell("A1").Value).IsEqualTo(0d).Within(1e-7);
        await Assert.That((double)ws.Cell("A2").Value).IsEqualTo(0d).Within(1e-7);
    }

    [Test]
    [Arguments("BESSELJ(1.5, -1)")] // The order may not be negative.
    [Arguments("BESSELI(1.5, -1)")]
    [Arguments("BESSELK(1.5, -1)")]
    [Arguments("BESSELY(1.5, -1)")]
    [Arguments("BESSELK(0, 1)")] // K and Y are singular at the origin.
    [Arguments("BESSELK(-1, 1)")]
    [Arguments("BESSELY(0, 1)")]
    [Arguments("BESSELY(-1, 1)")]
    public async Task Bessel_OutOfRangeArgumentsReturnNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    // Microsoft's ERF/ERFC examples, plus values from published error-function tables.
    [Arguments("ERF(0.745)", 0.70792892d)]
    [Arguments("ERF(1)", 0.842700792949715d)]
    [Arguments("ERF(0)", 0d)]
    [Arguments("ERF(-1)", -0.842700792949715d)] // erf is odd.
    [Arguments("ERF.PRECISE(0.745)", 0.70792892d)]
    [Arguments("ERF.PRECISE(1)", 0.842700792949715d)]
    [Arguments("ERFC(1)", 0.157299207050285d)]
    [Arguments("ERFC(0)", 1d)]
    [Arguments("ERFC.PRECISE(1)", 0.157299207050285d)]
    [Arguments("ERFC(-1)", 1.842700792949715d)]
    public async Task Erf_ReferenceExamplesFromExcelDocumentation(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-9);
    }

    [Test]
    public async Task Erf_WithTwoLimitsIntegratesBetweenThem()
    {
        // ERF(a, b) = erf(b) - erf(a), so splitting the interval has to add up.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = "ERF(0, 2)";
        ws.Cell("A2").FormulaA1 = "ERF(2)";
        ws.Cell("A3").FormulaA1 = "ERF(0.5, 2) + ERF(0, 0.5)";

        await Assert.That((double)ws.Cell("A1").Value).IsEqualTo((double)ws.Cell("A2").Value).Within(1e-12);
        await Assert.That((double)ws.Cell("A3").Value).IsEqualTo((double)ws.Cell("A2").Value).Within(1e-12);
    }

    [Test]
    public async Task Erfc_KeepsItsPrecisionInTheFarTail()
    {
        // 1 - erf(6) would round to zero in double precision; taking the tail directly does not.
        var actual = (double)XLWorkbook.EvaluateExpr("ERFC(6)");
        await Assert.That(actual).IsEqualTo(2.15197367124989e-17d).Within(1e-28);
    }

    [Test]
    // Microsoft's DELTA and GESTEP examples.
    [Arguments("DELTA(5, 4)", 0d)]
    [Arguments("DELTA(5, 5)", 1d)]
    [Arguments("DELTA(0.5, 0)", 0d)]
    [Arguments("DELTA(0)", 1d)] // The second number defaults to zero.
    [Arguments("GESTEP(5, 4)", 1d)]
    [Arguments("GESTEP(5, 5)", 1d)] // The step itself passes the test.
    [Arguments("GESTEP(-4, -5)", 1d)]
    [Arguments("GESTEP(-1)", 0d)]
    [Arguments("GESTEP(1)", 1d)]
    public async Task DeltaAndGeStep_ReferenceExamplesFromExcelDocumentation(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected);
    }

    [Test]
    // Microsoft's bitwise examples.
    [Arguments("BITAND(13, 25)", 9d)] // 01101 & 11001.
    [Arguments("BITOR(23, 10)", 31d)] // 10111 | 01010.
    [Arguments("BITXOR(5, 3)", 6d)] // 101 ^ 011.
    [Arguments("BITLSHIFT(4, 2)", 16d)]
    [Arguments("BITRSHIFT(13, 2)", 3d)]
    [Arguments("BITLSHIFT(4, -2)", 1d)] // A negative shift goes the other way.
    [Arguments("BITRSHIFT(4, -2)", 16d)]
    [Arguments("BITLSHIFT(3, 0)", 3d)]
    [Arguments("BITAND(0, 255)", 0d)]
    [Arguments("BITOR(0, 255)", 255d)]
    public async Task Bitwise_ReferenceExamplesFromExcelDocumentation(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected);
    }

    [Test]
    public async Task Bitwise_WorksAcrossTheFull48BitRange()
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr("BITAND(281474976710655, 281474976710655)")).IsEqualTo(281474976710655d);
        await Assert.That((double)XLWorkbook.EvaluateExpr("BITXOR(281474976710655, 281474976710655)")).IsEqualTo(0d);
        await Assert.That((double)XLWorkbook.EvaluateExpr("BITRSHIFT(281474976710655, 47)")).IsEqualTo(1d);
    }

    [Test]
    [Arguments("BITAND(-1, 5)")] // Operands must be non-negative.
    [Arguments("BITAND(1.5, 5)")] // And whole numbers.
    [Arguments("BITAND(281474976710656, 5)")] // 2^48 is one too many bits.
    [Arguments("BITOR(-1, 5)")]
    [Arguments("BITXOR(-1, 5)")]
    [Arguments("BITLSHIFT(1, 54)")] // A shift may not exceed 53 bits.
    [Arguments("BITLSHIFT(1, -54)")]
    [Arguments("BITRSHIFT(1, 54)")]
    [Arguments("BITLSHIFT(281474976710655, 10)")] // The result would exceed 2^53 - 1.
    public async Task Bitwise_OutOfRangeArgumentsReturnNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task Engineering_EvaluatesAgainstWorksheetCells()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = 100;
        ws.Cell("A2").Value = "km";
        ws.Cell("A3").Value = "mi";
        ws.Cell("A4").Value = 13;
        ws.Cell("A5").Value = 25;

        ws.Cell("B1").FormulaA1 = "CONVERT(A1, A2, A3)";
        ws.Cell("B2").FormulaA1 = "BITAND(A4, A5)";
        ws.Cell("B3").FormulaA1 = "ERF(A4 / 13)";

        await Assert.That((double)ws.Cell("B1").Value).IsEqualTo(100000d / 1609.344d).Within(1e-9);
        await Assert.That((double)ws.Cell("B2").Value).IsEqualTo(9d);
        await Assert.That((double)ws.Cell("B3").Value).IsEqualTo(0.842700792949715d).Within(1e-9);
    }
}
