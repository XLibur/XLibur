using System;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// Regression (SLOPE, INTERCEPT, CORREL, RSQ, STEYX, FORECAST, LINEST, LOGEST, TREND, GROWTH) and
/// the descriptive statistics beside it (AVEDEV, HARMEAN, SKEW, SKEW.P, KURT, TRIMMEAN, PROB,
/// COVARIANCE, FREQUENCY). Expected values are computed by hand from the definition — the data sets
/// are small and chosen so the arithmetic comes out exactly — and the working is in the comment.
/// </summary>
[SetCulture("en-US")]
public class RegressionTests
{
    private static XLWorksheet NewSheet(out XLWorkbook wb)
    {
        wb = new XLWorkbook();
        return (XLWorksheet)wb.AddWorksheet("Sheet1");
    }

    /// <summary>x = 9, 7, 5, 3, 1 in column A and y = 10, 6, 1, 5, 3 in column B.</summary>
    private static void SeedScatter(XLWorksheet ws)
    {
        double[] xs = [9, 7, 5, 3, 1];
        double[] ys = [10, 6, 1, 5, 3];
        for (var i = 0; i < xs.Length; i++)
        {
            ws.Cell(i + 1, 1).Value = xs[i];
            ws.Cell(i + 1, 2).Value = ys[i];
        }
    }

    /// <summary>x = 1..5 in column A and the exactly linear y = 2x + 1 in column B.</summary>
    private static void SeedExactLine(XLWorksheet ws)
    {
        for (var row = 1; row <= 5; row++)
        {
            ws.Cell(row, 1).Value = row;
            ws.Cell(row, 2).Value = 2 * row + 1;
        }
    }

    [Test]
    // Both means are 5, so dx = 4, 2, 0, -2, -4 and dy = 5, 1, -4, 0, -2. That gives
    // Sxy = 30, Sxx = 40 and Syy = 46, and every one of these follows from those three.
    [Arguments("CORREL(A1:A5, B1:B5)", 0.6993786061802354d)] // 30 / sqrt(40*46).
    [Arguments("PEARSON(A1:A5, B1:B5)", 0.6993786061802354d)]
    [Arguments("RSQ(A1:A5, B1:B5)", 0.48913043478260876d)] // 900/1840.
    [Arguments("SLOPE(B1:B5, A1:A5)", 0.75d)] // 30/40.
    [Arguments("INTERCEPT(B1:B5, A1:A5)", 1.25d)] // 5 - 0.75*5.
    [Arguments("STEYX(B1:B5, A1:A5)", 2.798809270624444d)] // sqrt((46 - 900/40) / 3).
    [Arguments("COVARIANCE.P(A1:A5, B1:B5)", 6d)] // 30/5.
    [Arguments("COVAR(A1:A5, B1:B5)", 6d)]
    [Arguments("COVARIANCE.S(A1:A5, B1:B5)", 7.5d)] // 30/4.
    [Arguments("FORECAST(9, B1:B5, A1:A5)", 8d)] // 1.25 + 0.75*9.
    [Arguments("FORECAST.LINEAR(9, B1:B5, A1:A5)", 8d)]
    public async Task PairedStatistics_FollowFromTheThreeSums(string formula, double expected)
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedScatter(ws);
            ws.Cell("D1").FormulaA1 = formula;

            await Assert.That((double)ws.Cell("D1").Value).IsEqualTo(expected).Within(1e-12);
        }
    }

    [Test]
    public async Task PairedStatistics_OnAPerfectLine()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedExactLine(ws);
            ws.Cell("D1").FormulaA1 = "SLOPE(B1:B5, A1:A5)";
            ws.Cell("D2").FormulaA1 = "INTERCEPT(B1:B5, A1:A5)";
            ws.Cell("D3").FormulaA1 = "CORREL(A1:A5, B1:B5)";
            ws.Cell("D4").FormulaA1 = "RSQ(A1:A5, B1:B5)";
            ws.Cell("D5").FormulaA1 = "STEYX(B1:B5, A1:A5)";
            ws.Cell("D6").FormulaA1 = "FORECAST(6, B1:B5, A1:A5)";

            await Assert.That((double)ws.Cell("D1").Value).IsEqualTo(2d).Within(1e-12);
            await Assert.That((double)ws.Cell("D2").Value).IsEqualTo(1d).Within(1e-12);
            await Assert.That((double)ws.Cell("D3").Value).IsEqualTo(1d).Within(1e-12);
            await Assert.That((double)ws.Cell("D4").Value).IsEqualTo(1d).Within(1e-12);
            await Assert.That((double)ws.Cell("D5").Value).IsEqualTo(0d).Within(1e-9); // Nothing left over.
            await Assert.That((double)ws.Cell("D6").Value).IsEqualTo(13d).Within(1e-12);
        }
    }

    [Test]
    public async Task PairedStatistics_RelateToEachOther()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedScatter(ws);
            // RSQ is the square of CORREL, and the slope is the covariance over the variance of x.
            ws.Cell("D1").FormulaA1 = "RSQ(A1:A5, B1:B5) - CORREL(A1:A5, B1:B5) ^ 2";
            ws.Cell("D2").FormulaA1 = "SLOPE(B1:B5, A1:A5) - COVARIANCE.S(A1:A5, B1:B5) / VAR.S(A1:A5)";
            // The regression line passes through the point of both means.
            ws.Cell("D3").FormulaA1 = "FORECAST(AVERAGE(A1:A5), B1:B5, A1:A5) - AVERAGE(B1:B5)";

            foreach (var address in new[] { "D1", "D2", "D3" })
                await Assert.That((double)ws.Cell(address).Value).IsEqualTo(0d).Within(1e-12);
        }
    }

    [Test]
    public async Task PairedStatistics_MismatchedRangesReturnNoValueAvailable()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedScatter(ws);
            ws.Cell("D1").FormulaA1 = "CORREL(A1:A5, B1:B4)";
            ws.Cell("D2").FormulaA1 = "SLOPE(B1:B4, A1:A5)";

            await Assert.That(ws.Cell("D1").Value).IsEqualTo(XLError.NoValueAvailable);
            await Assert.That(ws.Cell("D2").Value).IsEqualTo(XLError.NoValueAvailable);
        }
    }

    [Test]
    public async Task PairedStatistics_WithNoSpreadInXAreDivisionByZero()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            for (var row = 1; row <= 3; row++)
            {
                ws.Cell(row, 1).Value = 5; // Every x the same: no line to fit.
                ws.Cell(row, 2).Value = row;
            }

            ws.Cell("D1").FormulaA1 = "SLOPE(B1:B3, A1:A3)";
            ws.Cell("D2").FormulaA1 = "CORREL(A1:A3, B1:B3)";

            await Assert.That(ws.Cell("D1").Value).IsEqualTo(XLError.DivisionByZero);
            await Assert.That(ws.Cell("D2").Value).IsEqualTo(XLError.DivisionByZero);
        }
    }

    [Test]
    // 1, 2, 3, 4, 5 has mean 3 and sample variance 2.5, so the fourth standardised moment is
    // 34/6.25 = 5.44 and KURT = (5*6/(4*3*2))*5.44 - 3*16/(3*2) = 6.8 - 8.
    [Arguments("KURT(1, 2, 3, 4, 5)", -1.2d)]
    [Arguments("SKEW(1, 2, 3, 4, 5)", 0d)] // A symmetric set has no skew.
    [Arguments("SKEW.P(1, 2, 3, 4, 5)", 0d)]
    // 1, 1, 4 has mean 2 and population variance 2, so SKEW.P = (3/sqrt(2))/3 = 1/sqrt(2).
    [Arguments("SKEW.P(1, 1, 4)", 0.70710678118654746d)]
    [Arguments("AVEDEV(1, 2, 3, 4, 5)", 1.2d)] // (2+1+0+1+2)/5.
    [Arguments("HARMEAN(1, 2, 4)", 1.7142857142857142d)] // 3/(1 + 1/2 + 1/4).
    [Arguments("HARMEAN(2, 2, 2)", 2d)] // Equal values give back the value.
    [Arguments("GEOMEAN(1, 4, 16)", 4d)] // The cube root of 64.
    [Arguments("DEVSQ(1, 2, 3, 4, 5)", 10d)] // 4+1+0+1+4.
    public async Task ShapeStatistics_ComputedByHand(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-12);
    }

    [Test]
    public async Task ShapeStatistics_SkewIsSignedByTheLongerTail()
    {
        // A long right tail is a positive skew, and mirroring the data flips the sign.
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "SKEW(1, 1, 1, 5)";
            ws.Cell("A2").FormulaA1 = "SKEW(-1, -1, -1, -5)";

            await Assert.That((double)ws.Cell("A1").Value).IsGreaterThan(0d);
            await Assert.That((double)ws.Cell("A2").Value).IsEqualTo(-(double)ws.Cell("A1").Value).Within(1e-12);
        }
    }

    [Test]
    [Arguments("SKEW(1, 2)")] // Skewness needs at least three values.
    [Arguments("SKEW(3, 3, 3)")] // And some spread.
    [Arguments("KURT(1, 2, 3)")] // Kurtosis needs at least four.
    [Arguments("KURT(3, 3, 3, 3)")]
    [Arguments("SKEW.P(3, 3, 3)")]
    public async Task ShapeStatistics_UndefinedCasesReturnDivisionByZero(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    [Arguments("HARMEAN(1, 0, 4)")] // A reciprocal of zero has no meaning.
    [Arguments("HARMEAN(1, -2, 4)")]
    public async Task HarMean_NonPositiveValuesReturnNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task TrimMean_DiscardsAnEqualCountFromEachEnd()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            double[] values = [4, 5, 6, 7, 2, 3, 4, 5, 1, 2, 3];
            for (var i = 0; i < values.Length; i++)
                ws.Cell(i + 1, 1).Value = values[i];

            // Eleven values at 20% discards floor(1.1) rounded down to an even 2 — one from each
            // end, so the 1 and the 7 go and 34/9 is left.
            ws.Cell("C1").FormulaA1 = "TRIMMEAN(A1:A11, 0.2)";
            ws.Cell("C2").FormulaA1 = "TRIMMEAN(A1:A11, 0)"; // Nothing trimmed is just the mean.
            ws.Cell("C3").FormulaA1 = "AVERAGE(A1:A11)";

            await Assert.That((double)ws.Cell("C1").Value).IsEqualTo(34d / 9d).Within(1e-12);
            await Assert.That((double)ws.Cell("C2").Value).IsEqualTo((double)ws.Cell("C3").Value).Within(1e-12);
        }
    }

    [Test]
    [Arguments("TRIMMEAN(A1:A11, -0.1)")]
    [Arguments("TRIMMEAN(A1:A11, 1)")] // Trimming everything leaves nothing to average.
    public async Task TrimMean_OutOfRangePercentReturnsNumberInvalid(string formula)
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            for (var row = 1; row <= 11; row++)
                ws.Cell(row, 1).Value = row;

            ws.Cell("C1").FormulaA1 = formula;
            await Assert.That(ws.Cell("C1").Value).IsEqualTo(XLError.NumberInvalid);
        }
    }

    [Test]
    public async Task Prob_AddsUpTheProbabilitiesInRange()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            double[] outcomes = [0, 1, 2, 3];
            double[] probabilities = [0.2, 0.3, 0.1, 0.4];
            for (var i = 0; i < outcomes.Length; i++)
            {
                ws.Cell(i + 1, 1).Value = outcomes[i];
                ws.Cell(i + 1, 2).Value = probabilities[i];
            }

            ws.Cell("D1").FormulaA1 = "PROB(A1:A4, B1:B4, 2)"; // A single outcome.
            ws.Cell("D2").FormulaA1 = "PROB(A1:A4, B1:B4, 1, 3)";
            ws.Cell("D3").FormulaA1 = "PROB(A1:A4, B1:B4, 0, 3)"; // The whole distribution.
            ws.Cell("D4").FormulaA1 = "PROB(A1:A4, B1:B4, 5, 6)"; // Nothing in range.

            await Assert.That((double)ws.Cell("D1").Value).IsEqualTo(0.1d).Within(1e-12);
            await Assert.That((double)ws.Cell("D2").Value).IsEqualTo(0.8d).Within(1e-12);
            await Assert.That((double)ws.Cell("D3").Value).IsEqualTo(1d).Within(1e-12);
            await Assert.That((double)ws.Cell("D4").Value).IsEqualTo(0d).Within(1e-12);
        }
    }

    [Test]
    public async Task Prob_ProbabilitiesThatDoNotDescribeADistributionReturnNumberInvalid()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 2;
            ws.Cell("B1").Value = 0.3;
            ws.Cell("B2").Value = 0.3; // Sums to 0.6, not 1.
            ws.Cell("D1").FormulaA1 = "PROB(A1:A2, B1:B2, 1, 2)";

            ws.Cell("C1").Value = 1.5; // Not a probability at all.
            ws.Cell("C2").Value = -0.5;
            ws.Cell("D2").FormulaA1 = "PROB(A1:A2, C1:C2, 1, 2)";

            await Assert.That(ws.Cell("D1").Value).IsEqualTo(XLError.NumberInvalid);
            await Assert.That(ws.Cell("D2").Value).IsEqualTo(XLError.NumberInvalid);
        }
    }

    [Test]
    public async Task Frequency_CountsIntoBinsAndTheOverflow()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            double[] scores = [79, 85, 78, 85, 50, 81, 95, 88, 97];
            for (var i = 0; i < scores.Length; i++)
                ws.Cell(i + 1, 1).Value = scores[i];

            double[] bins = [70, 79, 89];
            for (var i = 0; i < bins.Length; i++)
                ws.Cell(i + 1, 2).Value = bins[i];

            ws.Range("D1:D4").FormulaArrayA1 = "FREQUENCY(A1:A9, B1:B3)";

            // 50 alone is at or below 70; 79 and 78 fall in the next bin; 85, 85, 81 and 88 in the
            // next; 95 and 97 overflow past the last.
            await Assert.That((double)ws.Cell("D1").Value).IsEqualTo(1d);
            await Assert.That((double)ws.Cell("D2").Value).IsEqualTo(2d);
            await Assert.That((double)ws.Cell("D3").Value).IsEqualTo(4d);
            await Assert.That((double)ws.Cell("D4").Value).IsEqualTo(2d);
        }
    }

    [Test]
    public async Task Frequency_SpillsAndTotalsToTheDataCount()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            for (var row = 1; row <= 10; row++)
                ws.Cell(row, 1).Value = row;

            ws.Cell("B1").Value = 3;
            ws.Cell("B2").Value = 7;
            ws.Cell("D1").SetDynamicFormulaA1("FREQUENCY(A1:A10, B1:B2)");
            ws.Cell("F1").FormulaA1 = "SUM(D1:D3)";

            await Assert.That((double)ws.Cell("D1").Value).IsEqualTo(3d); // 1, 2, 3.
            await Assert.That((double)ws.Cell("D2").Value).IsEqualTo(4d); // 4 through 7.
            await Assert.That((double)ws.Cell("D3").Value).IsEqualTo(3d); // 8, 9, 10.
            await Assert.That((double)ws.Cell("F1").Value).IsEqualTo(10d); // Every value is counted once.
        }
    }

    [Test]
    public async Task Linest_RecoversTheLineItWasGiven()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedExactLine(ws);
            ws.Range("D1:E1").FormulaArrayA1 = "LINEST(B1:B5, A1:A5)";
            // With no x given the predictor is the position, which here is the same 1..5.
            ws.Range("D2:E2").FormulaArrayA1 = "LINEST(B1:B5)";

            await Assert.That((double)ws.Cell("D1").Value).IsEqualTo(2d).Within(1e-9); // Slope.
            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(1d).Within(1e-9); // Intercept.
            await Assert.That((double)ws.Cell("D2").Value).IsEqualTo(2d).Within(1e-9);
            await Assert.That((double)ws.Cell("E2").Value).IsEqualTo(1d).Within(1e-9);
        }
    }

    [Test]
    public async Task Linest_AgreesWithSlopeAndIntercept()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedScatter(ws);
            ws.Range("D1:E1").FormulaArrayA1 = "LINEST(B1:B5, A1:A5)";
            ws.Cell("D3").FormulaA1 = "SLOPE(B1:B5, A1:A5)";
            ws.Cell("E3").FormulaA1 = "INTERCEPT(B1:B5, A1:A5)";

            await Assert.That((double)ws.Cell("D1").Value).IsEqualTo((double)ws.Cell("D3").Value).Within(1e-9);
            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo((double)ws.Cell("E3").Value).Within(1e-9);
        }
    }

    [Test]
    public async Task Linest_WithStatisticsReportsTheFitQuality()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedScatter(ws);
            ws.Range("D1:E5").FormulaArrayA1 = "LINEST(B1:B5, A1:A5, TRUE, TRUE)";
            ws.Cell("G1").FormulaA1 = "RSQ(A1:A5, B1:B5)";
            ws.Cell("G2").FormulaA1 = "STEYX(B1:B5, A1:A5)";

            await Assert.That((double)ws.Cell("D1").Value).IsEqualTo(0.75d).Within(1e-9); // Slope.
            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(1.25d).Within(1e-9); // Intercept.
            await Assert.That((double)ws.Cell("D3").Value).IsEqualTo((double)ws.Cell("G1").Value).Within(1e-9);
            await Assert.That((double)ws.Cell("E3").Value).IsEqualTo((double)ws.Cell("G2").Value).Within(1e-9);
            await Assert.That((double)ws.Cell("E4").Value).IsEqualTo(3d); // n - 2 degrees of freedom.
            // The regression and residual sums of squares add up to the total spread in y.
            await Assert.That((double)ws.Cell("D5").Value + (double)ws.Cell("E5").Value).IsEqualTo(46d).Within(1e-9);
        }
    }

    [Test]
    public async Task Linest_WithoutAConstantPinsTheLineThroughTheOrigin()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            for (var row = 1; row <= 4; row++)
            {
                ws.Cell(row, 1).Value = row;
                ws.Cell(row, 2).Value = 3 * row; // Exactly y = 3x, no intercept.
            }

            ws.Range("D1:E1").FormulaArrayA1 = "LINEST(B1:B4, A1:A4, FALSE)";

            await Assert.That((double)ws.Cell("D1").Value).IsEqualTo(3d).Within(1e-9);
            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(0d).Within(1e-12);
        }
    }

    [Test]
    public async Task Linest_FitsSeveralPredictorsAtOnce()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            // y = 5 + 2*x1 + 3*x2, exactly.
            double[,] rows = { { 1, 1 }, { 2, 1 }, { 1, 2 }, { 3, 2 }, { 2, 3 } };
            for (var i = 0; i < 5; i++)
            {
                ws.Cell(i + 1, 1).Value = rows[i, 0];
                ws.Cell(i + 1, 2).Value = rows[i, 1];
                ws.Cell(i + 1, 3).Value = 5 + 2 * rows[i, 0] + 3 * rows[i, 1];
            }

            ws.Range("E1:G1").FormulaArrayA1 = "LINEST(C1:C5, A1:B5)";

            // Excel reports the coefficients last predictor first, then the intercept.
            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(3d).Within(1e-8);
            await Assert.That((double)ws.Cell("F1").Value).IsEqualTo(2d).Within(1e-8);
            await Assert.That((double)ws.Cell("G1").Value).IsEqualTo(5d).Within(1e-8);
        }
    }

    [Test]
    public async Task Trend_PredictsAlongTheFittedLine()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedExactLine(ws);
            ws.Range("D1:D5").FormulaArrayA1 = "TREND(B1:B5, A1:A5)";

            ws.Cell("F1").Value = 6;
            ws.Cell("F2").Value = 10;
            ws.Range("G1:G2").FormulaArrayA1 = "TREND(B1:B5, A1:A5, F1:F2)";

            // Fitted onto its own data a perfect line reproduces it.
            for (var row = 1; row <= 5; row++)
                await Assert.That((double)ws.Cell(row, 4).Value).IsEqualTo(2d * row + 1).Within(1e-9);

            await Assert.That((double)ws.Cell("G1").Value).IsEqualTo(13d).Within(1e-9);
            await Assert.That((double)ws.Cell("G2").Value).IsEqualTo(21d).Within(1e-9);
        }
    }

    [Test]
    public async Task Logest_AndGrowth_RecoverAnExponentialCurve()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            // y = 2 * 3^x for x = 1..4.
            for (var row = 1; row <= 4; row++)
            {
                ws.Cell(row, 1).Value = row;
                ws.Cell(row, 2).Value = 2 * Math.Pow(3, row);
            }

            ws.Range("D1:E1").FormulaArrayA1 = "LOGEST(B1:B4, A1:A4)";
            ws.Cell("F1").Value = 5;
            ws.Range("G1:G1").FormulaArrayA1 = "GROWTH(B1:B4, A1:A4, F1)";

            await Assert.That((double)ws.Cell("D1").Value).IsEqualTo(3d).Within(1e-9); // The base.
            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(2d).Within(1e-9); // The factor.
            await Assert.That((double)ws.Cell("G1").Value).IsEqualTo(486d).Within(1e-6); // 2 * 3^5.
        }
    }

    [Test]
    public async Task Logest_RefusesANonPositiveY()
    {
        // Fitting y = b*m^x means fitting the logarithm of y, which a zero or negative value has not.
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 2;
            ws.Cell("B1").Value = 5;
            ws.Cell("B2").Value = 0;
            ws.Cell("D1").FormulaA1 = "LOGEST(B1:B2, A1:A2)";

            await Assert.That(ws.Cell("D1").Value).IsEqualTo(XLError.NumberInvalid);
        }
    }

    [Test]
    public async Task LeastSquares_MismatchedRangesReturnCellReference()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedScatter(ws);
            ws.Cell("D1").FormulaA1 = "LINEST(B1:B5, A1:A4)";
            ws.Cell("D2").FormulaA1 = "TREND(B1:B5, A1:A4)";

            await Assert.That(ws.Cell("D1").Value).IsEqualTo(XLError.CellReference);
            await Assert.That(ws.Cell("D2").Value).IsEqualTo(XLError.CellReference);
        }
    }

    [Test]
    public async Task LeastSquares_SpillIntoTheGrid()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedExactLine(ws);
            ws.Cell("D1").SetDynamicFormulaA1("LINEST(B1:B5, A1:A5)");
            ws.Cell("F1").SetDynamicFormulaA1("TREND(B1:B5, A1:A5)");

            // The anchor is read first: until it has been evaluated the spill footprint is unknown,
            // so the cells it will fill still read as blank.
            await Assert.That((double)ws.Cell("D1").Value).IsEqualTo(2d).Within(1e-9);
            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(1d).Within(1e-9);
            await Assert.That((double)ws.Cell("F1").Value).IsEqualTo(3d).Within(1e-9);
            await Assert.That((double)ws.Cell("F5").Value).IsEqualTo(11d).Within(1e-9);
        }
    }

    [Test]
    public async Task LeastSquares_WorkAcrossARowAsWellAsDownAColumn()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            for (var column = 1; column <= 5; column++)
            {
                ws.Cell(1, column).Value = column;
                ws.Cell(2, column).Value = 2 * column + 1;
            }

            ws.Range("A4:B4").FormulaArrayA1 = "LINEST(A2:E2, A1:E1)";
            ws.Range("A5:E5").FormulaArrayA1 = "TREND(A2:E2, A1:E1)";

            await Assert.That((double)ws.Cell("A4").Value).IsEqualTo(2d).Within(1e-9);
            await Assert.That((double)ws.Cell("B4").Value).IsEqualTo(1d).Within(1e-9);
            for (var column = 1; column <= 5; column++)
                await Assert.That((double)ws.Cell(5, column).Value).IsEqualTo(2d * column + 1).Within(1e-9);
        }
    }
}
