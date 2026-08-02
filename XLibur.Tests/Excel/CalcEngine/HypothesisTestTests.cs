using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// The hypothesis tests (CHISQ.TEST, F.TEST, T.TEST, Z.TEST) and the rank functions added alongside
/// them (RANK.AVG, MODE.MULT, PERCENTILE.EXC, QUARTILE.EXC, PERCENTRANK). The tests are checked
/// against the distribution they are defined in terms of, computed independently in the same
/// worksheet — which is exactly the identity Microsoft's documentation states for each of them.
/// </summary>
[SetCulture("en-US")]
public class HypothesisTestTests
{
    private static IXLWorksheet NewSheet(out XLWorkbook wb)
    {
        wb = new XLWorkbook();
        return wb.AddWorksheet("Sheet1");
    }

    private static void SeedTwoSamples(IXLWorksheet ws)
    {
        double[] first = [3, 4, 5, 8, 9, 1, 2, 4, 5];
        double[] second = [6, 19, 3, 2, 14, 4, 5, 17, 1];
        for (var i = 0; i < first.Length; i++)
        {
            ws.Cell(i + 1, 1).Value = first[i];
            ws.Cell(i + 1, 2).Value = second[i];
        }
    }

    [Test]
    public async Task ChiSqTest_IsTheRightTailOfTheStatisticItComputes()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            // A two-by-three table, so two degrees of freedom.
            double[,] actual = { { 58, 11, 10 }, { 35, 25, 23 } };
            double[,] expected = { { 45.35, 17.56, 16.09 }, { 47.65, 18.44, 16.91 } };
            for (var row = 0; row < 2; row++)
            {
                for (var column = 0; column < 3; column++)
                {
                    ws.Cell(row + 1, column + 1).Value = actual[row, column];
                    ws.Cell(row + 1, column + 5).Value = expected[row, column];
                }
            }

            ws.Cell("A5").FormulaA1 = "CHISQ.TEST(A1:C2, E1:G2)";
            ws.Cell("A6").FormulaA1 = "SUMPRODUCT((A1:C2 - E1:G2) ^ 2 / E1:G2)";
            ws.Cell("A7").FormulaA1 = "CHISQ.DIST.RT(A6, 2)";
            ws.Cell("A8").FormulaA1 = "CHITEST(A1:C2, E1:G2) - CHISQ.TEST(A1:C2, E1:G2)";

            await Assert.That((double)ws.Cell("A5").Value).IsEqualTo((double)ws.Cell("A7").Value).Within(1e-12);
            await Assert.That((double)ws.Cell("A5").Value).IsGreaterThan(0d);
            await Assert.That((double)ws.Cell("A5").Value).IsLessThan(1d);
            await Assert.That((double)ws.Cell("A8").Value).IsEqualTo(0d);
        }
    }

    [Test]
    public async Task ChiSqTest_UsesOneLessThanTheLengthForAVector()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            for (var row = 1; row <= 4; row++)
            {
                ws.Cell(row, 1).Value = row * 10;
                ws.Cell(row, 2).Value = 25;
            }

            ws.Cell("D1").FormulaA1 = "CHISQ.TEST(A1:A4, B1:B4)";
            ws.Cell("D2").FormulaA1 = "SUMPRODUCT((A1:A4 - B1:B4) ^ 2 / B1:B4)";
            ws.Cell("D3").FormulaA1 = "CHISQ.DIST.RT(D2, 3)";

            await Assert.That((double)ws.Cell("D1").Value).IsEqualTo((double)ws.Cell("D3").Value).Within(1e-12);
        }
    }

    [Test]
    public async Task ChiSqTest_MismatchedShapesAndZeroExpectedValues()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 10;
            ws.Cell("A2").Value = 20;
            ws.Cell("A3").Value = 30;
            ws.Cell("B1").Value = 15;
            ws.Cell("B2").Value = 15;
            ws.Cell("B3").Value = 0;

            ws.Cell("D1").FormulaA1 = "CHISQ.TEST(A1:A3, B1:B2)"; // Different shapes.
            ws.Cell("D2").FormulaA1 = "CHISQ.TEST(A1:A3, B1:B3)"; // A zero expected frequency.

            await Assert.That(ws.Cell("D1").Value).IsEqualTo(XLError.NoValueAvailable);
            await Assert.That(ws.Cell("D2").Value).IsEqualTo(XLError.DivisionByZero);
        }
    }

    [Test]
    public async Task FTest_IsTwiceTheOneTailedProbabilityOfTheVarianceRatio()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedTwoSamples(ws);

            ws.Cell("D1").FormulaA1 = "F.TEST(A1:A9, B1:B9)";
            ws.Cell("D2").FormulaA1 = "VAR.S(B1:B9) / VAR.S(A1:A9)"; // B is the more variable sample.
            ws.Cell("D3").FormulaA1 = "2 * F.DIST.RT(D2, 8, 8)";
            ws.Cell("D4").FormulaA1 = "FTEST(A1:A9, B1:B9) - F.TEST(A1:A9, B1:B9)";
            // The test does not care which sample is given first.
            ws.Cell("D5").FormulaA1 = "F.TEST(B1:B9, A1:A9) - F.TEST(A1:A9, B1:B9)";

            await Assert.That((double)ws.Cell("D1").Value).IsEqualTo((double)ws.Cell("D3").Value).Within(1e-12);
            await Assert.That((double)ws.Cell("D4").Value).IsEqualTo(0d);
            await Assert.That((double)ws.Cell("D5").Value).IsEqualTo(0d).Within(1e-15);
        }
    }

    [Test]
    public async Task FTest_OfASampleAgainstItselfIsCertainty()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedTwoSamples(ws);
            ws.Cell("D1").FormulaA1 = "F.TEST(A1:A9, A1:A9)";

            // Identical variances give a ratio of one, whose two-tailed probability is one.
            await Assert.That((double)ws.Cell("D1").Value).IsEqualTo(1d).Within(1e-12);
        }
    }

    [Test]
    public async Task TTest_MatchesTheTDistributionOfEachTestStatistic()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedTwoSamples(ws);

            // Type 1, paired: the mean difference over its own standard error, on n-1 degrees.
            ws.Cell("D1").FormulaA1 = "T.TEST(A1:A9, B1:B9, 2, 1)";
            ws.Cell("D2").FormulaA1 = "AVERAGE(A1:A9 - B1:B9) / (STDEV.S(A1:A9 - B1:B9) / SQRT(9))";
            ws.Cell("D3").FormulaA1 = "T.DIST.2T(ABS(D2), 8)";

            // Type 2, equal variances: the pooled standard error on n1+n2-2 degrees.
            ws.Cell("E1").FormulaA1 = "T.TEST(A1:A9, B1:B9, 2, 2)";
            ws.Cell("E2").FormulaA1 = "(AVERAGE(A1:A9) - AVERAGE(B1:B9)) / SQRT((8 * VAR.S(A1:A9) + 8 * VAR.S(B1:B9)) / 16 * (1/9 + 1/9))";
            ws.Cell("E3").FormulaA1 = "T.DIST.2T(ABS(E2), 16)";

            await Assert.That((double)ws.Cell("D1").Value).IsEqualTo((double)ws.Cell("D3").Value).Within(1e-10);
            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo((double)ws.Cell("E3").Value).Within(1e-10);
        }
    }

    [Test]
    public async Task TTest_UnequalVariancesUsesTheWelchDegreesOfFreedom()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedTwoSamples(ws);

            ws.Cell("D1").FormulaA1 = "T.TEST(A1:A9, B1:B9, 2, 3)";
            ws.Cell("D2").FormulaA1 = "(AVERAGE(A1:A9) - AVERAGE(B1:B9)) / SQRT(VAR.S(A1:A9)/9 + VAR.S(B1:B9)/9)";
            // Welch–Satterthwaite, written out.
            ws.Cell("D3").FormulaA1 = "(VAR.S(A1:A9)/9 + VAR.S(B1:B9)/9)^2 / ((VAR.S(A1:A9)/9)^2/8 + (VAR.S(B1:B9)/9)^2/8)";
            ws.Cell("D4").FormulaA1 = "T.DIST.2T(ABS(D2), D3)";

            // The adjusted degrees of freedom are generally not a whole number.
            await Assert.That((double)ws.Cell("D3").Value).IsNotEqualTo(16d);
            await Assert.That((double)ws.Cell("D1").Value).IsEqualTo((double)ws.Cell("D4").Value).Within(1e-10);
        }
    }

    [Test]
    public async Task TTest_OneTailedIsHalfTheTwoTailed()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedTwoSamples(ws);
            ws.Cell("D1").FormulaA1 = "T.TEST(A1:A9, B1:B9, 1, 2) * 2 - T.TEST(A1:A9, B1:B9, 2, 2)";
            ws.Cell("D2").FormulaA1 = "TTEST(A1:A9, B1:B9, 2, 2) - T.TEST(A1:A9, B1:B9, 2, 2)";

            await Assert.That((double)ws.Cell("D1").Value).IsEqualTo(0d).Within(1e-14);
            await Assert.That((double)ws.Cell("D2").Value).IsEqualTo(0d);
        }
    }

    [Test]
    public async Task TTest_OutOfRangeArguments()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedTwoSamples(ws);
            ws.Cell("D1").FormulaA1 = "T.TEST(A1:A9, B1:B9, 3, 2)"; // One tail or two.
            ws.Cell("D2").FormulaA1 = "T.TEST(A1:A9, B1:B9, 2, 4)"; // Types are 1, 2 and 3.
            ws.Cell("D3").FormulaA1 = "T.TEST(A1:A9, B1:B8, 2, 1)"; // Paired needs equal lengths.

            await Assert.That(ws.Cell("D1").Value).IsEqualTo(XLError.NumberInvalid);
            await Assert.That(ws.Cell("D2").Value).IsEqualTo(XLError.NumberInvalid);
            await Assert.That(ws.Cell("D3").Value).IsEqualTo(XLError.NoValueAvailable);
        }
    }

    [Test]
    public async Task ZTest_IsTheUpperTailOfTheStandardisedSampleMean()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            double[] sample = [3, 6, 7, 8, 6, 5, 4, 2, 1, 9];
            for (var i = 0; i < sample.Length; i++)
                ws.Cell(i + 1, 1).Value = sample[i];

            ws.Cell("C1").FormulaA1 = "Z.TEST(A1:A10, 4)";
            ws.Cell("C2").FormulaA1 = "1 - NORM.S.DIST((AVERAGE(A1:A10) - 4) / (STDEV.S(A1:A10) / SQRT(10)), TRUE)";
            ws.Cell("C3").FormulaA1 = "Z.TEST(A1:A10, 4, 3)"; // With a known population sigma.
            ws.Cell("C4").FormulaA1 = "1 - NORM.S.DIST((AVERAGE(A1:A10) - 4) / (3 / SQRT(10)), TRUE)";
            ws.Cell("C5").FormulaA1 = "ZTEST(A1:A10, 4) - Z.TEST(A1:A10, 4)";
            // Testing against the sample's own mean leaves exactly half the mass above it.
            ws.Cell("C6").FormulaA1 = "Z.TEST(A1:A10, AVERAGE(A1:A10))";

            await Assert.That((double)ws.Cell("C1").Value).IsEqualTo((double)ws.Cell("C2").Value).Within(1e-12);
            await Assert.That((double)ws.Cell("C3").Value).IsEqualTo((double)ws.Cell("C4").Value).Within(1e-12);
            await Assert.That((double)ws.Cell("C5").Value).IsEqualTo(0d);
            await Assert.That((double)ws.Cell("C6").Value).IsEqualTo(0.5d).Within(1e-12);
        }
    }

    [Test]
    public async Task ZTest_NonPositiveSigmaReturnsNumberInvalid()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 2;
            ws.Cell("C1").FormulaA1 = "Z.TEST(A1:A2, 1, 0)";

            await Assert.That(ws.Cell("C1").Value).IsEqualTo(XLError.NumberInvalid);
        }
    }

    [Test]
    public async Task RankAvg_SharesTheRanksATiedGroupOccupies()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            double[] values = [7, 3.5, 3.5, 1, 2];
            for (var i = 0; i < values.Length; i++)
                ws.Cell(i + 1, 1).Value = values[i];

            // Descending: 7 is first, the two 3.5s take ranks two and three, 2 is fourth, 1 is fifth.
            ws.Cell("C1").FormulaA1 = "RANK.EQ(3.5, A1:A5)";
            ws.Cell("C2").FormulaA1 = "RANK.AVG(3.5, A1:A5)";
            ws.Cell("C3").FormulaA1 = "RANK.AVG(7, A1:A5)";
            ws.Cell("C4").FormulaA1 = "RANK.AVG(1, A1:A5)";
            ws.Cell("C5").FormulaA1 = "RANK.AVG(1, A1:A5, 1)"; // Ascending.

            await Assert.That((double)ws.Cell("C1").Value).IsEqualTo(2d);
            await Assert.That((double)ws.Cell("C2").Value).IsEqualTo(2.5d);
            await Assert.That((double)ws.Cell("C3").Value).IsEqualTo(1d);
            await Assert.That((double)ws.Cell("C4").Value).IsEqualTo(5d);
            await Assert.That((double)ws.Cell("C5").Value).IsEqualTo(1d);
        }
    }

    [Test]
    public async Task ModeMult_ReturnsEveryTiedMode()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            double[] values = [1, 2, 2, 3, 3, 4];
            for (var i = 0; i < values.Length; i++)
                ws.Cell(i + 1, 1).Value = values[i];

            ws.Range("C1:C2").FormulaArrayA1 = "MODE.MULT(A1:A6)";
            ws.Cell("D1").FormulaA1 = "MODE.SNGL(A1:A6)";

            await Assert.That((double)ws.Cell("C1").Value).IsEqualTo(2d);
            await Assert.That((double)ws.Cell("C2").Value).IsEqualTo(3d);
            // MODE.SNGL reports only the first of them.
            await Assert.That((double)ws.Cell("D1").Value).IsEqualTo(2d);
        }
    }

    [Test]
    public async Task ModeMult_WithoutARepeatedValueReturnsNoValueAvailable()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 2;
            ws.Cell("C1").FormulaA1 = "MODE.MULT(A1:A2)";

            await Assert.That(ws.Cell("C1").Value).IsEqualTo(XLError.NoValueAvailable);
        }
    }

    [Test]
    public async Task PercentileAndQuartile_ExclusiveVariants()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            for (var row = 1; row <= 5; row++)
                ws.Cell(row, 1).Value = row;

            // With five values the exclusive rank of the median is 0.5*(5+1) = 3, the third value.
            ws.Cell("C1").FormulaA1 = "PERCENTILE.EXC(A1:A5, 0.5)";
            ws.Cell("C2").FormulaA1 = "PERCENTILE.INC(A1:A5, 0.5)";
            ws.Cell("C3").FormulaA1 = "QUARTILE.EXC(A1:A5, 2)";
            ws.Cell("C4").FormulaA1 = "QUARTILE.INC(A1:A5, 2)";
            // The exclusive form cannot reach the ends of the range.
            ws.Cell("C5").FormulaA1 = "PERCENTILE.EXC(A1:A5, 0.1)";
            ws.Cell("C6").FormulaA1 = "QUARTILE.EXC(A1:A5, 0)";

            await Assert.That((double)ws.Cell("C1").Value).IsEqualTo(3d).Within(1e-12);
            await Assert.That((double)ws.Cell("C2").Value).IsEqualTo(3d).Within(1e-12);
            await Assert.That((double)ws.Cell("C3").Value).IsEqualTo(3d).Within(1e-12);
            await Assert.That((double)ws.Cell("C4").Value).IsEqualTo(3d).Within(1e-12);
            await Assert.That(ws.Cell("C5").Value).IsEqualTo(XLError.NumberInvalid);
            await Assert.That(ws.Cell("C6").Value).IsEqualTo(XLError.NumberInvalid);
        }
    }

    [Test]
    public async Task PercentRank_IsTheInverseOfPercentile()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            for (var row = 1; row <= 5; row++)
                ws.Cell(row, 1).Value = row;

            // Five evenly spaced values sit at 0, 0.25, 0.5, 0.75 and 1 on the inclusive scale.
            ws.Cell("C1").FormulaA1 = "PERCENTRANK(A1:A5, 1)";
            ws.Cell("C2").FormulaA1 = "PERCENTRANK(A1:A5, 3)";
            ws.Cell("C3").FormulaA1 = "PERCENTRANK(A1:A5, 5)";
            ws.Cell("C4").FormulaA1 = "PERCENTRANK.INC(A1:A5, 2)";
            ws.Cell("C5").FormulaA1 = "PERCENTRANK(A1:A5, 2.5)"; // Interpolated between two values.
            ws.Cell("C6").FormulaA1 = "PERCENTILE.INC(A1:A5, PERCENTRANK(A1:A5, 4))";

            await Assert.That((double)ws.Cell("C1").Value).IsEqualTo(0d);
            await Assert.That((double)ws.Cell("C2").Value).IsEqualTo(0.5d);
            await Assert.That((double)ws.Cell("C3").Value).IsEqualTo(1d);
            await Assert.That((double)ws.Cell("C4").Value).IsEqualTo(0.25d);
            await Assert.That((double)ws.Cell("C5").Value).IsEqualTo(0.375d);
            await Assert.That((double)ws.Cell("C6").Value).IsEqualTo(4d).Within(1e-12);
        }
    }

    [Test]
    public async Task PercentRank_TruncatesToTheRequestedSignificance()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            for (var row = 1; row <= 4; row++)
                ws.Cell(row, 1).Value = row;

            // A rank of 1/3 truncated, not rounded: 0.333 at three digits, 0.3 at one.
            ws.Cell("C1").FormulaA1 = "PERCENTRANK(A1:A4, 2)";
            ws.Cell("C2").FormulaA1 = "PERCENTRANK(A1:A4, 2, 1)";
            ws.Cell("C3").FormulaA1 = "PERCENTRANK(A1:A4, 2, 5)";
            ws.Cell("C4").FormulaA1 = "PERCENTRANK(A1:A4, 2, 0)"; // At least one digit is required.

            await Assert.That((double)ws.Cell("C1").Value).IsEqualTo(0.333d);
            await Assert.That((double)ws.Cell("C2").Value).IsEqualTo(0.3d);
            await Assert.That((double)ws.Cell("C3").Value).IsEqualTo(0.33333d);
            await Assert.That(ws.Cell("C4").Value).IsEqualTo(XLError.NumberInvalid);
        }
    }

    [Test]
    public async Task PercentRank_ExclusiveKeepsTheEndsUnreachable()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            for (var row = 1; row <= 5; row++)
                ws.Cell(row, 1).Value = row;

            // On the exclusive scale the k-th of n values sits at k/(n+1).
            ws.Cell("C1").FormulaA1 = "PERCENTRANK.EXC(A1:A5, 3)";
            ws.Cell("C2").FormulaA1 = "PERCENTRANK.EXC(A1:A5, 1)";
            ws.Cell("C3").FormulaA1 = "PERCENTRANK.EXC(A1:A5, 5)";

            await Assert.That((double)ws.Cell("C1").Value).IsEqualTo(0.5d);
            await Assert.That((double)ws.Cell("C2").Value).IsEqualTo(0.166d); // 1/6, truncated.
            await Assert.That((double)ws.Cell("C3").Value).IsEqualTo(0.833d); // 5/6, truncated.
        }
    }

    [Test]
    public async Task PercentRank_OutsideTheDataIsNoValueAvailable()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            for (var row = 1; row <= 5; row++)
                ws.Cell(row, 1).Value = row;

            ws.Cell("C1").FormulaA1 = "PERCENTRANK(A1:A5, 0)";
            ws.Cell("C2").FormulaA1 = "PERCENTRANK(A1:A5, 6)";

            await Assert.That(ws.Cell("C1").Value).IsEqualTo(XLError.NoValueAvailable);
            await Assert.That(ws.Cell("C2").Value).IsEqualTo(XLError.NoValueAvailable);
        }
    }
}
