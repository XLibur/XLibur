using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// The statistical distributions. Expected values are, in order of preference: closed-form results
/// the distribution takes at a special point (the normal CDF at its mean is exactly a half, a
/// chi-squared with two degrees of freedom is an exponential), published critical values, and the
/// worked examples in Microsoft's per-function documentation. Every inverse is additionally checked
/// by round-tripping it through its own distribution, which pins both directions at once.
/// </summary>
[SetCulture("en-US")]
public class DistributionTests
{
    private const double Tolerance = 1e-9;

    private static IXLWorksheet NewSheet(out XLWorkbook wb)
    {
        wb = new XLWorkbook();
        return wb.AddWorksheet("Sheet1");
    }

    #region Normal

    [Test]
    [Arguments("NORM.S.DIST(0, TRUE)", 0.5d)] // Half the mass lies below the mean.
    [Arguments("NORM.S.DIST(0, FALSE)", 0.3989422804014327d)] // 1/sqrt(2*pi).
    [Arguments("NORM.DIST(40, 40, 1.5, TRUE)", 0.5d)]
    [Arguments("NORM.DIST(40, 40, 1.5, FALSE)", 0.26596152026762186d)] // 1/(1.5*sqrt(2*pi)).
    // 1.96 and 1.6449 are the standard normal's two best-known critical values.
    [Arguments("NORM.S.DIST(1.959963984540054, TRUE)", 0.975d)]
    [Arguments("NORM.S.DIST(1.6448536269514722, TRUE)", 0.95d)]
    [Arguments("NORM.S.INV(0.975)", 1.959963984540054d)]
    [Arguments("NORM.S.INV(0.95)", 1.6448536269514722d)]
    [Arguments("NORM.S.INV(0.5)", 0d)]
    // Microsoft's NORMDIST example: 42 against a mean of 40 and a standard deviation of 1.5.
    [Arguments("NORM.DIST(42, 40, 1.5, TRUE)", 0.9087887802741321d)]
    public async Task Normal_DistributionAndInverse(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(Tolerance);
    }

    [Test]
    public async Task Normal_IsSymmetricAboutItsMean()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "NORM.S.DIST(1.3, TRUE) + NORM.S.DIST(-1.3, TRUE)";
            ws.Cell("A2").FormulaA1 = "NORM.DIST(43, 40, 1.5, TRUE) + NORM.DIST(37, 40, 1.5, TRUE)";

            await Assert.That((double)ws.Cell("A1").Value).IsEqualTo(1d).Within(1e-12);
            await Assert.That((double)ws.Cell("A2").Value).IsEqualTo(1d).Within(1e-12);
        }
    }

    [Test]
    public async Task Normal_InverseUndoesTheDistribution()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "NORM.S.DIST(NORM.S.INV(0.123), TRUE)";
            ws.Cell("A2").FormulaA1 = "NORM.DIST(NORM.INV(0.876, 40, 1.5), 40, 1.5, TRUE)";
            ws.Cell("A3").FormulaA1 = "NORM.S.DIST(NORM.S.INV(0.0000001), TRUE)"; // Deep in the tail.

            await Assert.That((double)ws.Cell("A1").Value).IsEqualTo(0.123d).Within(1e-12);
            await Assert.That((double)ws.Cell("A2").Value).IsEqualTo(0.876d).Within(1e-12);
            await Assert.That((double)ws.Cell("A3").Value).IsEqualTo(0.0000001d).Within(1e-16);
        }
    }

    [Test]
    public async Task Normal_LegacyNamesMatchTheDottedOnes()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "NORMDIST(42, 40, 1.5, TRUE) - NORM.DIST(42, 40, 1.5, TRUE)";
            ws.Cell("A2").FormulaA1 = "NORMINV(0.9, 40, 1.5) - NORM.INV(0.9, 40, 1.5)";
            ws.Cell("A3").FormulaA1 = "NORMSDIST(1.3) - NORM.S.DIST(1.3, TRUE)";
            ws.Cell("A4").FormulaA1 = "NORMSINV(0.9) - NORM.S.INV(0.9)";

            foreach (var address in new[] { "A1", "A2", "A3", "A4" })
                await Assert.That((double)ws.Cell(address).Value).IsEqualTo(0d).Within(1e-15);
        }
    }

    [Test]
    [Arguments("NORM.DIST(1, 0, 0, TRUE)")] // The standard deviation must be positive.
    [Arguments("NORM.DIST(1, 0, -1, TRUE)")]
    [Arguments("NORM.INV(0, 0, 1)")] // The probability must be strictly inside 0..1.
    [Arguments("NORM.INV(1, 0, 1)")]
    [Arguments("NORM.INV(0.5, 0, 0)")]
    [Arguments("NORM.S.INV(0)")]
    [Arguments("NORM.S.INV(1)")]
    public async Task Normal_OutOfRangeArgumentsReturnNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    #endregion

    #region Lognormal

    [Test]
    // A lognormal variable is the exponential of a normal one, so its CDF at exp(mean) is a half.
    [Arguments("LOGNORM.DIST(1, 0, 1, TRUE)", 0.5d)]
    [Arguments("LOGNORM.DIST(2.718281828459045, 1, 1, TRUE)", 0.5d)]
    [Arguments("LOGNORM.INV(0.5, 0, 1)", 1d)]
    [Arguments("LOGNORM.INV(0.5, 1, 1)", 2.718281828459045d)]
    public async Task Lognormal_DistributionAndInverse(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(Tolerance);
    }

    [Test]
    public async Task Lognormal_TracksTheNormalOfTheLogarithm()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "LOGNORM.DIST(4, 3.5, 1.2, TRUE) - NORM.DIST(LN(4), 3.5, 1.2, TRUE)";
            ws.Cell("A2").FormulaA1 = "LOGNORM.DIST(LOGNORM.INV(0.42, 3.5, 1.2), 3.5, 1.2, TRUE)";
            ws.Cell("A3").FormulaA1 = "LOGNORMDIST(4, 3.5, 1.2) - LOGNORM.DIST(4, 3.5, 1.2, TRUE)";
            ws.Cell("A4").FormulaA1 = "LOGINV(0.42, 3.5, 1.2) - LOGNORM.INV(0.42, 3.5, 1.2)";

            await Assert.That((double)ws.Cell("A1").Value).IsEqualTo(0d).Within(1e-15);
            await Assert.That((double)ws.Cell("A2").Value).IsEqualTo(0.42d).Within(1e-12);
            await Assert.That((double)ws.Cell("A3").Value).IsEqualTo(0d).Within(1e-15);
            await Assert.That((double)ws.Cell("A4").Value).IsEqualTo(0d).Within(1e-15);
        }
    }

    [Test]
    [Arguments("LOGNORM.DIST(0, 0, 1, TRUE)")] // A lognormal variable is strictly positive.
    [Arguments("LOGNORM.DIST(-1, 0, 1, TRUE)")]
    [Arguments("LOGNORM.DIST(1, 0, 0, TRUE)")]
    [Arguments("LOGNORM.INV(0, 0, 1)")]
    [Arguments("LOGNORM.INV(1, 0, 1)")]
    public async Task Lognormal_OutOfRangeArgumentsReturnNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    #endregion

    #region Chi-squared

    [Test]
    // Chi-squared with two degrees of freedom is an exponential with mean two, so its right tail is
    // exactly exp(-x/2) and its inverse exactly -2*ln(p).
    [Arguments("CHISQ.DIST.RT(2, 2)", 0.36787944117144233d)]
    [Arguments("CHISQ.DIST.RT(0, 2)", 1d)]
    [Arguments("CHISQ.DIST(2, 2, TRUE)", 0.6321205588285577d)]
    [Arguments("CHISQ.INV.RT(0.5, 2)", 1.3862943611198906d)] // 2*ln(2).
    [Arguments("CHISQ.INV(0.5, 2)", 1.3862943611198906d)] // The same by symmetry of this case.
    // The 5% critical values of the chi-squared distribution, as published in every statistics text.
    [Arguments("CHISQ.INV.RT(0.05, 1)", 3.841458820694124d)]
    [Arguments("CHISQ.INV.RT(0.05, 2)", 5.991464547107979d)]
    [Arguments("CHISQ.INV.RT(0.05, 10)", 18.307038053275146d)]
    [Arguments("CHISQ.DIST.RT(3.841458820694124, 1)", 0.05d)]
    public async Task ChiSquared_DistributionAndInverse(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-8);
    }

    [Test]
    public async Task ChiSquared_TailsSumToOneAndInvertEachOther()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "CHISQ.DIST(7.3, 5, TRUE) + CHISQ.DIST.RT(7.3, 5)";
            ws.Cell("A2").FormulaA1 = "CHISQ.DIST.RT(CHISQ.INV.RT(0.037, 7), 7)";
            ws.Cell("A3").FormulaA1 = "CHISQ.DIST(CHISQ.INV(0.037, 7), 7, TRUE)";
            ws.Cell("A4").FormulaA1 = "CHIDIST(18.307038053275146, 10) - CHISQ.DIST.RT(18.307038053275146, 10)";
            ws.Cell("A5").FormulaA1 = "CHIINV(0.05, 10) - CHISQ.INV.RT(0.05, 10)";

            await Assert.That((double)ws.Cell("A1").Value).IsEqualTo(1d).Within(1e-12);
            await Assert.That((double)ws.Cell("A2").Value).IsEqualTo(0.037d).Within(1e-10);
            await Assert.That((double)ws.Cell("A3").Value).IsEqualTo(0.037d).Within(1e-10);
            await Assert.That((double)ws.Cell("A4").Value).IsEqualTo(0d).Within(1e-15);
            await Assert.That((double)ws.Cell("A5").Value).IsEqualTo(0d).Within(1e-15);
        }
    }

    [Test]
    [Arguments("CHISQ.DIST(-1, 2, TRUE)")] // A chi-squared variable is non-negative.
    [Arguments("CHISQ.DIST(1, 0, TRUE)")] // At least one degree of freedom.
    [Arguments("CHISQ.DIST.RT(-1, 2)")]
    [Arguments("CHISQ.INV(-0.1, 2)")]
    [Arguments("CHISQ.INV(1.1, 2)")]
    [Arguments("CHISQ.INV.RT(1.1, 2)")]
    public async Task ChiSquared_OutOfRangeArgumentsReturnNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    #endregion

    #region F distribution

    [Test]
    // The F distribution's median is not closed form, but F(1,1) has CDF (2/pi)*atan(sqrt(x)), so
    // its median is exactly 1 and its 1/2-quantile checkable by hand.
    [Arguments("F.DIST(1, 1, 1, TRUE)", 0.5d)]
    [Arguments("F.INV(0.5, 1, 1)", 1d)]
    [Arguments("F.DIST(3, 1, 1, TRUE)", 0.66666666666666663d)] // (2/pi)*atan(sqrt(3)) = 2/3.
    [Arguments("F.DIST.RT(1, 1, 1)", 0.5d)]
    public async Task FDistribution_ClosedFormCases(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-9);
    }

    [Test]
    public async Task FDistribution_TailsAndInversesAgree()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "F.DIST(2.5, 6, 4, TRUE) + F.DIST.RT(2.5, 6, 4)";
            ws.Cell("A2").FormulaA1 = "F.DIST.RT(F.INV.RT(0.05, 6, 4), 6, 4)";
            ws.Cell("A3").FormulaA1 = "F.DIST(F.INV(0.05, 6, 4), 6, 4, TRUE)";
            ws.Cell("A4").FormulaA1 = "FDIST(2.5, 6, 4) - F.DIST.RT(2.5, 6, 4)";
            ws.Cell("A5").FormulaA1 = "FINV(0.05, 6, 4) - F.INV.RT(0.05, 6, 4)";
            // The F distribution reciprocates when its degrees of freedom swap over.
            ws.Cell("A6").FormulaA1 = "F.INV.RT(0.05, 6, 4) * F.INV(0.05, 4, 6)";

            await Assert.That((double)ws.Cell("A1").Value).IsEqualTo(1d).Within(1e-12);
            await Assert.That((double)ws.Cell("A2").Value).IsEqualTo(0.05d).Within(1e-10);
            await Assert.That((double)ws.Cell("A3").Value).IsEqualTo(0.05d).Within(1e-10);
            await Assert.That((double)ws.Cell("A4").Value).IsEqualTo(0d).Within(1e-15);
            await Assert.That((double)ws.Cell("A5").Value).IsEqualTo(0d).Within(1e-15);
            await Assert.That((double)ws.Cell("A6").Value).IsEqualTo(1d).Within(1e-9);
        }
    }

    [Test]
    [Arguments("F.DIST(-1, 6, 4, TRUE)")]
    [Arguments("F.DIST(1, 0, 4, TRUE)")]
    [Arguments("F.DIST.RT(1, 6, 0)")]
    [Arguments("F.INV(-0.1, 6, 4)")]
    [Arguments("F.INV.RT(1.1, 6, 4)")]
    public async Task FDistribution_OutOfRangeArgumentsReturnNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    #endregion

    #region Student's t

    [Test]
    [Arguments("T.DIST(0, 10, TRUE)", 0.5d)] // The t-distribution is symmetric about zero.
    [Arguments("T.DIST.RT(0, 10)", 0.5d)]
    [Arguments("T.DIST.2T(0, 10)", 1d)]
    // With one degree of freedom the t-distribution is Cauchy: CDF = 1/2 + atan(x)/pi.
    [Arguments("T.DIST(1, 1, TRUE)", 0.75d)]
    [Arguments("T.DIST(1, 1, FALSE)", 0.15915494309189535d)] // 1/(pi*(1+x^2)) at x = 1.
    // The two-tailed 5% critical values, as published in every t-table.
    [Arguments("T.INV.2T(0.05, 10)", 2.2281388519649385d)]
    [Arguments("T.INV.2T(0.05, 30)", 2.0422724563012373d)]
    public async Task StudentT_DistributionAndInverse(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-8);
    }

    [Test]
    public async Task StudentT_TailsAndInversesAgree()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "T.DIST(1.7, 12, TRUE) + T.DIST.RT(1.7, 12)";
            ws.Cell("A2").FormulaA1 = "T.DIST.RT(T.INV(0.95, 12), 12)";
            ws.Cell("A3").FormulaA1 = "T.DIST.2T(T.INV.2T(0.05, 12), 12)";
            ws.Cell("A4").FormulaA1 = "T.DIST.2T(1.7, 12) - 2 * T.DIST.RT(1.7, 12)";
            ws.Cell("A5").FormulaA1 = "TDIST(1.7, 12, 2) - T.DIST.2T(1.7, 12)";
            ws.Cell("A6").FormulaA1 = "TDIST(1.7, 12, 1) - T.DIST.RT(1.7, 12)";

            await Assert.That((double)ws.Cell("A1").Value).IsEqualTo(1d).Within(1e-12);
            await Assert.That((double)ws.Cell("A2").Value).IsEqualTo(0.05d).Within(1e-10);
            await Assert.That((double)ws.Cell("A3").Value).IsEqualTo(0.05d).Within(1e-10);
            await Assert.That((double)ws.Cell("A4").Value).IsEqualTo(0d).Within(1e-14);
            await Assert.That((double)ws.Cell("A5").Value).IsEqualTo(0d).Within(1e-14);
            await Assert.That((double)ws.Cell("A6").Value).IsEqualTo(0d).Within(1e-14);
        }
    }

    [Test]
    [Arguments("T.DIST(1, 0, TRUE)")]
    [Arguments("T.DIST.2T(-1, 10)")] // The two-tailed form takes a non-negative argument.
    [Arguments("TDIST(-1, 10, 1)")]
    [Arguments("TDIST(1, 10, 3)")] // One tail or two.
    public async Task StudentT_OutOfRangeArgumentsReturnNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    #endregion

    #region Exponential families

    [Test]
    [Arguments("EXPON.DIST(0, 10, TRUE)", 0d)]
    [Arguments("EXPON.DIST(0, 10, FALSE)", 10d)] // The density at zero is the rate itself.
    [Arguments("EXPON.DIST(0.2, 10, TRUE)", 0.8646647167633873d)] // 1 - exp(-2).
    [Arguments("EXPON.DIST(0.2, 10, FALSE)", 1.353352832366127d)] // 10*exp(-2).
    [Arguments("POISSON.DIST(0, 1, FALSE)", 0.36787944117144233d)] // exp(-1).
    [Arguments("POISSON.DIST(0, 1, TRUE)", 0.36787944117144233d)]
    [Arguments("POISSON.DIST(1, 1, TRUE)", 0.7357588823428847d)] // 2*exp(-1).
    [Arguments("WEIBULL.DIST(1, 1, 1, TRUE)", 0.6321205588285577d)] // Shape one is the exponential.
    [Arguments("WEIBULL.DIST(1, 1, 1, FALSE)", 0.36787944117144233d)]
    [Arguments("GAMMA.DIST(1, 1, 1, TRUE)", 0.6321205588285577d)] // Shape one is again exponential.
    [Arguments("GAMMA.DIST(1, 1, 1, FALSE)", 0.36787944117144233d)]
    [Arguments("GAMMA(5)", 24d)] // Gamma extends the factorial: GAMMA(n) = (n-1)!.
    [Arguments("GAMMA(1)", 1d)]
    [Arguments("GAMMA(0.5)", 1.7724538509055159d)] // sqrt(pi).
    [Arguments("GAMMA(-0.5)", -3.5449077018110318d)] // -2*sqrt(pi).
    [Arguments("GAMMALN(5)", 3.1780538303479458d)] // ln(24).
    [Arguments("GAMMALN.PRECISE(5)", 3.1780538303479458d)]
    public async Task ExponentialFamilies_ClosedFormCases(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-9);
    }

    [Test]
    public async Task Gamma_InvertsItsOwnDistributionAndMatchesTheLegacyNames()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "GAMMA.DIST(GAMMA.INV(0.31, 9, 2), 9, 2, TRUE)";
            ws.Cell("A2").FormulaA1 = "GAMMADIST(10, 9, 2, TRUE) - GAMMA.DIST(10, 9, 2, TRUE)";
            ws.Cell("A3").FormulaA1 = "GAMMAINV(0.31, 9, 2) - GAMMA.INV(0.31, 9, 2)";
            ws.Cell("A4").FormulaA1 = "EXPONDIST(0.2, 10, TRUE) - EXPON.DIST(0.2, 10, TRUE)";
            ws.Cell("A5").FormulaA1 = "POISSON(2, 5, TRUE) - POISSON.DIST(2, 5, TRUE)";
            ws.Cell("A6").FormulaA1 = "WEIBULL(105, 20, 100, TRUE) - WEIBULL.DIST(105, 20, 100, TRUE)";
            // GAMMALN is the logarithm of GAMMA wherever both are defined.
            ws.Cell("A7").FormulaA1 = "EXP(GAMMALN(7.3)) - GAMMA(7.3)";

            await Assert.That((double)ws.Cell("A1").Value).IsEqualTo(0.31d).Within(1e-10);
            foreach (var address in new[] { "A2", "A3", "A4", "A5", "A6" })
                await Assert.That((double)ws.Cell(address).Value).IsEqualTo(0d).Within(1e-15);

            await Assert.That((double)ws.Cell("A7").Value).IsEqualTo(0d).Within(1e-6);
        }
    }

    [Test]
    public async Task Poisson_CumulativeIsTheRunningSumOfItsTerms()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "POISSON.DIST(0, 5, FALSE) + POISSON.DIST(1, 5, FALSE) + POISSON.DIST(2, 5, FALSE)";
            ws.Cell("A2").FormulaA1 = "POISSON.DIST(2, 5, TRUE)";

            await Assert.That((double)ws.Cell("A2").Value).IsEqualTo((double)ws.Cell("A1").Value).Within(1e-12);
        }
    }

    [Test]
    [Arguments("EXPON.DIST(-1, 10, TRUE)")]
    [Arguments("EXPON.DIST(1, 0, TRUE)")]
    [Arguments("POISSON.DIST(-1, 5, TRUE)")]
    [Arguments("POISSON.DIST(1, -1, TRUE)")]
    [Arguments("WEIBULL.DIST(-1, 1, 1, TRUE)")]
    [Arguments("WEIBULL.DIST(1, 0, 1, TRUE)")]
    [Arguments("WEIBULL.DIST(1, 1, 0, TRUE)")]
    [Arguments("GAMMA.DIST(-1, 1, 1, TRUE)")]
    [Arguments("GAMMA.DIST(1, 0, 1, TRUE)")]
    [Arguments("GAMMA.INV(1.1, 1, 1)")]
    [Arguments("GAMMA(0)")] // The gamma function has a pole at every non-positive integer.
    [Arguments("GAMMA(-1)")]
    [Arguments("GAMMALN(0)")]
    [Arguments("GAMMALN(-1)")]
    public async Task ExponentialFamilies_OutOfRangeArgumentsReturnNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    #endregion

    #region Beta

    [Test]
    // Beta(1,1) is the uniform distribution, so its CDF is the identity and its density is one.
    [Arguments("BETA.DIST(0.3, 1, 1, TRUE)", 0.3d)]
    [Arguments("BETA.DIST(0.3, 1, 1, FALSE)", 1d)]
    [Arguments("BETA.INV(0.3, 1, 1)", 0.3d)]
    [Arguments("BETA.DIST(0.5, 2, 2, TRUE)", 0.5d)] // Symmetric shapes put the median in the middle.
    [Arguments("BETA.INV(0.5, 2, 2)", 0.5d)]
    // Rescaled onto [1, 3]: the uniform CDF is (x-1)/2.
    [Arguments("BETA.DIST(2, 1, 1, TRUE, 1, 3)", 0.5d)]
    [Arguments("BETA.INV(0.5, 1, 1, 1, 3)", 2d)]
    public async Task Beta_DistributionAndInverse(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-9);
    }

    [Test]
    public async Task Beta_InverseUndoesTheDistributionAndMatchesTheLegacyNames()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "BETA.DIST(BETA.INV(0.685, 8, 10, 1, 3), 8, 10, TRUE, 1, 3)";
            ws.Cell("A2").FormulaA1 = "BETADIST(2, 8, 10, 1, 3) - BETA.DIST(2, 8, 10, TRUE, 1, 3)";
            ws.Cell("A3").FormulaA1 = "BETAINV(0.685, 8, 10, 1, 3) - BETA.INV(0.685, 8, 10, 1, 3)";

            await Assert.That((double)ws.Cell("A1").Value).IsEqualTo(0.685d).Within(1e-10);
            await Assert.That((double)ws.Cell("A2").Value).IsEqualTo(0d).Within(1e-15);
            await Assert.That((double)ws.Cell("A3").Value).IsEqualTo(0d).Within(1e-15);
        }
    }

    [Test]
    [Arguments("BETA.DIST(0.5, 0, 1, TRUE)")]
    [Arguments("BETA.DIST(0.5, 1, 0, TRUE)")]
    [Arguments("BETA.DIST(2, 1, 1, TRUE)")] // Outside the default 0..1 support.
    [Arguments("BETA.DIST(-1, 1, 1, TRUE)")]
    [Arguments("BETA.DIST(2, 1, 1, TRUE, 3, 1)")] // The bounds must be the right way round.
    [Arguments("BETA.INV(0, 1, 1)")]
    [Arguments("BETA.INV(1.1, 1, 1)")]
    public async Task Beta_OutOfRangeArgumentsReturnNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    #endregion

    #region Discrete distributions

    [Test]
    // Drawing 4 from an urn of 20 holding 8 successes: C(8,1)*C(12,3)/C(20,4) = 8*220/4845.
    [Arguments("HYPGEOM.DIST(1, 4, 8, 20, FALSE)", 0.36326109391124871d)]
    [Arguments("HYPGEOMDIST(1, 4, 8, 20)", 0.36326109391124871d)]
    // A sample of one is just the proportion of successes in the population.
    [Arguments("HYPGEOM.DIST(1, 1, 8, 20, FALSE)", 0.4d)]
    [Arguments("HYPGEOM.DIST(0, 1, 8, 20, FALSE)", 0.6d)]
    [Arguments("HYPGEOM.DIST(1, 1, 8, 20, TRUE)", 1d)]
    // Ten failures before the fifth success: C(14,4)*0.25^5*0.75^10.
    [Arguments("NEGBINOM.DIST(10, 5, 0.25, FALSE)", 0.05504866037517786d)]
    [Arguments("NEGBINOMDIST(10, 5, 0.25)", 0.05504866037517786d)]
    // With one success required this is the geometric distribution: p*(1-p)^f.
    [Arguments("NEGBINOM.DIST(2, 1, 0.5, FALSE)", 0.125d)]
    [Arguments("NEGBINOM.DIST(2, 1, 0.5, TRUE)", 0.875d)] // 1 - 0.5^3.
    public async Task DiscreteDistributions_ClosedFormCases(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-12);
    }

    [Test]
    public async Task HypGeom_CumulativeIsTheRunningSumOfItsTerms()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "HYPGEOM.DIST(0, 4, 8, 20, FALSE) + HYPGEOM.DIST(1, 4, 8, 20, FALSE) + HYPGEOM.DIST(2, 4, 8, 20, FALSE)";
            ws.Cell("A2").FormulaA1 = "HYPGEOM.DIST(2, 4, 8, 20, TRUE)";
            ws.Cell("A3").FormulaA1 = "HYPGEOM.DIST(4, 4, 8, 20, TRUE)"; // Every outcome.

            await Assert.That((double)ws.Cell("A2").Value).IsEqualTo((double)ws.Cell("A1").Value).Within(1e-12);
            await Assert.That((double)ws.Cell("A3").Value).IsEqualTo(1d).Within(1e-12);
        }
    }

    [Test]
    // BINOM.INV finds the first number of successes whose cumulative probability reaches alpha.
    [Arguments("BINOM.INV(6, 0.5, 0.75)", 4d)]
    [Arguments("BINOM.INV(10, 0.5, 0.5)", 5d)]
    [Arguments("BINOM.INV(10, 0.5, 0)", 0d)]
    [Arguments("BINOM.INV(10, 0.5, 1)", 10d)]
    [Arguments("CRITBINOM(6, 0.5, 0.75)", 4d)]
    public async Task BinomInv_FindsTheCriterion(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected);
    }

    [Test]
    public async Task BinomInv_AgreesWithTheBinomialCumulative()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "BINOM.DIST(BINOM.INV(20, 0.3, 0.9), 20, 0.3, TRUE)";
            ws.Cell("A2").FormulaA1 = "BINOM.DIST(BINOM.INV(20, 0.3, 0.9) - 1, 20, 0.3, TRUE)";

            // The chosen k is the first one at or above alpha, so the one before it falls short.
            await Assert.That((double)ws.Cell("A1").Value).IsGreaterThanOrEqualTo(0.9d);
            await Assert.That((double)ws.Cell("A2").Value).IsLessThan(0.9d);
        }
    }

    [Test]
    [Arguments("HYPGEOM.DIST(5, 4, 8, 20, FALSE)")] // More successes than the sample holds.
    [Arguments("HYPGEOM.DIST(1, 4, 8, 3, FALSE)")] // A sample larger than the population.
    [Arguments("HYPGEOM.DIST(-1, 4, 8, 20, FALSE)")]
    [Arguments("NEGBINOM.DIST(-1, 5, 0.25, FALSE)")]
    [Arguments("NEGBINOM.DIST(10, 0, 0.25, FALSE)")] // At least one success is required.
    [Arguments("NEGBINOM.DIST(10, 5, 0, FALSE)")]
    [Arguments("NEGBINOM.DIST(10, 5, 1.5, FALSE)")]
    [Arguments("BINOM.INV(-1, 0.5, 0.5)")]
    [Arguments("BINOM.INV(10, 1.5, 0.5)")]
    [Arguments("BINOM.INV(10, 0.5, 1.5)")]
    public async Task DiscreteDistributions_OutOfRangeArgumentsReturnNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    #endregion

    #region Confidence intervals

    [Test]
    // Microsoft's CONFIDENCE example: alpha 5%, standard deviation 2.5, sample of 50.
    // 1.959963985 * 2.5 / sqrt(50).
    [Arguments("CONFIDENCE.NORM(0.05, 2.5, 50)", 0.6929519121748394d)]
    [Arguments("CONFIDENCE(0.05, 2.5, 50)", 0.6929519121748394d)]
    public async Task Confidence_ReferenceExampleFromExcelDocumentation(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-9);
    }

    [Test]
    public async Task Confidence_TIsWiderThanNormAndNarrowsWithTheSample()
    {
        // The t interval allows for the standard deviation being estimated, so it is always wider,
        // and the gap closes as the sample grows.
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "CONFIDENCE.T(0.05, 2.5, 50) - CONFIDENCE.NORM(0.05, 2.5, 50)";
            ws.Cell("A2").FormulaA1 = "CONFIDENCE.T(0.05, 2.5, 5000) - CONFIDENCE.NORM(0.05, 2.5, 5000)";
            ws.Cell("A3").FormulaA1 = "CONFIDENCE.T(0.05, 2.5, 50)";
            // The t interval is the two-tailed critical value scaled by the standard error.
            ws.Cell("A4").FormulaA1 = "T.INV.2T(0.05, 49) * 2.5 / SQRT(50)";

            await Assert.That((double)ws.Cell("A1").Value).IsGreaterThan(0d);
            await Assert.That((double)ws.Cell("A2").Value).IsLessThan((double)ws.Cell("A1").Value);
            await Assert.That((double)ws.Cell("A3").Value).IsEqualTo((double)ws.Cell("A4").Value).Within(1e-9);
        }
    }

    [Test]
    [Arguments("CONFIDENCE.NORM(0, 2.5, 50)")]
    [Arguments("CONFIDENCE.NORM(1, 2.5, 50)")]
    [Arguments("CONFIDENCE.NORM(0.05, 0, 50)")]
    [Arguments("CONFIDENCE.NORM(0.05, 2.5, 0)")]
    [Arguments("CONFIDENCE.T(0.05, 2.5, 0)")]
    public async Task Confidence_OutOfRangeArgumentsReturnNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task ConfidenceT_WithASingleObservationIsDivisionByZero()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("CONFIDENCE.T(0.05, 2.5, 1)")).IsEqualTo(XLError.DivisionByZero);
    }

    #endregion

    [Test]
    public async Task Distributions_EvaluateAgainstWorksheetCells()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 42;
            ws.Cell("A2").Value = 40;
            ws.Cell("A3").Value = 1.5;

            ws.Cell("B1").FormulaA1 = "NORM.DIST(A1, A2, A3, TRUE)";
            ws.Cell("B2").FormulaA1 = "NORM.INV(B1, A2, A3)";

            await Assert.That((double)ws.Cell("B1").Value).IsEqualTo(0.9087887802741321d).Within(1e-9);
            await Assert.That((double)ws.Cell("B2").Value).IsEqualTo(42d).Within(1e-9);
        }
    }
}
