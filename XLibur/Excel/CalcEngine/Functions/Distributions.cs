using System;
using System.Collections.Generic;
using XLibur.Excel.CalcEngine.Functions;
using static XLibur.Excel.CalcEngine.Functions.SampleStatistics;
using static XLibur.Excel.CalcEngine.Functions.SignatureAdapter;

#pragma warning disable S1244 // Intentional exact float comparison for Excel formula compatibility

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// The statistical distributions and the hypothesis tests built on them.
/// <para>
/// Excel carries two spellings of most of these: the modern dotted names (NORM.DIST) and the
/// pre-2010 ones (NORMDIST). Where the two compute the same thing the legacy name is registered
/// against the same implementation rather than a copy of it; where they differ — a legacy name
/// that only offers the cumulative form, or that returns the right tail where the dotted name
/// returns the left — the difference is in a thin wrapper, and the comment says which.
/// </para>
/// <para>
/// Every distribution here reduces to one of four special functions in <see cref="XLMath"/>: the
/// regularized incomplete gamma (chi-squared, gamma, Poisson), the regularized incomplete beta
/// (F, t, binomial), the error function (normal), or plain elementary functions.
/// </para>
/// </summary>
internal static class Distributions
{
    public static void Register(FunctionRegistry ce)
    {
        ce.RegisterFunction("BETA.DIST", 4, 6, AdaptLastTwoOptional(BetaDist, 0, 1), FunctionFlags.Scalar | FunctionFlags.Future); // Beta probability distribution
        ce.RegisterFunction("BETA.INV", 3, 5, AdaptLastTwoOptional(BetaInv, 0, 1), FunctionFlags.Scalar | FunctionFlags.Future); // Inverse of the beta cumulative distribution
        ce.RegisterFunction("BETADIST", 3, 5, AdaptLastTwoOptional(BetaDistLegacy, 0, 1), FunctionFlags.Scalar); // Cumulative beta distribution
        ce.RegisterFunction("BETAINV", 3, 5, AdaptLastTwoOptional(BetaInv, 0, 1), FunctionFlags.Scalar);
        ce.RegisterFunction("BINOM.INV", 3, 3, Adapt(BinomInv), FunctionFlags.Scalar | FunctionFlags.Future); // Smallest value whose binomial cumulative distribution reaches a criterion
        ce.RegisterFunction("CHISQ.DIST", 3, 3, AdaptLastOptional(ChiSqDist, true), FunctionFlags.Scalar | FunctionFlags.Future); // Left-tailed chi-squared distribution
        ce.RegisterFunction("CHISQ.DIST.RT", 2, 2, Adapt(ChiSqDistRt), FunctionFlags.Scalar | FunctionFlags.Future); // Right-tailed chi-squared distribution
        ce.RegisterFunction("CHISQ.INV", 2, 2, Adapt(ChiSqInv), FunctionFlags.Scalar | FunctionFlags.Future); // Inverse of the left-tailed chi-squared distribution
        ce.RegisterFunction("CHISQ.INV.RT", 2, 2, Adapt(ChiSqInvRt), FunctionFlags.Scalar | FunctionFlags.Future); // Inverse of the right-tailed chi-squared distribution
        ce.RegisterFunction("CHISQ.TEST", 2, 2, ChiSqTest, FunctionFlags.Range | FunctionFlags.Future, AllowRange.All); // Chi-squared test of independence
        ce.RegisterFunction("CHIDIST", 2, 2, Adapt(ChiSqDistRt), FunctionFlags.Scalar);
        ce.RegisterFunction("CHIINV", 2, 2, Adapt(ChiSqInvRt), FunctionFlags.Scalar);
        ce.RegisterFunction("CHITEST", 2, 2, ChiSqTest, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("CONFIDENCE", 3, 3, Adapt(ConfidenceNorm), FunctionFlags.Scalar);
        ce.RegisterFunction("CONFIDENCE.NORM", 3, 3, Adapt(ConfidenceNorm), FunctionFlags.Scalar | FunctionFlags.Future); // Confidence interval using the normal distribution
        ce.RegisterFunction("CONFIDENCE.T", 3, 3, Adapt(ConfidenceT), FunctionFlags.Scalar | FunctionFlags.Future); // Confidence interval using the Student's t-distribution
        ce.RegisterFunction("CRITBINOM", 3, 3, Adapt(BinomInv), FunctionFlags.Scalar);
        ce.RegisterFunction("EXPON.DIST", 3, 3, AdaptLastOptional(ExponDist, true), FunctionFlags.Scalar | FunctionFlags.Future); // Exponential distribution
        ce.RegisterFunction("EXPONDIST", 3, 3, AdaptLastOptional(ExponDist, true), FunctionFlags.Scalar);
        ce.RegisterFunction("F.DIST", 4, 4, Adapt(FDist), FunctionFlags.Scalar | FunctionFlags.Future); // Left-tailed F probability distribution
        ce.RegisterFunction("F.DIST.RT", 3, 3, Adapt(FDistRt), FunctionFlags.Scalar | FunctionFlags.Future); // Right-tailed F probability distribution
        ce.RegisterFunction("F.INV", 3, 3, Adapt(FInv), FunctionFlags.Scalar | FunctionFlags.Future); // Inverse of the left-tailed F distribution
        ce.RegisterFunction("F.INV.RT", 3, 3, Adapt(FInvRt), FunctionFlags.Scalar | FunctionFlags.Future); // Inverse of the right-tailed F distribution
        ce.RegisterFunction("F.TEST", 2, 2, FTest, FunctionFlags.Range | FunctionFlags.Future, AllowRange.All); // Two-tailed F-test of two variances
        ce.RegisterFunction("FDIST", 3, 3, Adapt(FDistRt), FunctionFlags.Scalar);
        ce.RegisterFunction("FINV", 3, 3, Adapt(FInvRt), FunctionFlags.Scalar);
        ce.RegisterFunction("FTEST", 2, 2, FTest, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("GAMMA", 1, 1, Adapt(Gamma), FunctionFlags.Scalar | FunctionFlags.Future); // The gamma function
        ce.RegisterFunction("GAMMA.DIST", 4, 4, Adapt(GammaDist), FunctionFlags.Scalar | FunctionFlags.Future); // Gamma distribution
        ce.RegisterFunction("GAMMA.INV", 3, 3, Adapt(GammaInv), FunctionFlags.Scalar | FunctionFlags.Future); // Inverse of the gamma cumulative distribution
        ce.RegisterFunction("GAMMADIST", 4, 4, Adapt(GammaDist), FunctionFlags.Scalar);
        ce.RegisterFunction("GAMMAINV", 3, 3, Adapt(GammaInv), FunctionFlags.Scalar);
        ce.RegisterFunction("GAMMALN", 1, 1, Adapt(GammaLn), FunctionFlags.Scalar); // Natural logarithm of the gamma function
        ce.RegisterFunction("GAMMALN.PRECISE", 1, 1, Adapt(GammaLn), FunctionFlags.Scalar | FunctionFlags.Future);
        ce.RegisterFunction("HYPGEOM.DIST", 5, 5, Adapt(HypGeomDist), FunctionFlags.Scalar | FunctionFlags.Future); // Hypergeometric distribution
        ce.RegisterFunction("HYPGEOMDIST", 4, 4, Adapt(HypGeomDistLegacy), FunctionFlags.Scalar);
        ce.RegisterFunction("LOGINV", 3, 3, Adapt(LogNormInv), FunctionFlags.Scalar);
        ce.RegisterFunction("LOGNORM.DIST", 4, 4, Adapt(LogNormDist), FunctionFlags.Scalar | FunctionFlags.Future); // Lognormal distribution
        ce.RegisterFunction("LOGNORM.INV", 3, 3, Adapt(LogNormInv), FunctionFlags.Scalar | FunctionFlags.Future); // Inverse of the lognormal cumulative distribution
        ce.RegisterFunction("LOGNORMDIST", 3, 3, Adapt(LogNormDistLegacy), FunctionFlags.Scalar);
        ce.RegisterFunction("NEGBINOM.DIST", 4, 4, Adapt(NegBinomDist), FunctionFlags.Scalar | FunctionFlags.Future); // Negative binomial distribution
        ce.RegisterFunction("NEGBINOMDIST", 3, 3, Adapt(NegBinomDistLegacy), FunctionFlags.Scalar);
        ce.RegisterFunction("NORM.DIST", 4, 4, Adapt(NormDist), FunctionFlags.Scalar | FunctionFlags.Future); // Normal distribution
        ce.RegisterFunction("NORM.INV", 3, 3, Adapt(NormInv), FunctionFlags.Scalar | FunctionFlags.Future); // Inverse of the normal cumulative distribution
        ce.RegisterFunction("NORM.S.DIST", 2, 2, Adapt(NormSDist), FunctionFlags.Scalar | FunctionFlags.Future); // Standard normal distribution
        ce.RegisterFunction("NORM.S.INV", 1, 1, Adapt(NormSInv), FunctionFlags.Scalar | FunctionFlags.Future); // Inverse of the standard normal cumulative distribution
        ce.RegisterFunction("NORMDIST", 4, 4, Adapt(NormDist), FunctionFlags.Scalar);
        ce.RegisterFunction("NORMINV", 3, 3, Adapt(NormInv), FunctionFlags.Scalar);
        ce.RegisterFunction("NORMSDIST", 1, 1, Adapt(NormSDistLegacy), FunctionFlags.Scalar);
        ce.RegisterFunction("NORMSINV", 1, 1, Adapt(NormSInv), FunctionFlags.Scalar);
        ce.RegisterFunction("POISSON", 3, 3, AdaptLastOptional(PoissonDist, true), FunctionFlags.Scalar);
        ce.RegisterFunction("POISSON.DIST", 3, 3, AdaptLastOptional(PoissonDist, true), FunctionFlags.Scalar | FunctionFlags.Future); // Poisson distribution
        ce.RegisterFunction("T.DIST", 3, 3, AdaptLastOptional(TDist, true), FunctionFlags.Scalar | FunctionFlags.Future); // Left-tailed Student's t-distribution
        ce.RegisterFunction("T.DIST.2T", 2, 2, Adapt(TDist2T), FunctionFlags.Scalar | FunctionFlags.Future); // Two-tailed Student's t-distribution
        ce.RegisterFunction("T.DIST.RT", 2, 2, Adapt(TDistRt), FunctionFlags.Scalar | FunctionFlags.Future); // Right-tailed Student's t-distribution
        ce.RegisterFunction("T.TEST", 4, 4, TTest, FunctionFlags.Range | FunctionFlags.Future, AllowRange.Only, 0, 1); // Probability associated with a Student's t-test
        ce.RegisterFunction("TDIST", 3, 3, Adapt(TDistLegacy), FunctionFlags.Scalar);
        ce.RegisterFunction("TTEST", 4, 4, TTest, FunctionFlags.Range, AllowRange.Only, 0, 1);
        ce.RegisterFunction("WEIBULL", 4, 4, Adapt(WeibullDist), FunctionFlags.Scalar);
        ce.RegisterFunction("WEIBULL.DIST", 4, 4, Adapt(WeibullDist), FunctionFlags.Scalar | FunctionFlags.Future); // Weibull distribution
        ce.RegisterFunction("Z.TEST", 2, 3, ZTest, FunctionFlags.Range | FunctionFlags.Future, AllowRange.Only, 0); // One-tailed probability value of a z-test
        ce.RegisterFunction("ZTEST", 2, 3, ZTest, FunctionFlags.Range, AllowRange.Only, 0);
    }

    #region Normal and lognormal

    private static ScalarValue NormDist(CalcContext ctx, double x, double mean, double standardDeviation, bool cumulative)
    {
        if (standardDeviation <= 0)
            return XLError.NumberInvalid;

        var z = (x - mean) / standardDeviation;
        return cumulative ? XLMath.NormalSDist(z) : XLMath.NormalSPdf(z) / standardDeviation;
    }

    private static ScalarValue NormInv(CalcContext ctx, double probability, double mean, double standardDeviation)
    {
        if (probability <= 0 || probability >= 1 || standardDeviation <= 0)
            return XLError.NumberInvalid;

        return mean + standardDeviation * XLMath.NormalSInv(probability);
    }

    private static ScalarValue NormSDist(CalcContext ctx, double z, bool cumulative)
        => cumulative ? XLMath.NormalSDist(z) : XLMath.NormalSPdf(z);

    /// <summary>The pre-2010 NORMSDIST has no cumulative flag — it is always the cumulative form.</summary>
    private static ScalarValue NormSDistLegacy(CalcContext ctx, double z) => XLMath.NormalSDist(z);

    private static ScalarValue NormSInv(CalcContext ctx, double probability)
    {
        if (probability <= 0 || probability >= 1)
            return XLError.NumberInvalid;

        return XLMath.NormalSInv(probability);
    }

    private static ScalarValue LogNormDist(CalcContext ctx, double x, double mean, double standardDeviation, bool cumulative)
    {
        if (x <= 0 || standardDeviation <= 0)
            return XLError.NumberInvalid;

        // A lognormal variable is the exponential of a normal one, so its distribution is the
        // normal distribution of the logarithm — with the density divided by x, the derivative of
        // that change of variable.
        var z = (Math.Log(x) - mean) / standardDeviation;
        return cumulative ? XLMath.NormalSDist(z) : XLMath.NormalSPdf(z) / (x * standardDeviation);
    }

    private static ScalarValue LogNormDistLegacy(CalcContext ctx, double x, double mean, double standardDeviation)
        => LogNormDist(ctx, x, mean, standardDeviation, cumulative: true);

    private static ScalarValue LogNormInv(CalcContext ctx, double probability, double mean, double standardDeviation)
    {
        if (probability <= 0 || probability >= 1 || standardDeviation <= 0)
            return XLError.NumberInvalid;

        return Math.Exp(mean + standardDeviation * XLMath.NormalSInv(probability));
    }

    #endregion

    #region Chi-squared

    private static ScalarValue ChiSqDist(CalcContext ctx, double x, double degreesOfFreedom, bool cumulative)
    {
        if (x < 0 || !IsValidDegreesOfFreedom(degreesOfFreedom))
            return XLError.NumberInvalid;

        var df = Math.Truncate(degreesOfFreedom);
        if (cumulative)
            return XLMath.GammaP(df / 2, x / 2);

        // At the origin the density diverges below two degrees of freedom, is exactly a half at
        // two, and vanishes above.
        if (x == 0)
        {
            if (df < 2)
                return XLError.NumberInvalid;

            return df == 2 ? 0.5 : 0d;
        }

        var logDensity = (df / 2 - 1) * Math.Log(x) - x / 2 - df / 2 * Math.Log(2) - XLMath.LnGamma(df / 2);
        return Math.Exp(logDensity);
    }

    private static ScalarValue ChiSqDistRt(CalcContext ctx, double x, double degreesOfFreedom)
    {
        if (x < 0 || !IsValidDegreesOfFreedom(degreesOfFreedom))
            return XLError.NumberInvalid;

        return XLMath.GammaQ(Math.Truncate(degreesOfFreedom) / 2, x / 2);
    }

    private static ScalarValue ChiSqInv(CalcContext ctx, double probability, double degreesOfFreedom)
    {
        if (probability < 0 || probability > 1 || !IsValidDegreesOfFreedom(degreesOfFreedom))
            return XLError.NumberInvalid;

        // Chi-squared with k degrees of freedom is gamma with shape k/2 and scale 2.
        return 2 * XLMath.InverseGammaP(probability, Math.Truncate(degreesOfFreedom) / 2);
    }

    private static ScalarValue ChiSqInvRt(CalcContext ctx, double probability, double degreesOfFreedom)
    {
        if (probability < 0 || probability > 1 || !IsValidDegreesOfFreedom(degreesOfFreedom))
            return XLError.NumberInvalid;

        return 2 * XLMath.InverseGammaP(1 - probability, Math.Truncate(degreesOfFreedom) / 2);
    }

    /// <summary>
    /// CHISQ.TEST(actual, expected) — the chi-squared statistic Σ(a−e)²/e turned into the
    /// probability of seeing one at least that large by chance. The degrees of freedom come from
    /// the shape of the ranges: a two-dimensional table has (rows−1)(columns−1), a single row or
    /// column has one less than its length.
    /// </summary>
#pragma warning disable S3776 // A double loop over paired cells; the guards inside it are the function's error contract
    private static AnyValue ChiSqTest(CalcContext ctx, Span<AnyValue> args)
    {
        if (!args[0].TryPickCollectionArray(out var actual, ctx) ||
            !args[1].TryPickCollectionArray(out var expected, ctx))
        {
            return XLError.IncompatibleValue;
        }

        if (actual!.Height != expected!.Height || actual.Width != expected.Width)
            return XLError.NoValueAvailable;

        var statistic = 0d;
        var count = 0;
        for (var row = 0; row < actual.Height; row++)
        {
            for (var column = 0; column < actual.Width; column++)
            {
                if (!TryPairOfNumbers(ctx, actual[row, column], expected[row, column], out var a, out var e, out var error))
                    return error;

                if (e == 0)
                    return XLError.DivisionByZero;

                statistic += (a - e) * (a - e) / e;
                count++;
            }
        }

        var degreesOfFreedom = actual.Height > 1 && actual.Width > 1
            ? (actual.Height - 1) * (actual.Width - 1)
            : count - 1;

        if (degreesOfFreedom < 1)
            return XLError.NoValueAvailable;

        return XLMath.GammaQ(degreesOfFreedom / 2.0, statistic / 2);
    }
#pragma warning restore S3776

    private static bool IsValidDegreesOfFreedom(double degreesOfFreedom)
        => degreesOfFreedom >= 1 && degreesOfFreedom < 1e10;

    #endregion

    #region F distribution

    private static ScalarValue FDist(CalcContext ctx, double x, double degreesOfFreedom1, double degreesOfFreedom2, bool cumulative)
    {
        if (x < 0 || !IsValidDegreesOfFreedom(degreesOfFreedom1) || !IsValidDegreesOfFreedom(degreesOfFreedom2))
            return XLError.NumberInvalid;

        var d1 = Math.Truncate(degreesOfFreedom1);
        var d2 = Math.Truncate(degreesOfFreedom2);

        if (cumulative)
            return XLMath.BetaRegularized(d1 * x / (d1 * x + d2), d1 / 2, d2 / 2);

        // As for chi-squared, the origin is a special case governed by the numerator's degrees of
        // freedom alone.
        if (x == 0)
        {
            if (d1 < 2)
                return XLError.NumberInvalid;

            return d1 == 2 ? 1d : 0d;
        }

        var logDensity = d1 / 2 * Math.Log(d1) + d2 / 2 * Math.Log(d2)
            + (d1 / 2 - 1) * Math.Log(x)
            - (d1 + d2) / 2 * Math.Log(d2 + d1 * x)
            - (XLMath.LnGamma(d1 / 2) + XLMath.LnGamma(d2 / 2) - XLMath.LnGamma((d1 + d2) / 2));
        return Math.Exp(logDensity);
    }

    private static ScalarValue FDistRt(CalcContext ctx, double x, double degreesOfFreedom1, double degreesOfFreedom2)
    {
        if (x < 0 || !IsValidDegreesOfFreedom(degreesOfFreedom1) || !IsValidDegreesOfFreedom(degreesOfFreedom2))
            return XLError.NumberInvalid;

        var d1 = Math.Truncate(degreesOfFreedom1);
        var d2 = Math.Truncate(degreesOfFreedom2);

        // Taking the complementary tail through the beta function's own symmetry keeps the small
        // p-values that matter for a test from being lost to 1 − (something very close to 1).
        return XLMath.BetaRegularized(d2 / (d2 + d1 * x), d2 / 2, d1 / 2);
    }

    private static ScalarValue FInv(CalcContext ctx, double probability, double degreesOfFreedom1, double degreesOfFreedom2)
    {
        if (probability < 0 || probability > 1 || !IsValidDegreesOfFreedom(degreesOfFreedom1) || !IsValidDegreesOfFreedom(degreesOfFreedom2))
            return XLError.NumberInvalid;

        return InvertF(probability, Math.Truncate(degreesOfFreedom1), Math.Truncate(degreesOfFreedom2));
    }

    private static ScalarValue FInvRt(CalcContext ctx, double probability, double degreesOfFreedom1, double degreesOfFreedom2)
    {
        if (probability < 0 || probability > 1 || !IsValidDegreesOfFreedom(degreesOfFreedom1) || !IsValidDegreesOfFreedom(degreesOfFreedom2))
            return XLError.NumberInvalid;

        return InvertF(1 - probability, Math.Truncate(degreesOfFreedom1), Math.Truncate(degreesOfFreedom2));
    }

    /// <summary>
    /// Invert the F CDF by inverting the beta function it is built from: if
    /// <c>y = I⁻¹(p; d1/2, d2/2)</c> then <c>d1·x/(d1·x + d2) = y</c>, which rearranges to x.
    /// </summary>
    private static double InvertF(double probability, double d1, double d2)
    {
        var y = XLMath.InverseBetaRegularized(probability, d1 / 2, d2 / 2);
        if (y >= 1)
            return double.PositiveInfinity;

        return d2 * y / (d1 * (1 - y));
    }

    /// <summary>
    /// F.TEST(array1, array2) — the two-tailed probability that two samples come from populations
    /// with the same variance. The ratio is always taken with the larger variance on top, and the
    /// one-tailed probability of that is doubled.
    /// </summary>
    private static AnyValue FTest(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryGetSample(ctx, args[0], out var first, out var firstError))
            return firstError;
        if (!TryGetSample(ctx, args[1], out var second, out var secondError))
            return secondError;

        if (first.Count < 2 || second.Count < 2)
            return XLError.DivisionByZero;

        var varianceA = SampleVariance(first);
        var varianceB = SampleVariance(second);
        if (varianceA == 0 || varianceB == 0)
            return XLError.DivisionByZero;

        double ratio, dfNumerator, dfDenominator;
        if (varianceA > varianceB)
        {
            ratio = varianceA / varianceB;
            dfNumerator = first.Count - 1;
            dfDenominator = second.Count - 1;
        }
        else
        {
            ratio = varianceB / varianceA;
            dfNumerator = second.Count - 1;
            dfDenominator = first.Count - 1;
        }

        var oneTailed = XLMath.BetaRegularized(dfDenominator / (dfDenominator + dfNumerator * ratio), dfDenominator / 2, dfNumerator / 2);
        return Math.Min(2 * oneTailed, 1);
    }

    #endregion

    #region Student's t

    private static ScalarValue TDist(CalcContext ctx, double x, double degreesOfFreedom, bool cumulative)
    {
        if (!IsValidDegreesOfFreedom(degreesOfFreedom))
            return XLError.NumberInvalid;

        var df = Math.Truncate(degreesOfFreedom);
        if (cumulative)
            return StudentTCdf(x, df);

        var logDensity = -0.5 * Math.Log(df * Math.PI) + XLMath.LnGamma((df + 1) / 2) - XLMath.LnGamma(df / 2)
            - (df + 1) / 2 * Math.Log(1 + x * x / df);
        return Math.Exp(logDensity);
    }

    private static ScalarValue TDistRt(CalcContext ctx, double x, double degreesOfFreedom)
    {
        if (!IsValidDegreesOfFreedom(degreesOfFreedom))
            return XLError.NumberInvalid;

        return 1 - StudentTCdf(x, Math.Truncate(degreesOfFreedom));
    }

    private static ScalarValue TDist2T(CalcContext ctx, double x, double degreesOfFreedom)
    {
        if (x < 0 || !IsValidDegreesOfFreedom(degreesOfFreedom))
            return XLError.NumberInvalid;

        return Math.Min(2 * (1 - StudentTCdf(x, Math.Truncate(degreesOfFreedom))), 1);
    }

    /// <summary>The pre-2010 TDIST takes the number of tails instead of offering a left tail at all.</summary>
    private static ScalarValue TDistLegacy(CalcContext ctx, double x, double degreesOfFreedom, double tails)
    {
        var tailCount = Math.Truncate(tails);
        if (x < 0 || !IsValidDegreesOfFreedom(degreesOfFreedom) || tailCount is not (1 or 2))
            return XLError.NumberInvalid;

        var rightTail = 1 - StudentTCdf(x, Math.Truncate(degreesOfFreedom));
        return Math.Min(tailCount * rightTail, 1);
    }

    /// <summary>
    /// CDF of the t-distribution, through the regularized incomplete beta function. Written so the
    /// small tail is the one computed directly, which is where a test's p-value lives.
    /// </summary>
    private static double StudentTCdf(double t, double df)
    {
        var x = df / (df + t * t);
        var half = 0.5 * XLMath.BetaRegularized(x, df / 2, 0.5);
        return t >= 0 ? 1 - half : half;
    }

    /// <summary>
    /// T.TEST(array1, array2, tails, type) — type 1 pairs the observations, type 2 assumes the two
    /// populations share a variance, and type 3 (Welch) does not and adjusts the degrees of freedom
    /// to match.
    /// </summary>
    private static AnyValue TTest(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryGetSample(ctx, args[0], out var first, out var firstError))
            return firstError;
        if (!TryGetSample(ctx, args[1], out var second, out var secondError))
            return secondError;
        if (!TryGetScalarNumber(ctx, args[2], out var tailsValue, out var tailsError))
            return tailsError;
        if (!TryGetScalarNumber(ctx, args[3], out var typeValue, out var typeError))
            return typeError;

        var tails = Math.Truncate(tailsValue);
        var type = Math.Truncate(typeValue);
        if (tails is not (1 or 2) || type is not (1 or 2 or 3))
            return XLError.NumberInvalid;

        var statistic = type switch
        {
            1 => PairedStatistic(first, second),
            2 => PooledStatistic(first, second),
            _ => WelchStatistic(first, second),
        };

        if (!statistic.TryPickT0(out var test, out var statisticError))
            return statisticError;

        var rightTail = 1 - StudentTCdf(Math.Abs(test.T), test.DegreesOfFreedom);
        return Math.Min(tails * rightTail, 1);
    }

    /// <summary>A t statistic and the degrees of freedom to read it against.</summary>
    private readonly record struct TStatistic(double T, double DegreesOfFreedom);

    /// <summary>Type 1 — the observations are paired, so the test is a one-sample test of their differences.</summary>
    private static OneOf<TStatistic, XLError> PairedStatistic(List<double> first, List<double> second)
    {
        if (first.Count != second.Count)
            return XLError.NoValueAvailable;
        if (first.Count < 2)
            return XLError.DivisionByZero;

        var differences = new List<double>(first.Count);
        for (var i = 0; i < first.Count; i++)
            differences.Add(first[i] - second[i]);

        var standardError = Math.Sqrt(SampleVariance(differences) / differences.Count);
        if (standardError == 0)
            return XLError.DivisionByZero;

        return new TStatistic(Mean(differences) / standardError, differences.Count - 1);
    }

    /// <summary>Type 2 — the two populations are assumed to share a variance, which is pooled from both samples.</summary>
    private static OneOf<TStatistic, XLError> PooledStatistic(List<double> first, List<double> second)
    {
        if (first.Count < 2 || second.Count < 2)
            return XLError.DivisionByZero;

        var n1 = first.Count;
        var n2 = second.Count;
        var pooled = ((n1 - 1) * SampleVariance(first) + (n2 - 1) * SampleVariance(second)) / (n1 + n2 - 2);
        var standardError = Math.Sqrt(pooled * (1.0 / n1 + 1.0 / n2));
        if (standardError == 0)
            return XLError.DivisionByZero;

        return new TStatistic((Mean(first) - Mean(second)) / standardError, n1 + n2 - 2);
    }

    /// <summary>
    /// Type 3 — the variances are not assumed equal, so each sample contributes its own and the
    /// degrees of freedom are adjusted to match.
    /// </summary>
    private static OneOf<TStatistic, XLError> WelchStatistic(List<double> first, List<double> second)
    {
        if (first.Count < 2 || second.Count < 2)
            return XLError.DivisionByZero;

        var a = SampleVariance(first) / first.Count;
        var b = SampleVariance(second) / second.Count;
        if (a + b == 0)
            return XLError.DivisionByZero;

        var t = (Mean(first) - Mean(second)) / Math.Sqrt(a + b);

        // Welch–Satterthwaite: the variance of the difference behaves like a chi-squared with this
        // many degrees of freedom, which is generally not a whole number. Excel truncates it, as
        // every other function in the T.DIST family truncates its own — which is what lets
        // T.TEST(…, 3) be reproduced from T.DIST.2T and the Welch formula in a worksheet.
        var degreesOfFreedom = Math.Truncate(
            (a + b) * (a + b) / (a * a / (first.Count - 1) + b * b / (second.Count - 1)));

        return new TStatistic(t, degreesOfFreedom);
    }

    /// <summary>
    /// Z.TEST(array, x, [sigma]) — the one-tailed probability that the sample mean is as far above
    /// <c>x</c> as it is. With no sigma given the sample standard deviation stands in for it.
    /// </summary>
    private static AnyValue ZTest(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryGetSample(ctx, args[0], out var sample, out var sampleError))
            return sampleError;
        if (!TryGetScalarNumber(ctx, args[1], out var hypothesizedMean, out var meanError))
            return meanError;

        if (sample.Count < 1)
            return XLError.NoValueAvailable;

        var sigma = 0d;
        if (args.Length > 2)
        {
            if (!TryGetScalarNumber(ctx, args[2], out sigma, out var sigmaError))
                return sigmaError;
            if (sigma <= 0)
                return XLError.NumberInvalid;
        }
        else
        {
            if (sample.Count < 2)
                return XLError.DivisionByZero;

            sigma = Math.Sqrt(SampleVariance(sample));
            if (sigma == 0)
                return XLError.DivisionByZero;
        }

        var z = (Mean(sample) - hypothesizedMean) / (sigma / Math.Sqrt(sample.Count));
        return 1 - XLMath.NormalSDist(z);
    }

    #endregion

    #region Exponential families

    private static ScalarValue ExponDist(CalcContext ctx, double x, double lambda, bool cumulative)
    {
        if (x < 0 || lambda <= 0)
            return XLError.NumberInvalid;

        return cumulative ? 1 - Math.Exp(-lambda * x) : lambda * Math.Exp(-lambda * x);
    }

    private static ScalarValue PoissonDist(CalcContext ctx, double x, double mean, bool cumulative)
    {
        var k = Math.Truncate(x);
        if (k < 0 || mean < 0)
            return XLError.NumberInvalid;

        // The Poisson CDF is the regularized upper incomplete gamma, which stays accurate for a
        // large mean where summing the terms one at a time would not.
        if (cumulative)
            return mean == 0 ? 1d : XLMath.GammaQ(k + 1, mean);

        return Math.Exp(-mean + k * Math.Log(mean == 0 ? 1 : mean) - XLMath.LnGamma(k + 1)) * (mean == 0 && k > 0 ? 0 : 1);
    }

    private static ScalarValue WeibullDist(CalcContext ctx, double x, double alpha, double beta, bool cumulative)
    {
        if (x < 0 || alpha <= 0 || beta <= 0)
            return XLError.NumberInvalid;

        var scaled = Math.Pow(x / beta, alpha);
        if (cumulative)
            return 1 - Math.Exp(-scaled);

        return alpha / Math.Pow(beta, alpha) * Math.Pow(x, alpha - 1) * Math.Exp(-scaled);
    }

    private static ScalarValue GammaDist(CalcContext ctx, double x, double alpha, double beta, bool cumulative)
    {
        if (x < 0 || alpha <= 0 || beta <= 0)
            return XLError.NumberInvalid;

        if (cumulative)
            return XLMath.GammaP(alpha, x / beta);

        if (x == 0)
            return alpha switch
            {
                < 1 => XLError.NumberInvalid,
                1 => 1 / beta,
                _ => 0d,
            };

        var logDensity = (alpha - 1) * Math.Log(x) - x / beta - alpha * Math.Log(beta) - XLMath.LnGamma(alpha);
        return Math.Exp(logDensity);
    }

    private static ScalarValue GammaInv(CalcContext ctx, double probability, double alpha, double beta)
    {
        if (probability < 0 || probability > 1 || alpha <= 0 || beta <= 0)
            return XLError.NumberInvalid;

        return beta * XLMath.InverseGammaP(probability, alpha);
    }

    private static ScalarValue Gamma(CalcContext ctx, double x)
    {
        // The gamma function has a pole at every non-positive integer.
        if (x <= 0 && x == Math.Truncate(x))
            return XLError.NumberInvalid;

        var result = XLMath.Gamma(x);
        return double.IsNaN(result) || double.IsInfinity(result) ? XLError.NumberInvalid : result;
    }

    private static ScalarValue GammaLn(CalcContext ctx, double x)
    {
        if (x <= 0)
            return XLError.NumberInvalid;

        return XLMath.LnGamma(x);
    }

    #endregion

    #region Beta

    private static ScalarValue BetaDist(CalcContext ctx, double x, double alpha, double beta, bool cumulative, double lowerBound, double upperBound)
    {
        if (alpha <= 0 || beta <= 0 || upperBound <= lowerBound || x < lowerBound || x > upperBound)
            return XLError.NumberInvalid;

        var width = upperBound - lowerBound;
        var scaled = (x - lowerBound) / width;

        if (cumulative)
            return XLMath.BetaRegularized(scaled, alpha, beta);

        if (scaled is 0 or 1)
        {
            // The density is finite at an endpoint only when the matching shape parameter allows it.
            var shape = scaled == 0 ? alpha : beta;
            if (shape < 1)
                return XLError.NumberInvalid;
        }

        var logDensity = (alpha - 1) * Math.Log(scaled) + (beta - 1) * Math.Log(1 - scaled)
            - (XLMath.LnGamma(alpha) + XLMath.LnGamma(beta) - XLMath.LnGamma(alpha + beta));
        return Math.Exp(logDensity) / width;
    }

    /// <summary>The pre-2010 BETADIST offers only the cumulative form.</summary>
    private static ScalarValue BetaDistLegacy(CalcContext ctx, double x, double alpha, double beta, double lowerBound, double upperBound)
        => BetaDist(ctx, x, alpha, beta, cumulative: true, lowerBound, upperBound);

    private static AnyValue BetaInv(double probability, double alpha, double beta, double lowerBound, double upperBound)
    {
        if (probability <= 0 || probability > 1 || alpha <= 0 || beta <= 0 || upperBound <= lowerBound)
            return XLError.NumberInvalid;

        return lowerBound + (upperBound - lowerBound) * XLMath.InverseBetaRegularized(probability, alpha, beta);
    }

    #endregion

    #region Discrete distributions

    private static ScalarValue HypGeomDist(CalcContext ctx, double sampleSuccesses, double sampleSize, double populationSuccesses, double populationSize, bool cumulative)
    {
        var x = Math.Truncate(sampleSuccesses);
        var n = Math.Truncate(sampleSize);
        var successes = Math.Truncate(populationSuccesses);
        var population = Math.Truncate(populationSize);

        if (n <= 0 || n > population || successes <= 0 || successes > population || population <= 0)
            return XLError.NumberInvalid;

        // The number of successes drawn cannot exceed the sample, the successes available, or fall
        // below what the failures in the population force.
        if (x < 0 || x > n || x > successes || x < n - (population - successes))
            return XLError.NumberInvalid;

        if (!cumulative)
            return HypGeomTerm(x, n, successes, population);

        var total = 0d;
        var from = Math.Max(0, n - (population - successes));
        for (var k = from; k <= x; k++)
            total += HypGeomTerm(k, n, successes, population);

        return Math.Min(total, 1);
    }

    private static ScalarValue HypGeomDistLegacy(CalcContext ctx, double sampleSuccesses, double sampleSize, double populationSuccesses, double populationSize)
        => HypGeomDist(ctx, sampleSuccesses, sampleSize, populationSuccesses, populationSize, cumulative: false);

    /// <summary>
    /// One term of the hypergeometric distribution, C(M,k)·C(N−M,n−k)/C(N,n), evaluated through log
    /// factorials so that a large population does not overflow on the way to a small probability.
    /// </summary>
    private static double HypGeomTerm(double k, double n, double successes, double population)
    {
        var logProbability = LogChoose(successes, k)
            + LogChoose(population - successes, n - k)
            - LogChoose(population, n);
        return Math.Exp(logProbability);
    }

    private static double LogChoose(double n, double k)
        => XLMath.LnGamma(n + 1) - XLMath.LnGamma(k + 1) - XLMath.LnGamma(n - k + 1);

    private static AnyValue NegBinomDist(double failures, double successes, double probability, bool cumulative)
    {
        var f = Math.Truncate(failures);
        var s = Math.Truncate(successes);
        if (f < 0 || s < 1 || probability <= 0 || probability > 1)
            return XLError.NumberInvalid;

        if (!cumulative)
            return NegBinomTerm(f, s, probability);

        var total = 0d;
        for (var k = 0d; k <= f; k++)
            total += NegBinomTerm(k, s, probability);

        return Math.Min(total, 1);
    }

    private static ScalarValue NegBinomDistLegacy(CalcContext ctx, double failures, double successes, double probability)
    {
        var result = NegBinomDist(failures, successes, probability, cumulative: false);
        return result.TryPickScalar(out var scalar, out _) ? scalar : XLError.NumberInvalid;
    }

    private static double NegBinomTerm(double failures, double successes, double probability)
    {
        var logProbability = LogChoose(failures + successes - 1, successes - 1)
            + successes * Math.Log(probability)
            + failures * Math.Log(1 - probability);
        return Math.Exp(logProbability);
    }

    /// <summary>
    /// BINOM.INV(trials, probability, alpha) — the smallest number of successes whose cumulative
    /// binomial probability reaches <paramref name="alpha"/>.
    /// </summary>
    private static ScalarValue BinomInv(CalcContext ctx, double trials, double probability, double alpha)
    {
        var n = Math.Truncate(trials);
        if (n < 0 || probability < 0 || probability > 1 || alpha < 0 || alpha > 1)
            return XLError.NumberInvalid;

        var cumulative = 0d;
        for (var k = 0d; k <= n; k++)
        {
            cumulative += Math.Exp(LogChoose(n, k) + k * SafeLog(probability) + (n - k) * SafeLog(1 - probability));
            if (cumulative >= alpha)
                return k;
        }

        return n;
    }

    /// <summary>Logarithm that returns negative infinity at zero rather than throwing off the sum.</summary>
    private static double SafeLog(double value) => value <= 0 ? double.NegativeInfinity : Math.Log(value);

    #endregion

    #region Confidence intervals

    private static ScalarValue ConfidenceNorm(CalcContext ctx, double alpha, double standardDeviation, double size)
    {
        var n = Math.Truncate(size);
        if (alpha <= 0 || alpha >= 1 || standardDeviation <= 0 || n < 1)
            return XLError.NumberInvalid;

        return XLMath.NormalSInv(1 - alpha / 2) * standardDeviation / Math.Sqrt(n);
    }

    private static ScalarValue ConfidenceT(CalcContext ctx, double alpha, double standardDeviation, double size)
    {
        var n = Math.Truncate(size);
        if (alpha <= 0 || alpha >= 1 || standardDeviation <= 0 || n < 1)
            return XLError.NumberInvalid;

        // With a single observation there is no spread to estimate from.
        if (n == 1)
            return XLError.DivisionByZero;

        return XLMath.StudentTInv(1 - alpha / 2, n - 1) * standardDeviation / Math.Sqrt(n);
    }

    #endregion

    #region Sample helpers

    private static bool TryGetSample(CalcContext ctx, in AnyValue value, out List<double> numbers, out XLError error)
        => Statistical.TryGetNumbers(ctx, value, out numbers, out error);

    private static bool TryPairOfNumbers(CalcContext ctx, in ScalarValue left, in ScalarValue right, out double a, out double b, out XLError error)
    {
        a = 0;
        b = 0;
        if (!left.ToNumber(ctx.Culture).TryPickT0(out a, out error))
            return false;

        return right.ToNumber(ctx.Culture).TryPickT0(out b, out error);
    }

    #endregion
}
