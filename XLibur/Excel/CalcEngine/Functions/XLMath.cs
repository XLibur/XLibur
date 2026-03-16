using System;

#pragma warning disable S1244 // Intentional exact float comparison for Excel formula compatibility

namespace XLibur.Excel.CalcEngine.Functions;

public static class XLMath
{
    public static double DegreesToRadians(double degrees)
    {
        return (Math.PI / 180.0) * degrees;
    }

    public static double RadiansToDegrees(double radians)
    {
        return (180.0 / Math.PI) * radians;
    }

    public static double GradsToRadians(double grads)
    {
        return (grads / 200.0) * Math.PI;
    }

    public static double RadiansToGrads(double radians)
    {
        return (radians / Math.PI) * 200.0;
    }

    public static double DegreesToGrads(double degrees)
    {
        return (degrees / 9.0) * 10.0;
    }

    public static double GradsToDegrees(double grads)
    {
        return (grads / 10.0) * 9.0;
    }

    public static double Asinh(double x)
    {
        return (Math.Log(x + Math.Sqrt(x * x + 1.0)));
    }

    public static double ACosh(double x)
    {
        return (Math.Log(x + Math.Sqrt((x * x) - 1.0)));
    }

    public static double ATanh(double x)
    {
        return (Math.Log((1.0 + x) / (1.0 - x)) / 2.0);
    }

    public static double ACoth(double x)
    {
        //return (Math.Log((x + 1.0) / (x - 1.0)) / 2.0);
        return (ATanh(1.0 / x));
    }

    public static double ASech(double x)
    {
        return (ACosh(1.0 / x));
    }

    public static double ACsch(double x)
    {
        return (Asinh(1.0 / x));
    }

    public static double Sech(double x)
    {
        return (1.0 / Math.Cosh(x));
    }

    public static double Csch(double x)
    {
        return (1.0 / Math.Sinh(x));
    }

    public static double Coth(double x)
    {
        return (Math.Cosh(x) / Math.Sinh(x));
    }

    internal static OneOf<double, XLError> CombinChecked(double number, double numberChosen)
    {
        if (number < 0 || numberChosen < 0)
            return XLError.NumberInvalid;

        var n = Math.Floor(number);
        var k = Math.Floor(numberChosen);

        // Parameter doesn't fit into int. That's how many multiplications Excel allows.
        if (n >= int.MaxValue || k >= int.MaxValue)
            return XLError.NumberInvalid;

        if (n < k)
            return XLError.NumberInvalid;

        var combinations = Combin(n, k);
        if (double.IsInfinity(combinations) || double.IsNaN(combinations))
            return XLError.NumberInvalid;

        return combinations;
    }

    internal static double Combin(double n, double k)
    {
        if (k == 0) return 1;

        // Don't use recursion, malicious input could exhaust stack.
        // Don't calculate directly from factorials, could overflow.
        double result = 1;
        for (var i = 1; i <= k; i++, n--)
        {
            result *= n;
            result /= i;
        }

        return result;
    }

    internal static double Factorial(double n)
    {
        n = Math.Truncate(n);
        var factorial = 1d;
        while (n > 1)
        {
            factorial *= n--;

            // n can be very large, stop when we reach infinity.
            if (double.IsInfinity(factorial))
                return factorial;
        }

        return factorial;
    }

    public static bool IsEven(int value)
    {
        return Math.Abs(value % 2) == 0;
    }

    public static bool IsEven(double value)
    {
        // Check the number doesn't have any fractions and that it is even.
        // Due to rounding after division, only checking for % 2 could fail
        // for numbers really close to whole number.
        var hasNoFraction = value % 1 == 0;
        var isEven = value % 2 == 0;
        return hasNoFraction && isEven;
    }

    public static bool IsOdd(int value)
    {
        return Math.Abs(value % 2) != 0;
    }

    public static bool IsOdd(double value)
    {
        var hasNoFraction = value % 1 == 0;
        var isOdd = value % 2 != 0;
        return hasNoFraction && isOdd;
    }

    public static double Round(double value, double digits)
    {
        digits = Math.Truncate(digits);
        if (digits < 0)
        {
            var coef = Math.Pow(10, Math.Abs(digits));
            var shifted = value / coef;
            shifted = Math.Round(shifted, 0, MidpointRounding.AwayFromZero);

            // if coef is infinity
            if (shifted == 0)
                return 0;

            return shifted * coef;
        }

        // Double can store at most 15 digits and anything below that is float artefact
        return Math.Round(value, (int)Math.Min(digits, 15), MidpointRounding.AwayFromZero);
    }

    #region Statistical distribution helpers

    /// <summary>
    /// Natural logarithm of the gamma function using the Lanczos approximation (g=7, n=9).
    /// </summary>
    internal static double LnGamma(double x)
    {
        if (x <= 0 && x == Math.Floor(x))
            return double.PositiveInfinity;

        // Lanczos coefficients for g=7, n=9
        ReadOnlySpan<double> c = [
            0.99999999999980993,
            676.5203681218851,
            -1259.1392167224028,
            771.32342877765313,
            -176.61502916214059,
            12.507343278686905,
            -0.13857109526572012,
            9.9843695780195716e-6,
            1.5056327351493116e-7
        ];

        if (x < 0.5)
        {
            // Reflection formula: Gamma(x) * Gamma(1-x) = pi / sin(pi*x)
            return Math.Log(Math.PI / Math.Sin(Math.PI * x)) - LnGamma(1.0 - x);
        }

        x -= 1.0;
        var a = c[0];
        var t = x + 7.5; // x + g + 0.5
        for (var i = 1; i < 9; i++)
            a += c[i] / (x + i);

        return 0.5 * Math.Log(2.0 * Math.PI) + (x + 0.5) * Math.Log(t) - t + Math.Log(a);
    }

    /// <summary>
    /// Regularized incomplete beta function I_x(a, b) using the continued fraction expansion.
    /// </summary>
    internal static double BetaRegularized(double x, double a, double b)
    {
        if (x <= 0.0) return 0.0;
        if (x >= 1.0) return 1.0;

        // Use the symmetry relation when x > (a+1)/(a+b+2) for better convergence.
        if (x > (a + 1.0) / (a + b + 2.0))
#pragma warning disable S2234
            return 1.0 - BetaRegularized(1.0 - x, b, a);
#pragma warning restore S2234        

        // Compute the log of the prefix: x^a * (1-x)^b / (a * Beta(a,b))
        var lnPrefix = a * Math.Log(x) + b * Math.Log(1.0 - x)
                       - Math.Log(a) - LnBeta(a, b);

        var prefix = Math.Exp(lnPrefix);

        // Evaluate the continued fraction using the modified Lentz method.
        return prefix * BetaContinuedFraction(x, a, b);
    }

    /// <summary>
    /// Log of the beta function: ln(Beta(a,b)) = ln(Gamma(a)) + ln(Gamma(b)) - ln(Gamma(a+b)).
    /// </summary>
    private static double LnBeta(double a, double b)
    {
        return LnGamma(a) + LnGamma(b) - LnGamma(a + b);
    }

    /// <summary>
    /// Continued fraction for the regularized incomplete beta function (Lentz's method).
    /// </summary>
    private static double BetaContinuedFraction(double x, double a, double b)
    {
        const int maxIterations = 200;
        const double epsilon = 1e-15;
        const double tiny = 1e-30;

        var f = 1.0;
        var c = 1.0;
        var d = 1.0 - (a + b) * x / (a + 1.0);
        if (Math.Abs(d) < tiny) d = tiny;
        d = 1.0 / d;
        f = d;

        for (var m = 1; m <= maxIterations; m++)
        {
            // Even step: d_{2m}
            var m2 = 2 * m;
            var numerator = m * (b - m) * x / ((a + m2 - 1.0) * (a + m2));

            d = 1.0 + numerator * d;
            if (Math.Abs(d) < tiny) d = tiny;
            d = 1.0 / d;

            c = 1.0 + numerator / c;
            if (Math.Abs(c) < tiny) c = tiny;

            f *= c * d;

            // Odd step: d_{2m+1}
            numerator = -(a + m) * (a + b + m) * x / ((a + m2) * (a + m2 + 1.0));

            d = 1.0 + numerator * d;
            if (Math.Abs(d) < tiny) d = tiny;
            d = 1.0 / d;

            c = 1.0 + numerator / c;
            if (Math.Abs(c) < tiny) c = tiny;

            var delta = c * d;
            f *= delta;

            if (Math.Abs(delta - 1.0) < epsilon)
                break;
        }

        return f;
    }

    /// <summary>
    /// Inverse of the regularized incomplete beta function.
    /// Given p = I_x(a, b), finds x such that BetaRegularized(x, a, b) = p.
    /// Uses an initial approximation followed by Newton's method with bisection fallback.
    /// </summary>
    internal static double InverseBetaRegularized(double p, double a, double b)
    {
        if (p <= 0.0) return 0.0;
        if (p >= 1.0) return 1.0;

        // Use symmetry to keep p <= 0.5 for better numerical behavior.
        if (p > 0.5)
#pragma warning disable S2234
            return 1.0 - InverseBetaRegularized(1.0 - p, b, a);
#pragma warning restore S2234        

        var x = InverseBetaInitialGuess(p, a, b);
        x = Math.Max(1e-14, Math.Min(1.0 - 1e-14, x));

        return InverseBetaNewtonRefine(x, p, a, b);
    }

    private static double InverseBetaInitialGuess(double p, double a, double b)
    {
        if (a >= 1.0 && b >= 1.0)
        {
            var t = Math.Sqrt(-2.0 * Math.Log(p));
            var s = t - (2.30753 + 0.27061 * t) / (1.0 + (0.99229 + 0.04481 * t) * t);

            var al = (s * s - 3.0) / 6.0;
            var h = 2.0 / (1.0 / (2.0 * a - 1.0) + 1.0 / (2.0 * b - 1.0));
            var w = s * Math.Sqrt(al + h) / h - (1.0 / (2.0 * b - 1.0) - 1.0 / (2.0 * a - 1.0))
                    * (al + 5.0 / 6.0 - 2.0 / (3.0 * h));
            return a / (a + b * Math.Exp(2.0 * w));
        }

        var lnBetaAB = LnBeta(a, b);
        var lnX = (Math.Log(p) + Math.Log(a) + lnBetaAB) / a;
        var lnOneMinusX = (Math.Log(1.0 - p) + Math.Log(b) + lnBetaAB) / b;

        var t2 = Math.Exp(lnX);
        return t2 <= 1.0 ? t2 : 1.0 - Math.Exp(lnOneMinusX);
    }

    private static double InverseBetaNewtonRefine(double x, double p, double a, double b)
    {
        var lo = 0.0;
        var hi = 1.0;
        var lnBeta = -LnBeta(a, b);

        for (var i = 0; i < 100; i++)
        {
            var err = BetaRegularized(x, a, b) - p;

            if (err < 0)
                lo = x;
            else
                hi = x;

            if (Math.Abs(err) < 1e-15)
                break;

            var logPdf = (a - 1.0) * Math.Log(x) + (b - 1.0) * Math.Log(1.0 - x) + lnBeta;
            var pdf = Math.Exp(logPdf);

            if (pdf > 0)
            {
                var newX = x - err / pdf;
                x = (newX > lo && newX < hi) ? newX : (lo + hi) / 2.0;
            }
            else
            {
                x = (lo + hi) / 2.0;
            }

            if (hi - lo < 1e-15 * x)
                break;
        }

        return x;
    }

    /// <summary>
    /// One-tailed (left) inverse of the Student's t-distribution.
    /// Returns value t such that P(T &lt;= t) = probability, where T ~ t(df).
    /// Uses Newton's method directly on the t-distribution CDF/PDF.
    /// </summary>
    internal static double StudentTInv(double probability, double degreesOfFreedom)
    {
        if (probability == 0.5)
            return 0.0;

        // Use symmetry for p < 0.5: T.INV(p, df) = -T.INV(1-p, df)
        if (probability < 0.5)
            return -StudentTInv(1.0 - probability, degreesOfFreedom);

        var v = degreesOfFreedom;

        // Initial guess using inverse normal with Cornish-Fisher correction
        var z = InverseNormalCdf(probability);
        var t = z + (z * z * z + z) / (4.0 * v)
                  + (5.0 * z * z * z * z * z + 16.0 * z * z * z + 3.0 * z) / (96.0 * v * v);

        // Precompute constants for the t-PDF: f(t) = (1 + t²/v)^(-(v+1)/2) / (sqrt(v) * B(v/2, 1/2))
        var lnPdfConst = -0.5 * Math.Log(v) - LnBeta(v / 2.0, 0.5);

        // Newton's method refinement
        for (var i = 0; i < 50; i++)
        {
            var cdf = StudentTCdf(t, v);
            var err = cdf - probability;

            if (Math.Abs(err) < 1e-15)
                break;

            // PDF of the t-distribution
            var logPdf = lnPdfConst - (v + 1.0) / 2.0 * Math.Log(1.0 + t * t / v);
            var pdf = Math.Exp(logPdf);

            if (pdf < 1e-300)
                break;

            var correction = err / pdf;
            t -= correction;

            if (Math.Abs(correction) < 1e-14 * Math.Abs(t))
                break;
        }

        return t;
    }

    /// <summary>
    /// CDF of the Student's t-distribution: P(T &lt;= t) where T ~ t(v).
    /// Uses the relationship with the regularized incomplete beta function.
    /// </summary>
    private static double StudentTCdf(double t, double v)
    {
        var x = v / (v + t * t);
        var iBeta = BetaRegularized(x, v / 2.0, 0.5);

        if (t >= 0)
            return 1.0 - 0.5 * iBeta;

        return 0.5 * iBeta;
    }

    /// <summary>
    /// Inverse of the standard normal CDF (probit function) for p in (0, 1).
    /// Uses the rational approximation from Abramowitz and Stegun (26.2.23) with refinement.
    /// </summary>
    private static double InverseNormalCdf(double p)
    {
        if (p < 0.5)
            return -InverseNormalCdf(1.0 - p);

        // Rational approximation for the inverse normal tail
        var t = Math.Sqrt(-2.0 * Math.Log(1.0 - p));

        // Coefficients for the rational approximation (Abramowitz & Stegun, improved)
        const double c0 = 2.515517;
        const double c1 = 0.802853;
        const double c2 = 0.010328;
        const double d1 = 1.432788;
        const double d2 = 0.189269;
        const double d3 = 0.001308;

        return t - (c0 + c1 * t + c2 * t * t) / (1.0 + d1 * t + d2 * t * t + d3 * t * t * t);
    }

    #endregion
}
