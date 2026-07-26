using System;

namespace XLibur.Excel.CalcEngine.Functions;

/// <summary>
/// Bessel functions of integer order, backing BESSELJ, BESSELY, BESSELI and BESSELK.
/// <para>
/// Orders 0 and 1 come from the classic rational and asymptotic approximations (Abramowitz &amp;
/// Stegun 9.4 / 9.8, as tabulated in Numerical Recipes); higher orders are reached by recurrence.
/// The direction matters: J and I are recurred <em>downwards</em> from a high starting order and
/// normalised, because upward recurrence for those is numerically unstable, while Y and K are
/// dominant solutions and recur upwards safely. Accuracy is around 1e-8 relative, which is the
/// accuracy of the underlying order-0/1 approximations.
/// </para>
/// </summary>
internal static class Bessel
{
    private const double TwoOverPi = 0.636619772367581343;
    private const double Big = 1.0e10;
    private const double Small = 1.0e-10;

    /// <summary>Bessel function of the first kind, J_n(x).</summary>
    internal static double J(double x, int order)
    {
        // The rational approximation is only good to about 1e-8, and the origin is the one place
        // where the exact value matters and is trivially known.
        if (x == 0.0)
            return order == 0 ? 1.0 : 0.0;

        if (order == 0)
            return J0(x);
        if (order == 1)
            return J1(x);

        var ax = Math.Abs(x);
        double result;
        if (ax > order)
        {
            // Upward recurrence is stable once x exceeds the order.
            var twoOverX = 2.0 / ax;
            var previous = J0(ax);
            var current = J1(ax);
            for (var j = 1; j < order; j++)
            {
                var next = j * twoOverX * current - previous;
                previous = current;
                current = next;
            }

            result = current;
        }
        else
        {
            result = JByDownwardRecurrence(ax, order);
        }

        // J_n(-x) = (-1)^n J_n(x).
        return x < 0.0 && (order & 1) == 1 ? -result : result;
    }

    /// <summary>
    /// Miller's algorithm: start from a high order with an arbitrary seed, recur down, and rescale
    /// using the identity J_0 + 2·(J_2 + J_4 + …) = 1.
    /// </summary>
    private static double JByDownwardRecurrence(double ax, int order)
    {
        var twoOverX = 2.0 / ax;
        var start = 2 * ((order + (int)Math.Sqrt(40.0 * order)) / 2);

        double result = 0, sum = 0;
        double higher = 0, current = 1.0;
        var addToSum = false;

        for (var j = start; j > 0; j--)
        {
            var lower = j * twoOverX * current - higher;
            higher = current;
            current = lower;

            if (Math.Abs(current) > Big)
            {
                current *= Small;
                higher *= Small;
                result *= Small;
                sum *= Small;
            }

            if (addToSum)
                sum += current;
            addToSum = !addToSum;

            if (j == order)
                result = higher;
        }

        return result / (2.0 * sum - current);
    }

    /// <summary>Bessel function of the second kind, Y_n(x). Defined only for positive x.</summary>
    internal static double Y(double x, int order)
    {
        if (order == 0)
            return Y0(x);
        if (order == 1)
            return Y1(x);

        var twoOverX = 2.0 / x;
        var previous = Y0(x);
        var current = Y1(x);
        for (var j = 1; j < order; j++)
        {
            var next = j * twoOverX * current - previous;
            previous = current;
            current = next;
        }

        return current;
    }

    /// <summary>Modified Bessel function of the first kind, I_n(x).</summary>
    internal static double I(double x, int order)
    {
        if (x == 0.0)
            return order == 0 ? 1.0 : 0.0;

        if (order == 0)
            return I0(x);
        if (order == 1)
            return I1(x);

        var twoOverX = 2.0 / Math.Abs(x);
        double result = 0;
        double higher = 0, current = 1.0;

        for (var j = 2 * (order + (int)Math.Sqrt(40.0 * order)); j > 0; j--)
        {
            var lower = higher + j * twoOverX * current;
            higher = current;
            current = lower;

            if (Math.Abs(current) > Big)
            {
                result *= Small;
                current *= Small;
                higher *= Small;
            }

            if (j == order)
                result = higher;
        }

        result *= I0(x) / current;
        return x < 0.0 && (order & 1) == 1 ? -result : result;
    }

    /// <summary>Modified Bessel function of the second kind, K_n(x). Defined only for positive x.</summary>
    internal static double K(double x, int order)
    {
        if (order == 0)
            return K0(x);
        if (order == 1)
            return K1(x);

        var twoOverX = 2.0 / x;
        var previous = K0(x);
        var current = K1(x);
        for (var j = 1; j < order; j++)
        {
            var next = previous + j * twoOverX * current;
            previous = current;
            current = next;
        }

        return current;
    }

    private static double J0(double x)
    {
        var ax = Math.Abs(x);
        if (ax < 8.0)
        {
            var y = x * x;
            var numerator = 57568490574.0 + y * (-13362590354.0 + y * (651619640.7
                + y * (-11214424.18 + y * (77392.33017 + y * -184.9052456))));
            var denominator = 57568490411.0 + y * (1029532985.0 + y * (9494680.718
                + y * (59272.64853 + y * (267.8532712 + y))));
            return numerator / denominator;
        }

        var (cosPart, sinPart, z) = AsymptoticFirstKindOrderZero(ax);
        return Math.Sqrt(TwoOverPi / ax) * (Math.Cos(ax - 0.785398164) * cosPart - z * Math.Sin(ax - 0.785398164) * sinPart);
    }

    private static double J1(double x)
    {
        var ax = Math.Abs(x);
        double result;
        if (ax < 8.0)
        {
            var y = x * x;
            var numerator = x * (72362614232.0 + y * (-7895059235.0 + y * (242396853.1
                + y * (-2972611.439 + y * (15704.48260 + y * -30.16036606)))));
            var denominator = 144725228442.0 + y * (2300535178.0 + y * (18583304.74
                + y * (99447.43394 + y * (376.9991397 + y))));
            return numerator / denominator;
        }

        var (cosPart, sinPart, z) = AsymptoticFirstKindOrderOne(ax);
        result = Math.Sqrt(TwoOverPi / ax) * (Math.Cos(ax - 2.356194491) * cosPart - z * Math.Sin(ax - 2.356194491) * sinPart);
        return x < 0.0 ? -result : result;
    }

    private static double Y0(double x)
    {
        if (x < 8.0)
        {
            var y = x * x;
            var numerator = -2957821389.0 + y * (7062834065.0 + y * (-512359803.6
                + y * (10879881.29 + y * (-86327.92757 + y * 228.4622733))));
            var denominator = 40076544269.0 + y * (745249964.8 + y * (7189466.438
                + y * (47447.26470 + y * (226.1030244 + y))));
            return numerator / denominator + TwoOverPi * J0(x) * Math.Log(x);
        }

        var (cosPart, sinPart, z) = AsymptoticFirstKindOrderZero(x);
        return Math.Sqrt(TwoOverPi / x) * (Math.Sin(x - 0.785398164) * cosPart + z * Math.Cos(x - 0.785398164) * sinPart);
    }

    private static double Y1(double x)
    {
        if (x < 8.0)
        {
            var y = x * x;
            var numerator = x * (-0.4900604943e13 + y * (0.1275274390e13
                + y * (-0.5153438139e11 + y * (0.7349264551e9
                + y * (-0.4237922726e7 + y * 0.8511937935e4)))));
            var denominator = 0.2499580570e14 + y * (0.4244419664e12
                + y * (0.3733650367e10 + y * (0.2245904002e8
                + y * (0.1020426050e6 + y * (0.3549632885e3 + y)))));
            return numerator / denominator + TwoOverPi * (J1(x) * Math.Log(x) - 1.0 / x);
        }

        var (cosPart, sinPart, z) = AsymptoticFirstKindOrderOne(x);
        return Math.Sqrt(TwoOverPi / x) * (Math.Sin(x - 2.356194491) * cosPart + z * Math.Cos(x - 2.356194491) * sinPart);
    }

    /// <summary>The shared amplitude polynomials of the order-0 large-argument expansion.</summary>
    private static (double CosPart, double SinPart, double Z) AsymptoticFirstKindOrderZero(double ax)
    {
        var z = 8.0 / ax;
        var y = z * z;
        var cosPart = 1.0 + y * (-0.1098628627e-2 + y * (0.2734510407e-4
            + y * (-0.2073370639e-5 + y * 0.2093887211e-6)));
        var sinPart = -0.1562499995e-1 + y * (0.1430488765e-3
            + y * (-0.6911147651e-5 + y * (0.7621095161e-6 + y * -0.934935152e-7)));
        return (cosPart, sinPart, z);
    }

    /// <summary>The shared amplitude polynomials of the order-1 large-argument expansion.</summary>
    private static (double CosPart, double SinPart, double Z) AsymptoticFirstKindOrderOne(double ax)
    {
        var z = 8.0 / ax;
        var y = z * z;
        var cosPart = 1.0 + y * (0.183105e-2 + y * (-0.3516396496e-4
            + y * (0.2457520174e-5 + y * -0.240337019e-6)));
        var sinPart = 0.04687499995 + y * (-0.2002690873e-3
            + y * (0.8449199096e-5 + y * (-0.88228987e-6 + y * 0.105787412e-6)));
        return (cosPart, sinPart, z);
    }

    private static double I0(double x)
    {
        var ax = Math.Abs(x);
        if (ax < 3.75)
        {
            var y = x / 3.75;
            y *= y;
            return 1.0 + y * (3.5156229 + y * (3.0899424 + y * (1.2067492
                + y * (0.2659732 + y * (0.360768e-1 + y * 0.45813e-2)))));
        }

        var t = 3.75 / ax;
        return Math.Exp(ax) / Math.Sqrt(ax) * (0.39894228 + t * (0.1328592e-1
            + t * (0.225319e-2 + t * (-0.157565e-2 + t * (0.916281e-2
            + t * (-0.2057706e-1 + t * (0.2635537e-1 + t * (-0.1647633e-1
            + t * 0.392377e-2))))))));
    }

    private static double I1(double x)
    {
        var ax = Math.Abs(x);
        double result;
        if (ax < 3.75)
        {
            var y = x / 3.75;
            y *= y;
            result = ax * (0.5 + y * (0.87890594 + y * (0.51498869 + y * (0.15084934
                + y * (0.2658733e-1 + y * (0.301532e-2 + y * 0.32411e-3))))));
        }
        else
        {
            var t = 3.75 / ax;
            var tail = 0.2282967e-1 + t * (-0.2895312e-1 + t * (0.1787654e-1 - t * 0.420059e-2));
            tail = 0.39894228 + t * (-0.3988024e-1 + t * (-0.362018e-2
                + t * (0.163801e-2 + t * (-0.1031555e-1 + t * tail))));
            result = tail * (Math.Exp(ax) / Math.Sqrt(ax));
        }

        return x < 0.0 ? -result : result;
    }

    private static double K0(double x)
    {
        if (x <= 2.0)
        {
            var y = x * x / 4.0;
            return -Math.Log(x / 2.0) * I0(x) + (-0.57721566 + y * (0.42278420
                + y * (0.23069756 + y * (0.3488590e-1 + y * (0.262698e-2
                + y * (0.10750e-3 + y * 0.74e-5))))));
        }

        var t = 2.0 / x;
        return Math.Exp(-x) / Math.Sqrt(x) * (1.25331414 + t * (-0.7832358e-1
            + t * (0.2189568e-1 + t * (-0.1062446e-1 + t * (0.587872e-2
            + t * (-0.251540e-2 + t * 0.53208e-3))))));
    }

    private static double K1(double x)
    {
        if (x <= 2.0)
        {
            var y = x * x / 4.0;
            return Math.Log(x / 2.0) * I1(x) + 1.0 / x * (1.0 + y * (0.15443144
                + y * (-0.67278579 + y * (-0.18156897 + y * (-0.1919402e-1
                + y * (-0.110404e-2 + y * -0.4686e-4))))));
        }

        var t = 2.0 / x;
        return Math.Exp(-x) / Math.Sqrt(x) * (1.25331414 + t * (0.23498619
            + t * (-0.3655620e-1 + t * (0.1504268e-1 + t * (-0.780353e-2
            + t * (0.325614e-2 + t * -0.68245e-3))))));
    }
}
