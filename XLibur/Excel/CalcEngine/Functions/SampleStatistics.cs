using System.Collections.Generic;

namespace XLibur.Excel.CalcEngine.Functions;

/// <summary>
/// The handful of operations every statistical function needs before it can do anything
/// interesting: reduce an argument to one number, and take the mean and spread of a materialized
/// sample. They live here rather than in each caller because <see cref="Distributions"/> and
/// <see cref="Regression"/> both want all of them.
/// </summary>
internal static class SampleStatistics
{
    /// <summary>
    /// Reduce an argument to the single number a scalar parameter wants. The statistical functions
    /// mark only their data parameters as taking a range, so the rest arrive unreduced: a reference
    /// to one cell is unwrapped, a larger one goes through implicit intersection.
    /// </summary>
    internal static bool TryGetScalarNumber(CalcContext ctx, in AnyValue value, out double number, out XLError error)
    {
        number = 0;
        error = default;

        if (!value.TryPickScalar(out var scalar, out var collection))
        {
            if (collection.TryPickT0(out var array, out var reference))
            {
                scalar = array[0, 0];
            }
            else if (!reference.TryGetSingleCellValue(out scalar, ctx)
                     && !value.ImplicitIntersection(ctx).TryPickScalar(out scalar, out _))
            {
                error = XLError.IncompatibleValue;
                return false;
            }
        }

        return scalar.ToNumber(ctx.Culture).TryPickT0(out number, out error);
    }

    internal static double Mean(List<double> values)
    {
        var total = 0d;
        foreach (var value in values)
            total += value;

        return total / values.Count;
    }

    /// <summary>Σ(x − mean)², the numerator every variance and moment in the library is built on.</summary>
    internal static double SumOfSquaredDeviations(List<double> values, double mean)
    {
        var total = 0d;
        foreach (var value in values)
            total += (value - mean) * (value - mean);

        return total;
    }

    /// <summary>The variance with Bessel's correction — the estimate of a population's from a sample of it.</summary>
    internal static double SampleVariance(List<double> values)
        => SumOfSquaredDeviations(values, Mean(values)) / (values.Count - 1);
}
