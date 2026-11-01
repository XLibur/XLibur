using System.Collections.Generic;

namespace XLibur.Excel.CalcEngine.Functions;

/// <summary>
/// The handful of operations every statistical function needs before it can do anything
/// interesting: take the mean and spread of a materialized sample. They live here rather than in
/// each caller because <see cref="Distributions"/> and <see cref="Regression"/> both want all of
/// them. A scalar parameter that arrives unreduced is read with
/// <see cref="AnyValue.TryReduceToNumber"/>.
/// </summary>
internal static class SampleStatistics
{
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
