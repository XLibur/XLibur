using System;
using System.Collections.Generic;
using XLibur.Excel.CalcEngine.Functions;
using static XLibur.Excel.CalcEngine.Functions.SignatureAdapter;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// Regression and the descriptive statistics that go with it.
/// <para>
/// The paired-sample functions (CORREL, SLOPE, RSQ, …) all read the same three sums — of the
/// squared deviations in x, in y, and of their product — so they share one pass over the data and
/// differ only in what they do with those three numbers.
/// </para>
/// <para>
/// LINEST, LOGEST, TREND and GROWTH are one least-squares fit behind four names: the exponential
/// pair is the linear pair applied to the logarithm of y. All four return arrays and spill.
/// </para>
/// </summary>
internal static class Regression
{
    public static void Register(FunctionRegistry ce)
    {
        ce.RegisterFunction("AVEDEV", 1, 255, AveDev, FunctionFlags.Range, AllowRange.All); // Average of the absolute deviations from the mean
        ce.RegisterFunction("CORREL", 2, 2, Correl, FunctionFlags.Range, AllowRange.All); // Correlation coefficient between two data sets
        ce.RegisterFunction("COVAR", 2, 2, CovarianceP, FunctionFlags.Range, AllowRange.All); // Population covariance
        ce.RegisterFunction("COVARIANCE.P", 2, 2, CovarianceP, FunctionFlags.Range | FunctionFlags.Future, AllowRange.All);
        ce.RegisterFunction("COVARIANCE.S", 2, 2, CovarianceS, FunctionFlags.Range | FunctionFlags.Future, AllowRange.All); // Sample covariance
        ce.RegisterFunction("FORECAST", 3, 3, Forecast, FunctionFlags.Range, AllowRange.Only, 1, 2); // A value along a linear trend
        ce.RegisterFunction("FORECAST.LINEAR", 3, 3, Forecast, FunctionFlags.Range | FunctionFlags.Future, AllowRange.Only, 1, 2);
        ce.RegisterFunction("FREQUENCY", 2, 2, Frequency, FunctionFlags.Range | FunctionFlags.ReturnsArray, AllowRange.All); // Distribution of values across bins
        ce.RegisterFunction("GROWTH", 1, 4, Growth, FunctionFlags.Range | FunctionFlags.ReturnsArray, AllowRange.All); // Values along an exponential trend
        ce.RegisterFunction("HARMEAN", 1, 255, HarMean, FunctionFlags.Range, AllowRange.All); // Harmonic mean
        ce.RegisterFunction("INTERCEPT", 2, 2, Intercept, FunctionFlags.Range, AllowRange.All); // Where the regression line meets the y axis
        ce.RegisterFunction("KURT", 1, 255, Kurt, FunctionFlags.Range, AllowRange.All); // Kurtosis of a data set
        ce.RegisterFunction("LINEST", 1, 4, Linest, FunctionFlags.Range | FunctionFlags.ReturnsArray, AllowRange.All); // Parameters of a linear trend
        ce.RegisterFunction("LOGEST", 1, 4, Logest, FunctionFlags.Range | FunctionFlags.ReturnsArray, AllowRange.All); // Parameters of an exponential trend
        ce.RegisterFunction("PEARSON", 2, 2, Correl, FunctionFlags.Range, AllowRange.All); // The same coefficient CORREL returns
        ce.RegisterFunction("PROB", 3, 4, Prob, FunctionFlags.Range, AllowRange.Only, 0, 1); // Probability that values fall between two limits
        ce.RegisterFunction("RSQ", 2, 2, RSq, FunctionFlags.Range, AllowRange.All); // Square of the correlation coefficient
        ce.RegisterFunction("SKEW", 1, 255, Skew, FunctionFlags.Range, AllowRange.All); // Skewness of a distribution
        ce.RegisterFunction("SKEW.P", 1, 255, SkewP, FunctionFlags.Range | FunctionFlags.Future, AllowRange.All); // Skewness of a whole population
        ce.RegisterFunction("SLOPE", 2, 2, Slope, FunctionFlags.Range, AllowRange.All); // Slope of the regression line
        ce.RegisterFunction("STEYX", 2, 2, SteyX, FunctionFlags.Range, AllowRange.All); // Standard error of the predicted y
        ce.RegisterFunction("TREND", 1, 4, Trend, FunctionFlags.Range | FunctionFlags.ReturnsArray, AllowRange.All); // Values along a linear trend
        ce.RegisterFunction("TRIMMEAN", 2, 2, Adapt(TrimMean), FunctionFlags.Range, AllowRange.Only, 0); // Mean of the interior of a data set
    }

    #region Paired-sample statistics

    /// <summary>
    /// The sums every paired-sample function is built from: the count, both means, the two sums of
    /// squared deviations and the sum of their products. One pass, six numbers, and each function
    /// is a line of arithmetic over them.
    /// </summary>
    private readonly record struct PairedSums(int Count, double MeanX, double MeanY, double SumXX, double SumYY, double SumXY);

    private static AnyValue Correl(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryGetPairedSums(ctx, args[0], args[1], out var sums, out var error))
            return error;

        if (sums.Count < 2 || sums.SumXX == 0 || sums.SumYY == 0)
            return XLError.DivisionByZero;

        return sums.SumXY / Math.Sqrt(sums.SumXX * sums.SumYY);
    }

    private static AnyValue RSq(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryGetPairedSums(ctx, args[0], args[1], out var sums, out var error))
            return error;

        if (sums.Count < 2 || sums.SumXX == 0 || sums.SumYY == 0)
            return XLError.DivisionByZero;

        var correlation = sums.SumXY / Math.Sqrt(sums.SumXX * sums.SumYY);
        return correlation * correlation;
    }

    private static AnyValue CovarianceP(CalcContext ctx, Span<AnyValue> args)
        => Covariance(ctx, args, population: true);

    private static AnyValue CovarianceS(CalcContext ctx, Span<AnyValue> args)
        => Covariance(ctx, args, population: false);

    private static AnyValue Covariance(CalcContext ctx, Span<AnyValue> args, bool population)
    {
        if (!TryGetPairedSums(ctx, args[0], args[1], out var sums, out var error))
            return error;

        if (sums.Count < 1 || (!population && sums.Count < 2))
            return XLError.DivisionByZero;

        return sums.SumXY / (population ? sums.Count : sums.Count - 1);
    }

    /// <summary>SLOPE(known_y, known_x) — note that y comes first, as it does in every Excel regression function.</summary>
    private static AnyValue Slope(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryGetPairedSums(ctx, args[1], args[0], out var sums, out var error))
            return error;

        if (sums.Count < 2 || sums.SumXX == 0)
            return XLError.DivisionByZero;

        return sums.SumXY / sums.SumXX;
    }

    private static AnyValue Intercept(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryGetPairedSums(ctx, args[1], args[0], out var sums, out var error))
            return error;

        if (sums.Count < 2 || sums.SumXX == 0)
            return XLError.DivisionByZero;

        return sums.MeanY - sums.SumXY / sums.SumXX * sums.MeanX;
    }

    /// <summary>
    /// STEYX(known_y, known_x) — the standard error of the y the regression predicts: the spread of
    /// y that the line does not account for, over the n−2 degrees of freedom left after fitting it.
    /// </summary>
    private static AnyValue SteyX(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryGetPairedSums(ctx, args[1], args[0], out var sums, out var error))
            return error;

        if (sums.Count < 3 || sums.SumXX == 0)
            return XLError.DivisionByZero;

        var residual = sums.SumYY - sums.SumXY * sums.SumXY / sums.SumXX;
        return Math.Sqrt(Math.Max(residual, 0) / (sums.Count - 2));
    }

    private static AnyValue Forecast(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryGetScalarNumber(ctx, args[0], out var x, out var xError))
            return xError;
        if (!TryGetPairedSums(ctx, args[2], args[1], out var sums, out var error))
            return error;

        if (sums.Count < 1 || sums.SumXX == 0)
            return XLError.DivisionByZero;

        var slope = sums.SumXY / sums.SumXX;
        return sums.MeanY - slope * sums.MeanX + slope * x;
    }

    /// <summary>
    /// Read two equally sized ranges as paired observations. A pair is used only when both of its
    /// values are numbers, which is how Excel skips a row with a gap in it rather than pairing
    /// values that do not belong together.
    /// </summary>
    private static bool TryGetPairedSums(CalcContext ctx, in AnyValue xArg, in AnyValue yArg, out PairedSums sums, out XLError error)
    {
        sums = default;
        error = default;

        if (!TryGetPairs(ctx, xArg, yArg, out var xs, out var ys, out error))
            return false;

        var count = xs.Count;
        if (count == 0)
        {
            error = XLError.DivisionByZero;
            return false;
        }

        double sumX = 0, sumY = 0;
        for (var i = 0; i < count; i++)
        {
            sumX += xs[i];
            sumY += ys[i];
        }

        var meanX = sumX / count;
        var meanY = sumY / count;

        double sumXX = 0, sumYY = 0, sumXY = 0;
        for (var i = 0; i < count; i++)
        {
            var dx = xs[i] - meanX;
            var dy = ys[i] - meanY;
            sumXX += dx * dx;
            sumYY += dy * dy;
            sumXY += dx * dy;
        }

        sums = new PairedSums(count, meanX, meanY, sumXX, sumYY, sumXY);
        return true;
    }

    private static bool TryGetPairs(CalcContext ctx, in AnyValue xArg, in AnyValue yArg, out List<double> xs, out List<double> ys, out XLError error)
    {
        xs = [];
        ys = [];
        error = default;

        if (!xArg.TryPickCollectionArray(out var xArray, ctx) || !yArg.TryPickCollectionArray(out var yArray, ctx))
        {
            error = XLError.IncompatibleValue;
            return false;
        }

        var xValues = Flatten(xArray!);
        var yValues = Flatten(yArray!);
        if (xValues.Count != yValues.Count)
        {
            error = XLError.NoValueAvailable;
            return false;
        }

        for (var i = 0; i < xValues.Count; i++)
        {
            if (xValues[i].TryPickError(out error) || yValues[i].TryPickError(out error))
                return false;

            if (xValues[i].TryPickNumber(out var x) && yValues[i].TryPickNumber(out var y))
            {
                xs.Add(x);
                ys.Add(y);
            }
        }

        return true;
    }

    #endregion

    #region Shape statistics

    private static AnyValue AveDev(CalcContext ctx, Span<AnyValue> args)
    {
        if (!Statistical.CollectNumbers(ctx, args, TallyNumbers.Default).TryPickT0(out var numbers, out var error))
            return error;
        if (numbers.Count == 0)
            return XLError.NumberInvalid;

        var mean = Mean(numbers);
        var total = 0d;
        foreach (var value in numbers)
            total += Math.Abs(value - mean);

        return total / numbers.Count;
    }

    private static AnyValue HarMean(CalcContext ctx, Span<AnyValue> args)
    {
        if (!Statistical.CollectNumbers(ctx, args, TallyNumbers.Default).TryPickT0(out var numbers, out var error))
            return error;
        if (numbers.Count == 0)
            return XLError.NumberInvalid;

        var reciprocals = 0d;
        foreach (var value in numbers)
        {
            // The harmonic mean is the reciprocal of a mean of reciprocals, so a zero or negative
            // value has no meaning in it.
            if (value <= 0)
                return XLError.NumberInvalid;

            reciprocals += 1 / value;
        }

        return numbers.Count / reciprocals;
    }

    /// <summary>
    /// SKEW — the sample skewness, which carries the n/((n−1)(n−2)) correction that makes it an
    /// unbiased estimate of the population's.
    /// </summary>
    private static AnyValue Skew(CalcContext ctx, Span<AnyValue> args)
    {
        if (!Statistical.CollectNumbers(ctx, args, TallyNumbers.Default).TryPickT0(out var numbers, out var error))
            return error;

        var n = numbers.Count;
        if (n < 3)
            return XLError.DivisionByZero;

        var mean = Mean(numbers);
        var standardDeviation = Math.Sqrt(SumOfSquaredDeviations(numbers, mean) / (n - 1));
        if (standardDeviation == 0)
            return XLError.DivisionByZero;

        var total = 0d;
        foreach (var value in numbers)
            total += Math.Pow((value - mean) / standardDeviation, 3);

        return (double)n / ((n - 1) * (n - 2)) * total;
    }

    /// <summary>SKEW.P — the population skewness, with no small-sample correction.</summary>
    private static AnyValue SkewP(CalcContext ctx, Span<AnyValue> args)
    {
        if (!Statistical.CollectNumbers(ctx, args, TallyNumbers.Default).TryPickT0(out var numbers, out var error))
            return error;

        var n = numbers.Count;
        if (n < 1)
            return XLError.DivisionByZero;

        var mean = Mean(numbers);
        var standardDeviation = Math.Sqrt(SumOfSquaredDeviations(numbers, mean) / n);
        if (standardDeviation == 0)
            return XLError.DivisionByZero;

        var total = 0d;
        foreach (var value in numbers)
            total += Math.Pow((value - mean) / standardDeviation, 3);

        return total / n;
    }

    /// <summary>
    /// KURT — the excess kurtosis, so a normal distribution scores zero rather than three. The
    /// second term is what subtracts that three, adjusted for the sample size.
    /// </summary>
    private static AnyValue Kurt(CalcContext ctx, Span<AnyValue> args)
    {
        if (!Statistical.CollectNumbers(ctx, args, TallyNumbers.Default).TryPickT0(out var numbers, out var error))
            return error;

        var n = numbers.Count;
        if (n < 4)
            return XLError.DivisionByZero;

        var mean = Mean(numbers);
        var standardDeviation = Math.Sqrt(SumOfSquaredDeviations(numbers, mean) / (n - 1));
        if (standardDeviation == 0)
            return XLError.DivisionByZero;

        var total = 0d;
        foreach (var value in numbers)
            total += Math.Pow((value - mean) / standardDeviation, 4);

        var scale = (double)n * (n + 1) / ((n - 1.0) * (n - 2) * (n - 3));
        var correction = 3.0 * (n - 1) * (n - 1) / ((n - 2.0) * (n - 3));
        return scale * total - correction;
    }

    /// <summary>
    /// TRIMMEAN(array, percent) — the mean after discarding the extremes. The count discarded is
    /// rounded down to an even number so that the same many come off each end.
    /// </summary>
    private static AnyValue TrimMean(CalcContext ctx, AnyValue arrayParam, double percent)
    {
        if (percent < 0 || percent >= 1)
            return XLError.NumberInvalid;

        if (!Statistical.TryGetNumbers(ctx, arrayParam, out var numbers, out var error))
            return error;
        if (numbers.Count == 0)
            return XLError.NumberInvalid;

        numbers.Sort();

        var discarded = (int)Math.Floor(numbers.Count * percent / 2) * 2;
        var kept = numbers.Count - discarded;
        var from = discarded / 2;

        var total = 0d;
        for (var i = from; i < from + kept; i++)
            total += numbers[i];

        return total / kept;
    }

    /// <summary>
    /// PROB(x_range, prob_range, lower, [upper]) — the total probability of the outcomes that fall
    /// within the limits. With no upper limit only the outcomes equal to <c>lower</c> count.
    /// </summary>
    private static AnyValue Prob(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryGetPairs(ctx, args[0], args[1], out var values, out var probabilities, out var error))
            return error;
        if (!TryGetScalarNumber(ctx, args[2], out var lowerLimit, out var lowerError))
            return lowerError;

        var upperLimit = lowerLimit;
        if (args.Length > 3 && !TryGetScalarNumber(ctx, args[3], out upperLimit, out var upperError))
            return upperError;

        var total = 0d;
        var matched = 0d;
        for (var i = 0; i < values.Count; i++)
        {
            var probability = probabilities[i];
            if (probability <= 0 || probability > 1)
                return XLError.NumberInvalid;

            total += probability;
            if (values[i] >= lowerLimit && values[i] <= upperLimit)
                matched += probability;
        }

        // The probabilities have to describe a whole distribution, up to the rounding that summing
        // them introduces.
        if (Math.Abs(total - 1) > 1e-7)
            return XLError.NumberInvalid;

        return matched;
    }

    #endregion

    #region Frequency

    /// <summary>
    /// FREQUENCY(data, bins) — how many values fall into each bin, as a column one longer than the
    /// bins: one count per bin, plus everything above the last. Bins are taken in the order given,
    /// each one closing at its own value.
    /// </summary>
    private static AnyValue Frequency(CalcContext ctx, Span<AnyValue> args)
    {
        if (!args[0].TryPickCollectionArray(out var dataArray, ctx))
            return XLError.IncompatibleValue;

        var data = new List<double>();
        foreach (var value in dataArray!)
        {
            if (value.TryPickError(out var dataError))
                return dataError;
            if (value.TryPickNumber(out var number))
                data.Add(number);
        }

        var bins = new List<double>();
        if (args[1].TryPickCollectionArray(out var binArray, ctx))
        {
            foreach (var value in binArray!)
            {
                if (value.TryPickError(out var binError))
                    return binError;
                if (value.TryPickNumber(out var number))
                    bins.Add(number);
            }
        }
        else if (args[1].TryPickScalar(out var scalar, out _))
        {
            if (scalar.TryPickError(out var scalarError))
                return scalarError;
            if (scalar.TryPickNumber(out var number))
                bins.Add(number);
        }

        // With no bins at all every value lands in the single overflow bucket.
        var counts = new ScalarValue[bins.Count + 1, 1];
        for (var i = 0; i <= bins.Count; i++)
        {
            var count = 0;
            foreach (var value in data)
            {
                var aboveLower = i == 0 || value > bins[i - 1];
                var atOrBelowUpper = i == bins.Count || value <= bins[i];
                if (aboveLower && atOrBelowUpper)
                    count++;
            }

            counts[i, 0] = count;
        }

        return new ConstArray(counts);
    }

    #endregion

    #region Least squares

    private static AnyValue Linest(CalcContext ctx, Span<AnyValue> args)
        => LinearFit(ctx, args, exponential: false);

    private static AnyValue Logest(CalcContext ctx, Span<AnyValue> args)
        => LinearFit(ctx, args, exponential: true);

    private static AnyValue Trend(CalcContext ctx, Span<AnyValue> args)
        => Predict(ctx, args, exponential: false);

    private static AnyValue Growth(CalcContext ctx, Span<AnyValue> args)
        => Predict(ctx, args, exponential: true);

    /// <summary>
    /// The observations of a fit, laid out as a design matrix. Excel lets the caller orient the data
    /// either way — y down a column with each predictor in its own column, or y across a row with
    /// each predictor in its own row — so both are read into the same shape here.
    /// </summary>
    private readonly record struct Design(double[,] X, double[] Y, int Observations, int Predictors, bool Vertical);

    private static AnyValue LinearFit(CalcContext ctx, Span<AnyValue> args, bool exponential)
    {
        // LINEST(known_y, [known_x], [const], [stats]).
        if (!TryReadDesign(ctx, args, exponential, constIndex: 2, out var design, out var constant, out var error))
            return error;

        var wantsStatistics = false;
        if (args.Length > 3 && !TryGetBoolean(ctx, args[3], out wantsStatistics, out var statsError))
            return statsError;

        if (!TrySolve(design, constant, out var coefficients))
            return XLError.NumberInvalid;

        // Excel reports the coefficients from the last predictor back to the first, then the
        // intercept, which is the reverse of how the fit produces them.
        var width = design.Predictors + 1;
        var row = new double[width];
        for (var i = 0; i < design.Predictors; i++)
            row[i] = coefficients[design.Predictors - i];
        row[design.Predictors] = coefficients[0];

        if (exponential)
        {
            // LOGEST fits ln(y), so the coefficients come back as logarithms of the factors.
            for (var i = 0; i < width; i++)
                row[i] = Math.Exp(row[i]);
        }

        if (!wantsStatistics)
        {
            var simple = new ScalarValue[1, width];
            for (var i = 0; i < width; i++)
                simple[0, i] = row[i];

            return new ConstArray(simple);
        }

        return BuildStatistics(design, constant, coefficients, row);
    }

    private static AnyValue Predict(CalcContext ctx, Span<AnyValue> args, bool exponential)
    {
        // TREND(known_y, [known_x], [new_x], [const]) — the flag sits one place later than in LINEST.
        if (!TryReadDesign(ctx, args, exponential, constIndex: 3, out var design, out var constant, out var error))
            return error;

        if (!TrySolve(design, constant, out var coefficients))
            return XLError.NumberInvalid;

        // The points to predict at default to the ones the fit was made from.
        double[,] newX;
        int newCount;
        if (args.Length > 2 && !IsOmitted(args, 2))
        {
            if (!args[2].TryPickCollectionArray(out var newArray, ctx))
                return XLError.IncompatibleValue;

            if (!TryReadPredictors(ctx, newArray!, design.Predictors, design.Vertical, out newX, out newCount, out var newError))
                return newError;
        }
        else
        {
            newX = design.X;
            newCount = design.Observations;
        }

        var predictions = new double[newCount];
        for (var i = 0; i < newCount; i++)
        {
            var value = coefficients[0];
            for (var p = 1; p <= design.Predictors; p++)
                value += coefficients[p] * newX[i, p - 1];

            predictions[i] = exponential ? Math.Exp(value) : value;
        }

        var data = design.Vertical ? new ScalarValue[newCount, 1] : new ScalarValue[1, newCount];
        for (var i = 0; i < newCount; i++)
        {
            if (design.Vertical)
                data[i, 0] = predictions[i];
            else
                data[0, i] = predictions[i];
        }

        return new ConstArray(data);
    }

    /// <summary>
    /// Read known_y and known_x into a design matrix. When known_x is left out the predictor is the
    /// sequence 1, 2, 3, … , which is what makes TREND(known_y) a fit against position.
    /// </summary>
    private static bool TryReadDesign(CalcContext ctx, Span<AnyValue> args, bool exponential, int constIndex, out Design design, out bool constant, out XLError error)
    {
        design = default;
        constant = true;
        error = default;

        if (constIndex < args.Length && !IsOmitted(args, constIndex)
            && !TryGetBoolean(ctx, args[constIndex], out constant, out error))
        {
            return false;
        }

        if (!args[0].TryPickCollectionArray(out var yArray, ctx))
        {
            error = XLError.IncompatibleValue;
            return false;
        }

        // A column of y means each predictor is a column too; a row of y means each is a row.
        var vertical = yArray!.Width == 1;
        var observations = vertical ? yArray.Height : yArray.Width;

        var y = new double[observations];
        for (var i = 0; i < observations; i++)
        {
            var scalar = vertical ? yArray[i, 0] : yArray[0, i];
            if (scalar.TryPickError(out error))
                return false;

            if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var value, out error))
                return false;

            if (exponential)
            {
                // Fitting y = b·m^x means fitting ln(y) linearly, which needs a positive y.
                if (value <= 0)
                {
                    error = XLError.NumberInvalid;
                    return false;
                }

                value = Math.Log(value);
            }

            y[i] = value;
        }

        double[,] x;
        int predictors;
        if (args.Length > 1 && !IsOmitted(args, 1))
        {
            if (!args[1].TryPickCollectionArray(out var xArray, ctx))
            {
                error = XLError.IncompatibleValue;
                return false;
            }

            if (!TryReadPredictors(ctx, xArray!, 0, vertical, out x, out var xObservations, out error))
                return false;

            if (xObservations != observations)
            {
                error = XLError.CellReference;
                return false;
            }

            predictors = x.GetLength(1);
        }
        else
        {
            predictors = 1;
            x = new double[observations, 1];
            for (var i = 0; i < observations; i++)
                x[i, 0] = i + 1;
        }

        design = new Design(x, y, observations, predictors, vertical);
        return true;
    }

    /// <summary>
    /// Read a block of predictor values. <paramref name="expected"/> is the number of predictors to
    /// insist on, or zero to take whatever the block holds.
    /// </summary>
    private static bool TryReadPredictors(CalcContext ctx, Array array, int expected, bool vertical, out double[,] x, out int observations, out XLError error)
    {
        error = default;
        observations = vertical ? array.Height : array.Width;
        var predictors = vertical ? array.Width : array.Height;

        // A single-vector block that runs the other way is still one predictor read along its length.
        if (expected == 1 && predictors != 1 && observations == 1)
        {
            (observations, predictors) = (predictors, 1);
            vertical = !vertical;
        }

        if (expected > 0 && predictors != expected)
        {
            x = new double[0, 0];
            error = XLError.CellReference;
            return false;
        }

        x = new double[observations, predictors];
        for (var i = 0; i < observations; i++)
        {
            for (var p = 0; p < predictors; p++)
            {
                var scalar = vertical ? array[i, p] : array[p, i];
                if (scalar.TryPickError(out error))
                    return false;

                if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var value, out error))
                    return false;

                x[i, p] = value;
            }
        }

        return true;
    }

    /// <summary>
    /// Solve the normal equations XᵀX·β = Xᵀy. <paramref name="constant"/> false pins the intercept
    /// at zero by dropping its column from the fit and reporting it as zero.
    /// </summary>
    private static bool TrySolve(in Design design, bool constant, out double[] coefficients)
    {
        var terms = design.Predictors + 1;
        coefficients = new double[terms];

        var columns = constant ? terms : design.Predictors;
        var normal = new XLMatrix(columns, columns);
        var rhs = new XLMatrix(columns, 1);

        for (var a = 0; a < columns; a++)
        {
            for (var b = 0; b < columns; b++)
                normal[a, b] = DotProduct(design, constant, a, b);

            rhs[a, 0] = DotProductWithY(design, constant, a);
        }

        try
        {
            var solution = normal.SolveWith(rhs);
            if (constant)
            {
                for (var i = 0; i < terms; i++)
                    coefficients[i] = solution[i, 0];
            }
            else
            {
                for (var i = 0; i < design.Predictors; i++)
                    coefficients[i + 1] = solution[i, 0];
            }
        }
        catch (InvalidOperationException)
        {
            return false;
        }

        return true;
    }

    /// <summary>Column <paramref name="index"/> of the design matrix, where column zero is the intercept's constant one.</summary>
    private static double Column(in Design design, bool constant, int index, int observation)
    {
        var predictor = constant ? index - 1 : index;
        return predictor < 0 ? 1 : design.X[observation, predictor];
    }

    private static double DotProduct(in Design design, bool constant, int a, int b)
    {
        var total = 0d;
        for (var i = 0; i < design.Observations; i++)
            total += Column(design, constant, a, i) * Column(design, constant, b, i);

        return total;
    }

    private static double DotProductWithY(in Design design, bool constant, int a)
    {
        var total = 0d;
        for (var i = 0; i < design.Observations; i++)
            total += Column(design, constant, a, i) * design.Y[i];

        return total;
    }

    /// <summary>
    /// The five-row statistics block LINEST and LOGEST return when asked for it: the coefficients
    /// and their standard errors, then r² and the standard error of y, then the F statistic and its
    /// degrees of freedom, then the regression and residual sums of squares. The cells to the right
    /// of the short rows are <c>#N/A</c>, as Excel leaves them.
    /// </summary>
    private static AnyValue BuildStatistics(in Design design, bool constant, double[] coefficients, double[] reportedRow)
    {
        var width = design.Predictors + 1;
        var n = design.Observations;
        var degreesOfFreedom = n - design.Predictors - (constant ? 1 : 0);

        var meanY = 0d;
        foreach (var value in design.Y)
            meanY += value;
        meanY /= n;

        double residualSumOfSquares = 0, totalSumOfSquares = 0;
        for (var i = 0; i < n; i++)
        {
            var fitted = constant ? coefficients[0] : 0;
            for (var p = 1; p <= design.Predictors; p++)
                fitted += coefficients[p] * design.X[i, p - 1];

            var residual = design.Y[i] - fitted;
            residualSumOfSquares += residual * residual;
            totalSumOfSquares += constant ? (design.Y[i] - meanY) * (design.Y[i] - meanY) : design.Y[i] * design.Y[i];
        }

        var regressionSumOfSquares = totalSumOfSquares - residualSumOfSquares;
        var standardErrorOfY = degreesOfFreedom > 0 ? Math.Sqrt(residualSumOfSquares / degreesOfFreedom) : 0;
        var rSquared = totalSumOfSquares > 0 ? regressionSumOfSquares / totalSumOfSquares : 0;
        var fStatistic = design.Predictors > 0 && residualSumOfSquares > 0
            ? regressionSumOfSquares / design.Predictors / (residualSumOfSquares / degreesOfFreedom)
            : 0;

        var standardErrors = StandardErrors(design, constant, standardErrorOfY);

        var data = new ScalarValue[5, width];
        for (var column = 0; column < width; column++)
        {
            data[0, column] = reportedRow[column];
            data[1, column] = standardErrors[column];
            data[2, column] = column switch { 0 => rSquared, 1 => standardErrorOfY, _ => XLError.NoValueAvailable };
            data[3, column] = column switch { 0 => fStatistic, 1 => degreesOfFreedom, _ => XLError.NoValueAvailable };
            data[4, column] = column switch { 0 => regressionSumOfSquares, 1 => residualSumOfSquares, _ => XLError.NoValueAvailable };
        }

        return new ConstArray(data);
    }

    /// <summary>
    /// Standard errors of the coefficients, in the same reversed order as the coefficients: the
    /// square roots of the diagonal of (XᵀX)⁻¹ scaled by the standard error of y. When the
    /// intercept is pinned at zero its error is reported as <c>#N/A</c>, as Excel does.
    /// </summary>
    private static ScalarValue[] StandardErrors(in Design design, bool constant, double standardErrorOfY)
    {
        var width = design.Predictors + 1;
        var errors = new ScalarValue[width];
        for (var i = 0; i < width; i++)
            errors[i] = XLError.NoValueAvailable;

        var columns = constant ? width : design.Predictors;
        var normal = new XLMatrix(columns, columns);
        for (var a = 0; a < columns; a++)
        {
            for (var b = 0; b < columns; b++)
                normal[a, b] = DotProduct(design, constant, a, b);
        }

        double[,] inverse;
        try
        {
            var inverted = normal.Invert();
            inverse = new double[columns, columns];
            for (var a = 0; a < columns; a++)
            {
                for (var b = 0; b < columns; b++)
                    inverse[a, b] = inverted[a, b];
            }
        }
        catch (InvalidOperationException)
        {
            return errors;
        }

        for (var p = 0; p < design.Predictors; p++)
        {
            var index = constant ? p + 1 : p;
            errors[design.Predictors - 1 - p] = standardErrorOfY * Math.Sqrt(Math.Abs(inverse[index, index]));
        }

        if (constant)
            errors[design.Predictors] = standardErrorOfY * Math.Sqrt(Math.Abs(inverse[0, 0]));

        return errors;
    }

    #endregion

    #region Shared helpers

    private static List<ScalarValue> Flatten(Array array)
    {
        var values = new List<ScalarValue>(array.Height * array.Width);
        foreach (var value in array)
            values.Add(value);

        return values;
    }

    private static bool IsOmitted(Span<AnyValue> args, int index)
        => args.Length <= index || (args[index].TryPickScalar(out var scalar, out _) && scalar.IsBlank);

    private static bool TryGetBoolean(CalcContext ctx, in AnyValue value, out bool flag, out XLError error)
    {
        flag = false;
        error = default;
        if (!value.TryPickScalar(out var scalar, out _))
        {
            if (!value.ImplicitIntersection(ctx).TryPickScalar(out scalar, out _))
            {
                error = XLError.IncompatibleValue;
                return false;
            }
        }

        if (scalar.IsBlank)
            return true;

        return scalar.TryCoerceLogicalOrBlankOrNumberOrText(out flag, out error);
    }

    private static bool TryGetScalarNumber(CalcContext ctx, in AnyValue value, out double number, out XLError error)
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

    private static double Mean(List<double> values)
    {
        var total = 0d;
        foreach (var value in values)
            total += value;

        return total / values.Count;
    }

    private static double SumOfSquaredDeviations(List<double> values, double mean)
    {
        var total = 0d;
        foreach (var value in values)
            total += (value - mean) * (value - mean);

        return total;
    }

    #endregion
}
