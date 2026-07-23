using System;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel.CalcEngine.Functions;
using static XLibur.Excel.CalcEngine.Functions.SignatureAdapter;

namespace XLibur.Excel.CalcEngine;

internal static class Statistical
{
    // Argument positions that may be ranges in *IFS functions that take a leading value range:
    // the value range at position 0, then the criteria ranges at odd positions (criteria values,
    // at even positions, are scalars). Mirrors MathTrig.SumIfsAllowedRangeParams.
    private static readonly int[] ValueAndCriteriaRangeParams =
        new[] { 0 }.Concat(Enumerable.Range(0, 128).Select(x => x * 2 + 1)).ToArray();

    public static void Register(FunctionRegistry ce)
    {
        //ce.RegisterFunction("AVEDEV", AveDev, 1, int.MaxValue);
        ce.RegisterFunction("AVERAGE", 1, int.MaxValue, Average, FunctionFlags.Range, AllowRange.All); // Returns the average (arithmetic mean) of the arguments
        ce.RegisterFunction("AVERAGEA", 1, int.MaxValue, AverageA, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("AVERAGEIF", 2, 3, AdaptLastOptional(AverageIf), FunctionFlags.Range, AllowRange.Only, 0, 2); // Returns the average of cells that meet a criterion
        ce.RegisterFunction("AVERAGEIFS", 3, 255, AdaptIfs(AverageIfs), FunctionFlags.Range, AllowRange.Only, ValueAndCriteriaRangeParams); // Returns the average of cells that meet multiple criteria
        //BETADIST	Returns the beta cumulative distribution function
        //BETAINV   Returns the inverse of the cumulative distribution function for a specified beta distribution
        ce.RegisterFunction("BINOMDIST", 4, 4, Adapt(BinomDist), FunctionFlags.Scalar); //BINOMDIST	Returns the individual term binomial distribution probability
        ce.RegisterFunction("BINOM.DIST", 4, 4, Adapt(BinomDist), FunctionFlags.Scalar); // In theory more precise BINOMDIST.
        //CHIDIST	Returns the one-tailed probability of the chi-squared distribution
        //CHIINV	Returns the inverse of the one-tailed probability of the chi-squared distribution
        //CHITEST	Returns the test for independence
        //CONFIDENCE	Returns the confidence interval for a population mean
        //CORREL	Returns the correlation coefficient between two data sets
        ce.RegisterFunction("COUNT", 1, int.MaxValue, Count, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("COUNTA", 1, 255, CountA, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("COUNTBLANK", 1, 1, Adapt(CountBlank), FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("COUNTIF", 2, 2, Adapt((Func<CalcContext, AnyValue, ScalarValue, AnyValue>)CountIf), FunctionFlags.Range, AllowRange.Only, 0);
        ce.RegisterFunction("COUNTIFS", 2, 255, AdaptIfs(CountIfs), FunctionFlags.Range, AllowRange.Only, Enumerable.Range(0, 128).Select(x => x * 2).ToArray());
        //COVAR	Returns covariance, the average of the products of paired deviations
        //CRITBINOM	Returns the smallest value for which the cumulative binomial distribution is less than or equal to a criterion value
        ce.RegisterFunction("DEVSQ", 1, 255, DevSq, FunctionFlags.Range, AllowRange.All); // Returns the sum of squares of deviations
        //EXPONDIST	Returns the exponential distribution
        //FDIST	Returns the F probability distribution
        //FINV	Returns the inverse of the F probability distribution
        ce.RegisterFunction("FISHER", 1, 1, Adapt(Fisher), FunctionFlags.Scalar); // Returns the Fisher transformation
        //FISHERINV	Returns the inverse of the Fisher transformation
        //FORECAST	Returns a value along a linear trend
        //FREQUENCY	Returns a frequency distribution as a vertical array
        //FTEST	Returns the result of an F-test
        //GAMMADIST	Returns the gamma distribution
        //GAMMAINV	Returns the inverse of the gamma cumulative distribution
        //GAMMALN	Returns the natural logarithm of the gamma function, Γ(x)
        ce.RegisterFunction("GEOMEAN", 1, 255, GeoMean, FunctionFlags.Range, AllowRange.All); // Returns the geometric mean
        //GROWTH	Returns values along an exponential trend
        //HARMEAN	Returns the harmonic mean
        //HYPGEOMDIST	Returns the hypergeometric distribution
        //INTERCEPT	Returns the intercept of the linear regression line
        //KURT	Returns the kurtosis of a data set
        //LARGE	Returns the k-th largest value in a data set
        ce.RegisterFunction("LARGE", 2, 2, Adapt(Large), FunctionFlags.Range, AllowRange.Only, 0);
        //LINEST	Returns the parameters of a linear trend
        //LOGEST	Returns the parameters of an exponential trend
        //LOGINV	Returns the inverse of the lognormal distribution
        //LOGNORMDIST	Returns the cumulative lognormal distribution
        ce.RegisterFunction("MAX", 1, 255, Max, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("MAXA", 1, int.MaxValue, MaxA, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("MAXIFS", 3, 255, AdaptIfs(MaxIfs), FunctionFlags.Range, AllowRange.Only, ValueAndCriteriaRangeParams); // Returns the maximum of cells that meet multiple criteria
        ce.RegisterFunction("MEDIAN", 1, int.MaxValue, Median, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("MIN", 1, int.MaxValue, Min, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("MINA", 1, int.MaxValue, MinA, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("MINIFS", 3, 255, AdaptIfs(MinIfs), FunctionFlags.Range, AllowRange.Only, ValueAndCriteriaRangeParams); // Returns the minimum of cells that meet multiple criteria
        ce.RegisterFunction("MODE", 1, 255, Mode, FunctionFlags.Range, AllowRange.All); // Returns the most common value in a data set
        ce.RegisterFunction("MODE.SNGL", 1, 255, Mode, FunctionFlags.Range, AllowRange.All);
        //NEGBINOMDIST	Returns the negative binomial distribution
        //NORMDIST	Returns the normal cumulative distribution
        //NORMINV	Returns the inverse of the normal cumulative distribution
        //NORMSDIST	Returns the standard normal cumulative distribution
        //NORMSINV	Returns the inverse of the standard normal cumulative distribution
        //PEARSON	Returns the Pearson product moment correlation coefficient
        ce.RegisterFunction("PERCENTILE", 2, 2, Adapt(Percentile), FunctionFlags.Range, AllowRange.Only, 0); // Returns the k-th percentile of values in a range
        ce.RegisterFunction("PERCENTILE.INC", 2, 2, Adapt(Percentile), FunctionFlags.Range, AllowRange.Only, 0);
        //PERCENTRANK	Returns the percentage rank of a value in a data set
        //PERMUT	Returns the number of permutations for a given number of objects
        //POISSON	Returns the Poisson distribution
        //PROB	Returns the probability that values in a range are between two limits
        ce.RegisterFunction("QUARTILE", 2, 2, Adapt(Quartile), FunctionFlags.Range, AllowRange.Only, 0); // Returns the quartile of a data set
        ce.RegisterFunction("QUARTILE.INC", 2, 2, Adapt(Quartile), FunctionFlags.Range, AllowRange.Only, 0);
        ce.RegisterFunction("RANK", 2, 3, Rank, FunctionFlags.Range, AllowRange.Only, 1); // Returns the rank of a number in a list of numbers
        ce.RegisterFunction("RANK.EQ", 2, 3, Rank, FunctionFlags.Range, AllowRange.Only, 1);
        //RSQ	Returns the square of the Pearson product moment correlation coefficient
        //SKEW	Returns the skewness of a distribution
        //SLOPE	Returns the slope of the linear regression line
        ce.RegisterFunction("SMALL", 2, 2, Adapt(Small), FunctionFlags.Range, AllowRange.Only, 0); // Returns the k-th smallest value in a data set
        //STANDARDIZE	Returns a normalized value
        ce.RegisterFunction("STDEV", 1, int.MaxValue, StDev, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("STDEVA", 1, int.MaxValue, StDevA, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("STDEVP", 1, int.MaxValue, StDevP, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("STDEVPA", 1, int.MaxValue, StDevPA, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("STDEV.S", 1, int.MaxValue, StDev, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("STDEV.P", 1, int.MaxValue, StDevP, FunctionFlags.Range, AllowRange.All);
        //STEYX	Returns the standard error of the predicted y-value for each x in the regression
        //TDIST	Returns the Student's t-distribution
        ce.RegisterFunction("TINV", 2, 2, Adapt(TInv2S), FunctionFlags.Scalar); // Returns the two-tailed inverse of the Student's t-distribution (legacy)
        ce.RegisterFunction("T.INV", 2, 2, Adapt(TInv), FunctionFlags.Scalar); // Returns the left-tailed inverse of the Student's t-distribution
        ce.RegisterFunction("T.INV.2T", 2, 2, Adapt(TInv2S), FunctionFlags.Scalar); // Returns the two-tailed inverse of the Student's t-distribution
        //TREND	Returns values along a linear trend
        //TRIMMEAN	Returns the mean of the interior of a data set
        //TTEST	Returns the probability associated with a Student's t-test
        ce.RegisterFunction("VAR", 1, int.MaxValue, Var, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("VARA", 1, int.MaxValue, VarA, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("VARP", 1, int.MaxValue, VarP, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("VARPA", 1, int.MaxValue, VarPA, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("VAR.S", 1, int.MaxValue, Var, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("VAR.P", 1, int.MaxValue, VarP, FunctionFlags.Range, AllowRange.All);
        //WEIBULL	Returns the Weibull distribution
        //ZTEST	Returns the one-tailed probability-value of a z-test
    }

    private static AnyValue Average(CalcContext ctx, Span<AnyValue> args)
    {
        return Average(ctx, args, TallyNumbers.Default);
    }

    internal static AnyValue Average(CalcContext ctx, Span<AnyValue> args, ITally tally)
    {
        if (args.Length < 1)
            return XLError.IncompatibleValue;

        if (!tally.Tally(ctx, args, new SumState()).TryPickT0(out var state, out var error))
            return error;

        if (state.SampleCount == 0)
            return XLError.DivisionByZero;

        return state.Sum / state.SampleCount;
    }

    private static AnyValue AverageA(CalcContext ctx, Span<AnyValue> args)
    {
        return Average(ctx, args, TallyAll.WithArrayText);
    }

    private static AnyValue BinomDist(double numberSuccesses, double numberTrials, double successProbability, bool cumulativeFlag)
    {
        if (successProbability is < 0 or > 1)
            return XLError.NumberInvalid;

        if (cumulativeFlag)
        {
            var cdf = 0d;
            for (var y = 0; y <= numberSuccesses; ++y)
            {
                var result = BinomDist(y, numberTrials, successProbability);
                if (!result.TryPickT0(out var pf, out var error))
                    return error;

                cdf += pf;
            }

            if (double.IsNaN(cdf) || double.IsInfinity(cdf))
                return XLError.NumberInvalid;

            return cdf;
        }

        {
            var result = BinomDist(numberSuccesses, numberTrials, successProbability);
            if (!result.TryPickT0(out var binomDist, out var error))
                return error;

            return binomDist;
        }
    }

    private static OneOf<double, XLError> BinomDist(double x, double n, double p)
    {
        if (!XLMath.CombinChecked(n, x).TryPickT0(out var combinations, out var error))
            return error;

        x = Math.Floor(x);
        n = Math.Floor(n);
        var binomDist = combinations * Math.Pow(p, x) * Math.Pow(1 - p, n - x);
        if (double.IsNaN(binomDist) || double.IsInfinity(binomDist))
            return XLError.NumberInvalid;

        return binomDist;
    }

    private static AnyValue Count(CalcContext ctx, Span<AnyValue> args)
    {
        return Count(ctx, args, TallyNumbers.IgnoreErrors);
    }

    internal static AnyValue Count(CalcContext ctx, Span<AnyValue> args, ITally tally)
    {
        if (args.Length < 1)
            return XLError.IncompatibleValue;

        var result = tally.Tally(ctx, args, new CountState(0));
        if (!result.TryPickT0(out var state, out var error))
            return error;

        return state.TallyCount;
    }

    private static AnyValue CountA(CalcContext ctx, Span<AnyValue> args)
    {
        return Count(ctx, args, TallyAll.IncludeErrors);
    }

    private static AnyValue CountBlank(CalcContext ctx, AnyValue arg)
    {
        if (!arg.TryPickArea(out var area, out var error))
            return error;

        // To be efficient for cases like whole sheet with only few values, calculate
        // the blank count as number of total area size without non-blank cells.
        var nonBlankCount = ctx.GetNonBlankValues(new Reference(area))
            .LongCount(static value => !value.IsBlank && !(value.IsText && value.GetText().Length == 0));

        return area.Size - nonBlankCount;
    }

    private static AnyValue CountIf(CalcContext ctx, AnyValue countRange, ScalarValue selectionCriteria)
    {
        // Excel doesn't support anything but area in the syntax, but we need to deal with it somehow.
        if (!countRange.TryPickArea(out var countArea, out var areaError))
            return areaError;

        var tally = new TallyCriteria(static _ => 1);
        var criteria = Criteria.Create(selectionCriteria, ctx.Culture);
        tally.Add(countArea, criteria);

        // TallyCriteria only sums up the value.
        var result = tally.Tally(ctx, new[] { countRange }, new CountState(0));
        if (!result.TryPickT0(out var state, out var error))
            return error;

        return state.TallyCount;
    }

    private static AnyValue CountIfs(CalcContext ctx, List<(AnyValue Range, ScalarValue Criteria)> criteriaRanges)
    {
        if (!criteriaRanges[0].Range.TryPickArea(out var countArea, out var areaError))
            return areaError;

        var tally = new TallyCriteria(static _ => 1);
        foreach (var (selectionRange, selectionCriteria) in criteriaRanges)
        {
            var criteria = Criteria.Create(selectionCriteria, ctx.Culture);
            if (!selectionRange.TryPickArea(out var selectionArea, out var selectionAreaError))
                return selectionAreaError;

            // All areas must have same size.
            if (countArea.RowSpan != selectionArea.RowSpan ||
                countArea.ColumnSpan != selectionArea.ColumnSpan)
                return XLError.IncompatibleValue;

            tally.Add(selectionArea, criteria);
        }

        // The values in the range aren't used, so just use first area
        var result = tally.Tally(ctx, new[] { criteriaRanges[0].Range }, new CountState(0));
        if (!result.TryPickT0(out var state, out var error))
            return error;

        return state.TallyCount;
    }

    private static AnyValue AverageIf(CalcContext ctx, AnyValue range, ScalarValue selectionCriteria, AnyValue averageRange)
    {
        // Average range is optional. If not specified, use the criteria range as the average range.
        if (averageRange.IsBlank)
            averageRange = range;

        if (!range.TryPickArea(out var area, out var areaError))
            return areaError;

        if (!averageRange.TryPickArea(out _, out var averageAreaError))
            return averageAreaError;

        var tally = new TallyCriteria();
        tally.Add(area, Criteria.Create(selectionCriteria, ctx.Culture));

        // Average returns #DIV/0! when no cell satisfies the criterion, matching Excel.
        return Average(ctx, new[] { averageRange }, tally);
    }

    private static AnyValue AverageIfs(CalcContext ctx, AnyValue averageRange, List<(AnyValue Range, ScalarValue Criteria)> criteriaRanges)
    {
        if (!TryBuildCriteriaTally(ctx, averageRange, criteriaRanges, out var tally, out var error))
            return error;

        // #DIV/0! when nothing matches, matching Excel.
        return Average(ctx, new[] { averageRange }, tally);
    }

    private static AnyValue MaxIfs(CalcContext ctx, AnyValue maxRange, List<(AnyValue Range, ScalarValue Criteria)> criteriaRanges)
    {
        if (!TryBuildCriteriaTally(ctx, maxRange, criteriaRanges, out var tally, out var error))
            return error;

        // Max returns 0 when nothing matches, matching Excel.
        return Max(ctx, new[] { maxRange }, tally);
    }

    private static AnyValue MinIfs(CalcContext ctx, AnyValue minRange, List<(AnyValue Range, ScalarValue Criteria)> criteriaRanges)
    {
        if (!TryBuildCriteriaTally(ctx, minRange, criteriaRanges, out var tally, out var error))
            return error;

        // Min returns 0 when nothing matches, matching Excel.
        return Min(ctx, new[] { minRange }, tally);
    }

    /// <summary>
    /// Build a <see cref="TallyCriteria"/> for the <c>{AVERAGE,MAX,MIN}IFS</c> family: every
    /// criteria area must match the size of the leading value area (Excel invariant), unlike the
    /// single-criteria <c>*IF</c> forms. Returns <c>false</c> with <paramref name="error"/> set when
    /// an argument isn't an area or the sizes disagree.
    /// </summary>
    private static bool TryBuildCriteriaTally(
        CalcContext ctx,
        AnyValue valueRange,
        List<(AnyValue Range, ScalarValue Criteria)> criteriaRanges,
        out TallyCriteria tally,
        out AnyValue error)
    {
        tally = new TallyCriteria();
        error = default;

        if (!valueRange.TryPickArea(out var valueArea, out var valueAreaError))
        {
            error = valueAreaError;
            return false;
        }

        foreach (var (selectionRange, selectionCriteria) in criteriaRanges)
        {
            if (!selectionRange.TryPickArea(out var selectionArea, out var selectionAreaError))
            {
                error = selectionAreaError;
                return false;
            }

            if (valueArea.RowSpan != selectionArea.RowSpan ||
                valueArea.ColumnSpan != selectionArea.ColumnSpan)
            {
                error = XLError.IncompatibleValue;
                return false;
            }

            tally.Add(selectionArea, Criteria.Create(selectionCriteria, ctx.Culture));
        }

        return true;
    }

    private static AnyValue DevSq(CalcContext ctx, Span<AnyValue> args)
    {
        var result = GetSquareDiffSum(ctx, args, TallyNumbers.Default);
        if (!result.TryPickT0(out var squareDiff, out var error))
            return error;

        // An outlier, most others return #DIV/0! when they can't calculate mean.
        if (squareDiff.SampleCount == 0)
            return XLError.NumberInvalid;

        return squareDiff.Sum;
    }

    private static ScalarValue Fisher(CalcContext ctx, double x)
    {
        if (x is <= -1 or >= 1)
            return XLError.NumberInvalid;

        return 0.5 * Math.Log((1 + x) / (1 - x));
    }

    private static ScalarValue TInv(CalcContext ctx, double probability, double degreesOfFreedom)
    {
        // T.INV: left-tailed inverse. probability in (0, 1), df >= 1.
        if (probability is <= 0 or >= 1)
            return XLError.NumberInvalid;

        var df = Math.Floor(degreesOfFreedom);
        if (df < 1)
            return XLError.NumberInvalid;

        return XLMath.StudentTInv(probability, df);
    }

    private static ScalarValue TInv2S(CalcContext ctx, double probability, double degreesOfFreedom)
    {
        // TINV / T.INV.2T: two-tailed inverse. probability in (0, 1), df >= 1.
        if (probability is <= 0 or >= 1)
            return XLError.NumberInvalid;

        var df = Math.Floor(degreesOfFreedom);
        if (df < 1)
            return XLError.NumberInvalid;

        // Two-tailed: result is always positive.
        // TINV(p, df) = T.INV(1 - p/2, df)
        return XLMath.StudentTInv(1.0 - probability / 2.0, df);
    }

    private static AnyValue GeoMean(CalcContext ctx, Span<AnyValue> args)
    {
        // Rather than interrupting a cycle early, just add it all
        // go through all values anyway. I don't want to code same
        // loop 1000 times and non-positive numbers will be rare.
        var tally = TallyNumbers.Default.Tally(ctx, args, new LogSumState(0.0, 0));
        if (!tally.TryPickT0(out var geoMean, out var error))
            return error;

        if (geoMean.SampleCount == 0)
            return XLError.NumberInvalid;

        // Some value was negative or zero. NaN plus whatever is NaN, infinity
        // plus whatever is also infinity.
        if (double.IsInfinity(geoMean.LogSum) || double.IsNaN(geoMean.LogSum))
            return XLError.NumberInvalid;

        return Math.Exp(geoMean.LogSum / geoMean.SampleCount);
    }

    private static AnyValue Max(CalcContext ctx, Span<AnyValue> args)
    {
        return Max(ctx, args, TallyNumbers.Default);
    }

    internal static AnyValue Max(CalcContext ctx, Span<AnyValue> args, ITally tally)
    {
        var result = tally.Tally(ctx, args, new MaxState());
        if (!result.TryPickT0(out var state, out var error))
            return error;

        if (!state.HasValues)
            return 0;

        return state.MaxValue;
    }

    private static AnyValue MaxA(CalcContext ctx, Span<AnyValue> args)
    {
        return Max(ctx, args, TallyAll.Default);
    }

    private static AnyValue Median(CalcContext ctx, Span<AnyValue> args)
    {
        // There is a better median algorithm that uses two heaps, but NetFx
        // doesn't have heap structure.
        var result = TallyNumbers.Default.Tally(ctx, args, new ValuesState([]));
        if (!result.TryPickT0(out var state, out var error))
            return error;

        var allNumbers = state.Values;
        if (allNumbers.Count == 0)
            return XLError.NumberInvalid;

        allNumbers.Sort();

        var halfIndex = allNumbers.Count / 2;
        var hasEvenCount = allNumbers.Count % 2 == 0;
        if (hasEvenCount)
            return (allNumbers[halfIndex - 1] + allNumbers[halfIndex]) / 2;

        return allNumbers[halfIndex];
    }

    private static AnyValue Min(CalcContext ctx, Span<AnyValue> args)
    {
        return Min(ctx, args, TallyNumbers.Default);
    }

    internal static AnyValue Min(CalcContext ctx, Span<AnyValue> args, ITally tally)
    {
        var result = tally.Tally(ctx, args, new MinState());

        if (!result.TryPickT0(out var state, out var error))
            return error;

        // Not even one non-ignored value found, return 0.
        if (!state.HasValues)
            return 0;

        return state.MinValue;
    }

    private static AnyValue MinA(CalcContext ctx, Span<AnyValue> args)
    {
        return Min(ctx, args, TallyAll.Default);
    }

    private static AnyValue StDev(CalcContext ctx, Span<AnyValue> args)
    {
        return StDev(ctx, args, TallyNumbers.Default);
    }

    internal static AnyValue StDev(CalcContext ctx, Span<AnyValue> args, ITally tally)
    {
        if (!GetSquareDiffSum(ctx, args, tally).TryPickT0(out var squareDiff, out var error))
            return error;

        if (squareDiff.SampleCount <= 1)
            return XLError.DivisionByZero;

        return Math.Sqrt(squareDiff.Sum / (squareDiff.SampleCount - 1));
    }

    private static AnyValue StDevA(CalcContext ctx, Span<AnyValue> args)
    {
        return StDev(ctx, args, TallyAll.Default);
    }

    private static AnyValue StDevP(CalcContext ctx, Span<AnyValue> args)
    {
        return StDevP(ctx, args, TallyNumbers.Default);
    }

    internal static AnyValue StDevP(CalcContext ctx, Span<AnyValue> args, ITally tally)
    {
        if (!GetSquareDiffSum(ctx, args, tally).TryPickT0(out var squareDiff, out var error))
            return error;

        if (squareDiff.SampleCount < 1)
            return XLError.DivisionByZero;

        return Math.Sqrt(squareDiff.Sum / squareDiff.SampleCount);
    }

    private static AnyValue StDevPA(CalcContext ctx, Span<AnyValue> args)
    {
        return StDevP(ctx, args, TallyAll.Default);
    }

    private static AnyValue Var(CalcContext ctx, Span<AnyValue> args)
    {
        return Var(ctx, args, TallyNumbers.Default);
    }

    internal static AnyValue Var(CalcContext ctx, Span<AnyValue> args, ITally tally)
    {
        if (!GetSquareDiffSum(ctx, args, tally).TryPickT0(out var squareDiff, out var error))
            return error;

        if (squareDiff.SampleCount <= 1)
            return XLError.DivisionByZero;

        return squareDiff.Sum / (squareDiff.SampleCount - 1);
    }

    private static AnyValue VarA(CalcContext ctx, Span<AnyValue> args)
    {
        return Var(ctx, args, TallyAll.Default);
    }

    private static AnyValue VarP(CalcContext ctx, Span<AnyValue> args)
    {
        return VarP(ctx, args, TallyNumbers.Default);
    }

    internal static AnyValue VarP(CalcContext ctx, Span<AnyValue> args, ITally tally)
    {
        if (!GetSquareDiffSum(ctx, args, tally).TryPickT0(out var squareDiff, out var error))
            return error;

        if (squareDiff.SampleCount < 1)
            return XLError.DivisionByZero;

        return squareDiff.Sum / squareDiff.SampleCount;
    }

    private static AnyValue VarPA(CalcContext ctx, Span<AnyValue> args)
    {
        return VarP(ctx, args, TallyAll.Default);
    }

    private static AnyValue Large(CalcContext ctx, AnyValue arrayParam, double kParam)
    {
        if (kParam < 1)
            return XLError.NumberInvalid;

        var k = (int)Math.Ceiling(kParam);
        if (!TryGetNumbers(ctx, arrayParam, out var total, out var error))
            return error;

        if (k > total.Count)
            return XLError.NumberInvalid;

        total.Sort();

        // k-th largest.
        return total[^k];
    }

    private static AnyValue Small(CalcContext ctx, AnyValue arrayParam, double kParam)
    {
        if (kParam < 1)
            return XLError.NumberInvalid;

        var k = (int)Math.Ceiling(kParam);
        if (!TryGetNumbers(ctx, arrayParam, out var total, out var error))
            return error;

        if (k > total.Count)
            return XLError.NumberInvalid;

        total.Sort();

        // k-th smallest.
        return total[k - 1];
    }

    private static AnyValue Rank(CalcContext ctx, Span<AnyValue> args)
    {
        // RANK(number, ref, [order]). number/order are scalars (implicitly intersected); ref is the
        // range/array (marked param 1). order = 0 or omitted ranks descending, non-zero ascending.
        if (!args[0].TryPickScalar(out var numberScalar, out _))
            return XLError.IncompatibleValue;
        if (!numberScalar.ToNumber(ctx.Culture).TryPickT0(out var number, out var numberError))
            return numberError;

        if (!TryGetNumbers(ctx, args[1], out var numbers, out var refError))
            return refError;

        var ascending = false;
        if (args.Length > 2)
        {
            if (!args[2].TryPickScalar(out var orderScalar, out _))
                return XLError.IncompatibleValue;
            if (!orderScalar.ToNumber(ctx.Culture).TryPickT0(out var order, out var orderError))
                return orderError;
            ascending = order != 0;
        }

        if (!numbers.Contains(number))
            return XLError.NoValueAvailable;

        // Tied values share the top rank of their group.
        var rank = ascending
            ? numbers.Count(v => v < number) + 1
            : numbers.Count(v => v > number) + 1;
        return rank;
    }

    private static AnyValue Mode(CalcContext ctx, Span<AnyValue> args)
    {
        var result = TallyNumbers.Default.Tally(ctx, args, new ValuesState([]));
        if (!result.TryPickT0(out var state, out var error))
            return error;

        var values = state.Values;
        var counts = new Dictionary<double, int>();
        var maxCount = 0;
        foreach (var value in values)
        {
            var count = counts.GetValueOrDefault(value) + 1;
            counts[value] = count;
            if (count > maxCount)
                maxCount = count;
        }

        // No value repeats (or there are no numbers) -> no mode.
        if (maxCount <= 1)
            return XLError.NoValueAvailable;

        // Among the tied modes, return the one whose first occurrence is earliest (Excel order).
        foreach (var value in values)
        {
            if (counts[value] == maxCount)
                return value;
        }

        return XLError.NoValueAvailable;
    }

    private static AnyValue Percentile(CalcContext ctx, AnyValue arrayParam, double k)
    {
        if (!TryGetNumbers(ctx, arrayParam, out var numbers, out var error))
            return error;

        return PercentileInclusive(numbers, k);
    }

    private static AnyValue Quartile(CalcContext ctx, AnyValue arrayParam, double quartParam)
    {
        if (!TryGetNumbers(ctx, arrayParam, out var numbers, out var error))
            return error;

        // Excel truncates the quart argument toward zero and accepts only 0..4.
        var quart = (int)quartParam;
        if (quart < 0 || quart > 4)
            return XLError.NumberInvalid;

        return PercentileInclusive(numbers, quart * 0.25);
    }

    /// <summary>
    /// PERCENTILE.INC over a materialized list: the <paramref name="k"/>-th percentile
    /// (<c>k</c> in <c>[0, 1]</c>) with linear interpolation between the two closest ranks.
    /// </summary>
    private static AnyValue PercentileInclusive(List<double> numbers, double k)
    {
        if (numbers.Count == 0 || k < 0 || k > 1)
            return XLError.NumberInvalid;

        numbers.Sort();

        var rank = k * (numbers.Count - 1);
        var low = (int)Math.Floor(rank);
        if (low + 1 >= numbers.Count)
            return numbers[low];

        var fraction = rank - low;
        return numbers[low] + fraction * (numbers[low + 1] - numbers[low]);
    }

    /// <summary>
    /// Collect the numeric values of an array/range/scalar argument (skipping blanks and text,
    /// short-circuiting on an error value), mirroring how <see cref="Large"/> reads its data set.
    /// </summary>
    private static bool TryGetNumbers(CalcContext ctx, AnyValue arrayParam, out List<double> numbers, out XLError error)
    {
        error = default;

        if (arrayParam.TryPickScalar(out var scalar, out var collection))
        {
            if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var number, out var scalarError))
            {
                numbers = null!;
                error = scalarError;
                return false;
            }

            numbers = new List<double>(1) { number };
            return true;
        }

        IEnumerable<ScalarValue> values;
        int size;
        if (collection.TryPickT0(out var array, out var reference))
        {
            values = array;
            size = array.Width * array.Height;
        }
        else
        {
            values = reference.GetCellsValues(ctx);
            size = reference.NumberOfCells;
        }

        // Pre-allocate to reduce allocations during doubling of the buffer.
        var total = new List<double>(size);
        foreach (var value in values)
        {
            if (value.IsError)
            {
                numbers = null!;
                error = value.GetError();
                return false;
            }

            if (value.IsNumber)
                total.Add(value.GetNumber());
        }

        numbers = total;
        return true;
    }

    /// <summary>
    /// Calculate <c>SUM((x_i - mean_x)^2)</c> and number of samples. This method uses two-pass algorithm.
    /// There are several one-pass algorithms, but they are not numerically stable. In this case, accuracy
    /// takes precedence (plus VAR/STDEV are not a very frequently used function). Excel might have used
    /// those one-pass formulas in the past (see <em>Statistical flaws in Excel</em>), but doesn't seem to
    /// be using them anymore.
    /// </summary>
    private static OneOf<SquareDiff, XLError> GetSquareDiffSum(CalcContext ctx, Span<AnyValue> args, ITally tally)
    {
        if (!tally.Tally(ctx, args, new SumState(0.0, 0)).TryPickT0(out var sumState, out var sumError))
            return sumError;

        if (sumState.SampleCount == 0)
            return new SquareDiff(0.0, 0, double.NaN);

        var sampleMean = sumState.Sum / sumState.SampleCount;

        // Calculate sum of squares of deviations from sample mean
        var initialSquareDiffState = new SquareDiff(Sum: 0.0, SampleCount: 0, SampleMean: sampleMean);
        var result = tally.Tally(ctx, args, initialSquareDiffState);

        if (!result.TryPickT0(out var squareDiff, out var error))
            return error;

        return squareDiff;
    }

    private readonly record struct SumState(double Sum, int SampleCount) : ITallyState<SumState>
    {
        public SumState Tally(double number) => new(Sum + number, SampleCount + 1);
    }

    private readonly record struct SquareDiff(double Sum, int SampleCount, double SampleMean) : ITallyState<SquareDiff>
    {
        public SquareDiff Tally(double number)
        {
            var diff = number - SampleMean;
            var sum = Sum + diff * diff;
            return new SquareDiff(sum, SampleCount + 1, SampleMean);
        }
    }

    private readonly record struct MinState(double MinValue, bool HasValues) : ITallyState<MinState>
    {
        public MinState() : this(double.MaxValue, false)
        {
        }

        public MinState Tally(double number) => new(Math.Min(MinValue, number), true);
    }

    private readonly record struct MaxState(double MaxValue, bool HasValues) : ITallyState<MaxState>
    {
        public MaxState() : this(double.MinValue, false)
        {
        }

        public MaxState Tally(double number) => new(Math.Max(MaxValue, number), true);
    }

    private readonly record struct LogSumState(double LogSum, int SampleCount) : ITallyState<LogSumState>
    {
        public LogSumState Tally(double number)
        {
            var logSum = LogSum + Math.Log(number);
            return new LogSumState(logSum, SampleCount + 1);
        }
    }

    private readonly record struct ValuesState(List<double> Values) : ITallyState<ValuesState>
    {
        public ValuesState Tally(double number)
        {
            Values.Add(number);
            return new ValuesState(Values);
        }
    }

    private readonly record struct CountState(int TallyCount) : ITallyState<CountState>
    {
        public CountState Tally(double number) => new(TallyCount + 1);
    }
}
