using System;
using System.Collections.Generic;
using static XLibur.Excel.CalcEngine.Functions.SignatureAdapter;

#pragma warning disable S1244 // Intentional exact float comparison for Excel formula compatibility

namespace XLibur.Excel.CalcEngine;

internal static class Financial
{
    public static void Register(FunctionRegistry ce)
    {
        // The day-count-basis bond family (ACCRINT, ACCRINTM, COUP*, DURATION, MDURATION, ODD*,
        // PRICE*, YIELD*, AMORDEGRC, AMORLINC) needs a full 30/360 & actual/actual coupon-period
        // engine and is deliberately left out — see spec 07, "wave A2".
        ce.RegisterFunction("CUMIPMT", 6, 6, Adapt(CumIpmt), FunctionFlags.Scalar); // Returns the cumulative interest paid between two periods
        ce.RegisterFunction("CUMPRINC", 6, 6, Adapt(CumPrinc), FunctionFlags.Scalar); // Returns the cumulative principal paid on a loan between two periods
        ce.RegisterFunction("DB", 4, 5, AdaptLastOptional(Db, 12), FunctionFlags.Scalar); // Returns the depreciation of an asset for a specified period by using the fixed-declining balance method
        ce.RegisterFunction("DDB", 4, 5, AdaptLastOptional(Ddb, 2), FunctionFlags.Scalar); // Returns the depreciation of an asset for a specified period by using the double-declining balance method or some other method that you specify
        ce.RegisterFunction("DISC", 4, 5, AdaptLastOptional(Disc, 0), FunctionFlags.Scalar); // Returns the discount rate for a security
        ce.RegisterFunction("DOLLARDE", 2, 2, Adapt(DollarDe), FunctionFlags.Scalar); // Converts a dollar price, expressed as a fraction, into a dollar price, expressed as a decimal number
        ce.RegisterFunction("DOLLARFR", 2, 2, Adapt(DollarFr), FunctionFlags.Scalar); // Converts a dollar price, expressed as a decimal number, into a dollar price, expressed as a fraction
        ce.RegisterStrictNumeric("EFFECT", 2, 2, Adapt(Effect), [0, 1], FunctionFlags.Scalar); // Returns the effective annual interest rate
        ce.RegisterFunction("FV", 3, 5, AdaptLastTwoOptional(Fv, 0, 0), FunctionFlags.Scalar); // Returns the future value of an investment
        ce.RegisterFunction("FVSCHEDULE", 2, 2, FvSchedule, FunctionFlags.Range, AllowRange.Only, 1); // Returns the future value of an initial principal after applying a series of compound interest rates
        ce.RegisterFunction("INTRATE", 4, 5, AdaptLastOptional(IntRate, 0), FunctionFlags.Scalar); // Returns the interest rate for a fully invested security
        ce.RegisterFunction("IPMT", 4, 6, AdaptLastTwoOptional(Ipmt, 0, 0), FunctionFlags.Scalar); // Returns the interest payment for an investment for a given period
        ce.RegisterFunction("IRR", 1, 2, Irr, FunctionFlags.Range, AllowRange.Only, 0); // Returns the internal rate of return for a series of cash flows
        ce.RegisterFunction("ISPMT", 4, 4, Adapt(IsPmt), FunctionFlags.Scalar); // Calculates the interest paid during a specific period of an investment
        ce.RegisterFunction("MIRR", 3, 3, Mirr, FunctionFlags.Range, AllowRange.Only, 0); // Returns the internal rate of return where positive and negative cash flows are financed at different rates
        ce.RegisterStrictNumeric("NOMINAL", 2, 2, Adapt(Nominal), [0, 1], FunctionFlags.Scalar); // Returns the annual nominal interest rate
        ce.RegisterFunction("NPER", 3, 5, AdaptLastTwoOptional(Nper, 0, 0), FunctionFlags.Scalar); // Returns the number of periods for an investment
        ce.RegisterFunction("NPV", 2, 255, Npv, FunctionFlags.Range, AllowRange.Except, 0); // Returns the net present value of an investment based on a series of periodic cash flows and a discount rate
        ce.RegisterFunction("PDURATION", 3, 3, Adapt(PDuration), FunctionFlags.Scalar | FunctionFlags.Future); // Returns the number of periods required by an investment to reach a specified value
        ce.RegisterFunction("PMT", 3, 5, AdaptLastTwoOptional(Pmt, 0, 0), FunctionFlags.Scalar); // Returns the periodic payment for an annuity
        ce.RegisterFunction("PPMT", 4, 6, AdaptLastTwoOptional(Ppmt, 0, 0), FunctionFlags.Scalar); // Returns the payment on the principal for an investment for a given period
        ce.RegisterFunction("PV", 3, 5, AdaptLastTwoOptional(Pv, 0, 0), FunctionFlags.Scalar); // Returns the present value of an investment
        ce.RegisterFunction("RATE", 3, 6, Rate, FunctionFlags.Scalar); // Returns the interest rate per period of an annuity
        ce.RegisterFunction("RECEIVED", 4, 5, AdaptLastOptional(Received, 0), FunctionFlags.Scalar); // Returns the amount received at maturity for a fully invested security
        ce.RegisterFunction("RRI", 3, 3, Adapt(Rri), FunctionFlags.Scalar | FunctionFlags.Future); // Returns an equivalent interest rate for the growth of an investment
        ce.RegisterFunction("SLN", 3, 3, Adapt(Sln), FunctionFlags.Scalar); // Returns the straight-line depreciation of an asset for one period
        ce.RegisterFunction("SYD", 4, 4, Adapt(Syd), FunctionFlags.Scalar); // Returns the sum-of-years' digits depreciation of an asset for a specified period
        ce.RegisterFunction("TBILLEQ", 3, 3, Adapt(TBillEq), FunctionFlags.Scalar); // Returns the bond-equivalent yield for a Treasury bill
        ce.RegisterFunction("TBILLPRICE", 3, 3, Adapt(TBillPrice), FunctionFlags.Scalar); // Returns the price per $100 face value for a Treasury bill
        ce.RegisterFunction("TBILLYIELD", 3, 3, Adapt(TBillYield), FunctionFlags.Scalar); // Returns the yield for a Treasury bill
        ce.RegisterFunction("VDB", 5, 7, AdaptLastTwoOptional(Vdb, 2, false), FunctionFlags.Scalar); // Returns the depreciation of an asset for a specified or partial period by using a declining balance method
        ce.RegisterFunction("XIRR", 2, 3, XIrr, FunctionFlags.Range, AllowRange.Only, 0, 1); // Returns the internal rate of return for a schedule of cash flows that is not necessarily periodic
        ce.RegisterFunction("XNPV", 3, 3, XNpv, FunctionFlags.Range, AllowRange.Only, 1, 2); // Returns the net present value for a schedule of cash flows that is not necessarily periodic
    }

    private static AnyValue Fv(double rate, double numberOfPayments, double pmt, double presentValue, double type)
    {
        if (numberOfPayments == 0)
            return -presentValue;

        return FvInternal(rate, numberOfPayments, pmt, presentValue, type);
    }

    private static double FvInternal(double rate, double numberOfPayments, double pmt, double presentValue, double type)
    {
        if (rate == 0.0)
            return -(pmt * numberOfPayments + presentValue);

        if (type != 0.0)
            pmt *= (1 + rate);

        return -(pmt * (Math.Pow(1 + rate, numberOfPayments) - 1) / rate + presentValue * Math.Pow(1 + rate, numberOfPayments));
    }

    private static AnyValue Ipmt(double rate, double period, double numberOfPayments, double presentValue, double futureValue, double type)
    {
        if (!IsPaymentPeriod(rate, period, numberOfPayments))
            return XLError.NumberInvalid;

        // InterestOfPeriod ignores the future value; the payment passed in carries it.
        return InterestOfPeriod(rate, period, presentValue, type, PmtInternal(rate, Math.Ceiling(numberOfPayments), presentValue, futureValue, type));
    }

    private static bool IsPaymentPeriod(double rate, double period, double numberOfPayments)
        => numberOfPayments > 0 && rate > -1 && period >= 1 && period <= Math.Ceiling(numberOfPayments);

    private static AnyValue Pmt(double rate, double numberOfPayments, double presentValue, double futureValue, double type)
    {
        if (numberOfPayments == 0 || rate <= -1)
            return XLError.NumberInvalid;

        return PmtInternal(rate, numberOfPayments, presentValue, futureValue, type);
    }

    private static double PmtInternal(double rate, double numberOfPayments, double presentValue, double futureValue, double type)
    {
        if (rate == 0.0)
            return -(presentValue + futureValue) / numberOfPayments;

        var (_, oneMinusDiscount) = DiscountFactor(rate, numberOfPayments);
        var payment = -rate * (AmortizedValue(presentValue, futureValue, oneMinusDiscount) - futureValue);
        return payment / (1 + rate * PaymentTiming(type));
    }

    /// <summary>
    /// <c>(pv + fv) / (1 - v)</c>, the term PMT and PPMT share. Excel computes it before scaling by the
    /// rate, so it overflows to <c>#NUM!</c> for PMT(0.05,10,9E307) and PMT(0.05,10,1000,9E307).
    /// </summary>
    private static double AmortizedValue(double presentValue, double futureValue, double oneMinusDiscount)
        => (presentValue + futureValue) / oneMinusDiscount;

    /// <summary>
    /// The discount factor <c>v = (1+rate)^-n</c> and <c>1 - v</c>, both from <c>n·ln(1+rate)</c>, so that
    /// <c>1 - v</c> keeps its digits for a tiny rate (PMT(1E-10,10,1000) is -100.000000055) and a
    /// growth that overflows only drives <c>v</c> to zero (PMT(0.05,1000000,1000) is -50).
    /// </summary>
    private static (double Discount, double OneMinusDiscount) DiscountFactor(double rate, double numberOfPayments)
    {
        var logGrowth = numberOfPayments * LogOnePlus(rate);
        return (Math.Exp(-logGrowth), -ExpMinusOne(-logGrowth));
    }

    private static AnyValue Pv(double rate, double numberOfPayments, double pmt, double futureValue, double type)
    {
        if (rate == 0.0)
            return -(futureValue + pmt * numberOfPayments);

        var pow = Math.Pow(1 + rate, numberOfPayments);
        if (pow == 0.0)
            return XLError.DivisionByZero;

        return -(futureValue + pmt * (1 + rate * PaymentTiming(type)) * (pow - 1) / rate) / pow;
    }

    private static double PaymentTiming(double type) => type != 0.0 ? 1 : 0;

    private static AnyValue Nper(double rate, double pmt, double presentValue, double futureValue, double type)
    {
        if (rate <= -1)
            return XLError.NumberInvalid;

        if (rate == 0.0)
        {
            if (pmt == 0.0)
                return XLError.NumberInvalid;

            return -(presentValue + futureValue) / pmt;
        }

        var timing = pmt * (1 + rate * PaymentTiming(type));
        var numerator = timing - futureValue * rate;
        var denominator = presentValue * rate + timing;
        if (denominator == 0.0 || numerator / denominator <= 0.0)
            return XLError.NumberInvalid;

        return Math.Log(numerator / denominator) / Math.Log(1 + rate);
    }

    private static AnyValue Ppmt(double rate, double period, double numberOfPayments, double presentValue, double futureValue, double type)
    {
        if (!IsPaymentPeriod(rate, period, numberOfPayments))
            return XLError.NumberInvalid;

        if (rate == 0.0)
            return -(presentValue + futureValue) / numberOfPayments;

        var annuityDue = type != 0.0;
        if (annuityDue && period == 1)
            return PmtInternal(rate, numberOfPayments, presentValue, futureValue, type);

        // Principal grows by (1+rate) a period from the first end-of-period payment's -r(pv+fv)v/(1-v).
        // Computed directly rather than as PMT - IPMT, which cancels: PPMT(1E15,1,10,1000) is
        // -1E-132 and PPMT(9E307,1,10,1000) is 0. An annuity-due's later periods are a period behind.
        var (discount, oneMinusDiscount) = DiscountFactor(rate, numberOfPayments);
        var firstPrincipal = -AmortizedValue(presentValue, futureValue, oneMinusDiscount) * discount * rate;
        return firstPrincipal * Math.Pow(1 + rate, period - (annuityDue ? 2 : 1));
    }

    private static AnyValue Npv(CalcContext ctx, Span<AnyValue> args)
    {
        // NPV(rate, value1, [value2], ...). rate is a scalar (marked param 0), values may be ranges.
        if (!args[0].TryReduceToNumber(ctx, out var rate, out var rateError))
            return rateError;
        if (rate < -1)
            return XLError.NumberInvalid;

        double npv = 0;
        var period = 1;
        for (var i = 1; i < args.Length; i++)
        {
            if (!TryAddDiscountedValues(ctx, args[i], rate, ref npv, ref period, out var valueError))
                return valueError;
        }

        return npv;
    }

    /// <summary>
    /// Add the numbers of one NPV value argument to <paramref name="npv"/>, each discounted by its
    /// sequential <paramref name="period"/>. The first error stops the walk.
    /// </summary>
    private static bool TryAddDiscountedValues(CalcContext ctx, in AnyValue arg, double rate, ref double npv, ref int period, out XLError error)
    {
        // A value given directly coerces like any number argument: NPV(0.1,"5") discounts 5 and
        // NPV(0.1,"x") is #VALUE!.
        if (arg.TryPickScalar(out var direct, out _))
        {
            if (!direct.ToNumber(ctx.Culture).TryPickT0(out var number, out error))
                return false;

            return TryAddDiscounted(number, rate, ref npv, ref period, out error);
        }

        foreach (var scalar in EnumerateScalars(ctx, arg))
        {
            if (scalar.IsError)
            {
                error = scalar.GetError();
                return false;
            }

            // NPV ignores blanks, text and logicals in references and arrays.
            if (scalar.IsNumber && !TryAddDiscounted(scalar.GetNumber(), rate, ref npv, ref period, out error))
                return false;
        }

        error = default;
        return true;
    }

    /// <summary>
    /// Discount a value by its sequential position. A divisor of zero (rate -1) or one that overflows
    /// (NPV(9E307,...) at the second period) is <c>#DIV/0!</c> in Excel.
    /// </summary>
    private static bool TryAddDiscounted(double value, double rate, ref double npv, ref int period, out XLError error)
    {
        var divisor = Math.Pow(1 + rate, period);
        if (divisor == 0.0 || double.IsInfinity(divisor))
        {
            error = XLError.DivisionByZero;
            return false;
        }

        npv += value / divisor;
        period++;
        error = default;
        return true;
    }

    private static AnyValue Irr(CalcContext ctx, Span<AnyValue> args)
    {
        // IRR(values, [guess]). values is the cash-flow range (marked param 0), starting at period 0.
        if (!TryCollectNumbers(ctx, args[..1], out var cashflows, out var valuesError))
            return valuesError;
        if (cashflows.Count < 2)
            return XLError.NumberInvalid;

        var guess = 0.1;
        if (args.Length > 1 && !args[1].TryReduceToNumber(ctx, out guess, out var guessError))
            return guessError;

        return SolveIrr(cashflows, guess);
    }

    /// <summary>
    /// Excel answers <c>#VALUE!</c> for a guess at or below -1 or above about 2^53 (measured: 9E15
    /// is solved, 9.0072E15 is not).
    /// <para>
    /// When Newton's method fails, a guess of 1 or more falls back to bisection: Excel finds
    /// IRR(-100,50,70) = 0.123 from a guess of 5 or 1E15. A smaller guess does not, because Excel
    /// answers <c>#NUM!</c> for IRR(-100,0,...,0,1) from the default guess of 0.1.
    /// </para>
    /// </summary>
    private static AnyValue SolveIrr(List<double> cashflows, double guess)
    {
        const double maxGuess = 9007199254740992d; // 2^53
        if (guess <= -1 || guess > maxGuess)
            return XLError.IncompatibleValue;

        // A rate of return only exists when there is both an inflow and an outflow.
        if (!cashflows.Exists(c => c > 0) || !cashflows.Exists(c => c < 0))
            return XLError.NumberInvalid;

        if (TryNewtonRoot(rate => NpvWithDerivative(cashflows, rate), guess, out var root) || (guess >= 1 && TryBisectNpv(cashflows, out root)))
            return root;

        return XLError.NumberInvalid;
    }

    private static (double Value, double Derivative) NpvWithDerivative(List<double> cashflows, double rate)
    {
        double npv = 0, derivative = 0;
        for (var t = 0; t < cashflows.Count; t++)
        {
            var factor = Math.Pow(1 + rate, t);
            npv += cashflows[t] / factor;
            derivative -= t * cashflows[t] / (factor * (1 + rate));
        }

        return (npv, derivative);
    }

    private static bool TryNewtonRoot(Func<double, (double Value, double Derivative)> function, double guess, out double root)
    {
        const int maxIterations = 100;
        const double relativeStep = 1e-12;
        root = guess;
        for (var iteration = 0; iteration < maxIterations; iteration++)
        {
            var (value, derivative) = function(root);
            if (value == 0.0)
                return true;

            var step = value / derivative;
            if (!double.IsFinite(step))
                return false;

            root -= step;
            if (Math.Abs(step) <= relativeStep * Math.Max(1, Math.Abs(root)))
                return true;
        }

        return false;
    }

    private static bool TryBisectNpv(List<double> cashflows, out double root)
    {
        for (var d = 0; d < 30; d++)
        {
            if (TryBisectNpv(cashflows, Math.ScaleB(1, d) - 1, Math.ScaleB(1, d + 1) - 1, out root)
                || TryBisectNpv(cashflows, Math.ScaleB(1, -d - 1) - 1, Math.ScaleB(1, -d) - 1, out root))
                return true;
        }

        root = default;
        return false;
    }

    private static bool TryBisectNpv(List<double> cashflows, double low, double high, out double root)
    {
        root = default;
        var lowValue = NpvWithDerivative(cashflows, low).Value;
        var highValue = NpvWithDerivative(cashflows, high).Value;
        if (!double.IsFinite(lowValue) || !double.IsFinite(highValue) || Math.Sign(lowValue) == Math.Sign(highValue))
            return false;

        while (true)
        {
            var middle = low + (high - low) / 2;
            if (middle <= low || middle >= high)
            {
                root = Math.Abs(lowValue) <= Math.Abs(highValue) ? low : high;
                return true;
            }

            var value = NpvWithDerivative(cashflows, middle).Value;
            if (value == 0.0)
            {
                root = middle;
                return true;
            }

            if (Math.Sign(value) == Math.Sign(lowValue))
                (low, lowValue) = (middle, value);
            else
                (high, highValue) = (middle, value);
        }
    }

    private static AnyValue Rate(CalcContext ctx, Span<AnyValue> args)
    {
        // RATE(nper, pmt, pv, [fv], [type], [guess]) - solved iteratively. All arguments are scalars.
        if (!args[0].TryReduceToNumber(ctx, out var nper, out var nperError))
            return nperError;
        if (!args[1].TryReduceToNumber(ctx, out var pmt, out var pmtError))
            return pmtError;
        if (!args[2].TryReduceToNumber(ctx, out var pv, out var pvError))
            return pvError;

        if (!TryOptionalScalarNumber(ctx, args, 3, 0, out var fv, out var fvError))
            return fvError;
        if (!TryOptionalScalarNumber(ctx, args, 4, 0, out var type, out var typeError))
            return typeError;
        if (!TryOptionalScalarNumber(ctx, args, 5, 0.1, out var guess, out var guessError))
            return guessError;

        return SolveRate(nper, pmt, pv, fv, type, guess);
    }

    /// <summary>
    /// Read the optional scalar number at <paramref name="index"/>, or <paramref name="defaultValue"/>
    /// when the argument is left out.
    /// </summary>
    private static bool TryOptionalScalarNumber(CalcContext ctx, Span<AnyValue> args, int index, double defaultValue, out double value, out XLError error)
    {
        value = defaultValue;
        error = default;
        return args.Length <= index || args[index].TryReduceToNumber(ctx, out value, out error);
    }

    /// <summary>
    /// Excel answers <c>#VALUE!</c> for a guess at or below -1, and <c>#NUM!</c> for no periods or
    /// a rate at or below -1.
    /// </summary>
    private static AnyValue SolveRate(double nper, double pmt, double pv, double fv, double type, double guess)
    {
        if (guess <= -1)
            return XLError.IncompatibleValue;

        // With no periods the equation is pv + fv at every rate, so there is no single rate.
        if (nper == 0)
            return XLError.NumberInvalid;

        if (!TryNewtonRoot(rate => TvmEquation(rate, nper, pmt, pv, fv, type), guess, out var root) || root <= -1)
            return XLError.NumberInvalid;

        return root;
    }

    private static (double Value, double Derivative) TvmEquation(double rate, double nper, double pmt, double pv, double fv, double type)
    {
        var (growth, annuity, annuityDerivative) = GrowthAndAnnuity(rate, nper);
        var value = pv * growth + pmt * (1 + rate * type) * annuity + fv;
        var derivative = pv * nper * Math.Pow(1 + rate, nper - 1) + pmt * (type * annuity + (1 + rate * type) * annuityDerivative);
        return (value, derivative);
    }

    private static (double Growth, double Annuity, double AnnuityDerivative) GrowthAndAnnuity(double rate, double nper)
    {
        const double binomialSeriesRadius = 1e-5;
        if (Math.Abs(rate) * Math.Max(1, Math.Abs(nper)) < binomialSeriesRadius)
            return BinomialSeriesGrowthAndAnnuity(rate, nper);

        var (growth, growthMinusOne) = GrowthWithoutCancellation(rate, nper);
        var annuity = growthMinusOne / rate;
        return (growth, annuity, (nper * Math.Pow(1 + rate, nper - 1) - annuity) / rate);
    }

    private static (double Growth, double Annuity, double AnnuityDerivative) BinomialSeriesGrowthAndAnnuity(double rate, double nper)
    {
        var c2 = nper * (nper - 1) / 2;
        var c3 = c2 * (nper - 2) / 3;
        var c4 = c3 * (nper - 3) / 4;
        var annuity = nper + rate * (c2 + rate * (c3 + rate * c4));
        return (1 + rate * annuity, annuity, c2 + rate * (2 * c3 + rate * 3 * c4));
    }

    private static (double Growth, double GrowthMinusOne) GrowthWithoutCancellation(double rate, double nper)
    {
        const double cancellingLogGrowth = 0.5;
        var logGrowth = rate > -1 ? nper * LogOnePlus(rate) : double.NaN;
        if (Math.Abs(logGrowth) <= cancellingLogGrowth)
        {
            var growthMinusOne = ExpMinusOne(logGrowth);
            return (1 + growthMinusOne, growthMinusOne);
        }

        var growth = Math.Pow(1 + rate, nper);
        return (growth, growth - 1);
    }

    /// <summary>
    /// <c>ln(1+x)</c> accurate for small <c>x</c>; <see cref="double.LogP1"/> is <c>Math.Log(x + 1)</c>, which
    /// loses the digits of <c>x</c> to rounding.
    /// </summary>
    private static double LogOnePlus(double x)
    {
        var u = 1 + x;
        return u == 1 ? x : Math.Log(u) * x / (u - 1);
    }

    /// <summary>
    /// <c>e^x-1</c> accurate for small <c>x</c>; <see cref="double.ExpM1"/> is <c>Math.Exp(x) - 1</c>, which
    /// cancels.
    /// </summary>
    private static double ExpMinusOne(double x)
    {
        var u = Math.Exp(x);
        if (u == 1)
            return x;

        return Math.Abs(x) < 1 ? (u - 1) * x / Math.Log(u) : u - 1;
    }

    #region Depreciation

    private static ScalarValue Sln(CalcContext ctx, double cost, double salvage, double life)
    {
        if (life == 0)
            return XLError.DivisionByZero;

        return (cost - salvage) / life;
    }

    private static ScalarValue Syd(CalcContext ctx, double cost, double salvage, double life, double period)
    {
        if (life <= 0 || salvage < 0)
            return XLError.NumberInvalid;
        if (period <= 0 || period > life)
            return XLError.NumberInvalid;

        // A sum of the years' digits that overflows is #DIV/0! in Excel, as SYD(1000,100,9E307,1) is.
        var yearsDigits = life * (life + 1);
        if (double.IsInfinity(yearsDigits))
            return XLError.DivisionByZero;

        return (cost - salvage) * (life - period + 1) * 2 / yearsDigits;
    }

    /// <summary>
    /// Fixed-declining balance depreciation. The rate is <c>1 - (salvage/cost)^(1/life)</c> rounded
    /// to three decimals (Excel rounds before applying it, which is why DB and DDB disagree), and
    /// the first and the optional stub period <c>life+1</c> are pro-rated by <paramref name="month"/>.
    /// </summary>
    private static ScalarValue Db(CalcContext ctx, double cost, double salvage, double life, double period, double month)
    {
        month = Math.Truncate(month);
        if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || month < 1 || month > 12)
            return XLError.NumberInvalid;

        // A stub period after the asset's life only exists when the first year was partial.
        var lastPeriod = month < 12 ? life + 1 : life;
        if (period > lastPeriod)
            return XLError.NumberInvalid;

        if (cost == 0)
            return 0d;

        var rate = Math.Round(1 - Math.Pow(salvage / cost, 1 / life), 3, MidpointRounding.AwayFromZero);

        // Excel scales by the fraction of a year, so DB(9E307,100,10,1) is 9E307 rather than an overflow.
        var firstPeriod = cost * rate * (month / 12);
        if (period == 1)
            return firstPeriod;

        var accumulated = firstPeriod;
        var wholePeriods = Math.Min(Math.Truncate(period), Math.Truncate(life));
        double current = firstPeriod;
        for (var i = 2d; i <= wholePeriods; i++)
        {
            current = (cost - accumulated) * rate;
            accumulated += current;
        }

        if (period <= life)
            return current;

        // Stub period: what is left of the year the asset ran past its life.
        return (cost - accumulated) * rate * (12 - month) / 12;
    }

    private static ScalarValue Ddb(CalcContext ctx, double cost, double salvage, double life, double period, double factor)
    {
        if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || period > life || factor <= 0)
            return XLError.NumberInvalid;

        return DdbPeriod(cost, salvage, life, period, factor);
    }

    /// <summary>
    /// Declining-balance depreciation of a single period, expressed in closed form so that a
    /// fractional <paramref name="period"/> works the same way Excel's does. The charge is the opening
    /// book value times the rate, not the difference of two book values, which cancels
    /// (DDB(1000,100,10,1,1E-10) is 1E-08).
    /// </summary>
    private static double DdbPeriod(double cost, double salvage, double life, double period, double factor)
    {
        var rate = factor / life;
        var openingValue = rate >= 1
            ? period == 1 ? cost : 0
            : cost * Math.Pow(1 - rate, period - 1);

        // Excel charges the uncapped rate before the salvage cap, so DDB(1000,100,10,1,9E307) overflows.
        var charge = openingValue * rate;
        if (double.IsInfinity(charge))
            return charge;

        // Depreciation stops once the book value would drop below the salvage value.
        return Math.Max(Math.Min(charge, openingValue - salvage), 0);
    }

    /// <summary>
    /// Depreciation charged between <paramref name="startPeriod"/> and <paramref name="endPeriod"/>,
    /// as the difference of two cumulative runs from period zero. Replaying from the start is what
    /// makes the straight-line cross-over land in the same place regardless of which slice is asked
    /// for, so consecutive slices add up to the whole and a full-life run charges exactly
    /// <c>cost - salvage</c>.
    /// </summary>
#pragma warning disable S107 // VDB takes seven arguments in Excel; the eighth is the calc context
    private static ScalarValue Vdb(CalcContext ctx, double cost, double salvage, double life, double startPeriod, double endPeriod, double factor, bool noSwitch)
#pragma warning restore S107
    {
        if (startPeriod < 0 || endPeriod < startPeriod || endPeriod > life || cost < 0 || salvage > cost || factor <= 0)
            return XLError.NumberInvalid;

        return VdbCumulative(cost, salvage, life, endPeriod, factor, noSwitch)
               - VdbCumulative(cost, salvage, life, startPeriod, factor, noSwitch);
    }

    /// <summary>
    /// Total depreciation from period zero up to <paramref name="upTo"/>, which may be fractional.
    /// Each period is charged on the declining balance until straight-line over the remaining life
    /// of the remaining basis is the larger charge, after which the rest of the life is straight
    /// line (unless <paramref name="noSwitch"/> keeps it on declining balance throughout).
    /// </summary>
    private static double VdbCumulative(double cost, double salvage, double life, double upTo, double factor, bool noSwitch)
    {
        var lastPeriod = Math.Ceiling(upTo);
        var total = 0d;
        var remaining = cost - salvage;
        var straightLine = 0d;
        var switched = false;

        for (var period = 1d; period <= lastPeriod; period++)
        {
            var charge = straightLine;
            if (!switched)
            {
                var declining = DdbPeriod(cost, salvage, life, period, factor);
                straightLine = StraightLineCharge(remaining, life, period, noSwitch);
                switched = straightLine > declining;
                charge = switched ? straightLine : declining;
            }

            remaining -= charge;

            // The final period is only charged for the fraction of it that the range covers.
            if (period == lastPeriod)
                charge *= upTo - (lastPeriod - 1);

            total += charge;
        }

        return total;
    }

    /// <summary>
    /// Straight-line charge that spreads the <paramref name="remaining"/> basis over the life left
    /// from <paramref name="period"/>, or zero when <paramref name="noSwitch"/> rules it out.
    /// </summary>
    private static double StraightLineCharge(double remaining, double life, double period, bool noSwitch)
        => noSwitch ? 0 : remaining / (life - (period - 1));

    #endregion

    #region Interest rate conversion and growth

    private static ScalarValue Effect(CalcContext ctx, double nominalRate, double periodsPerYear)
    {
        periodsPerYear = Math.Truncate(periodsPerYear);
        if (nominalRate <= 0 || periodsPerYear < 1)
            return XLError.NumberInvalid;

        return PowerBySquaring(1 + nominalRate / periodsPerYear, periodsPerYear) - 1;
    }

    /// <summary>
    /// <c>x^n</c> for a whole <paramref name="n"/> by right-to-left binary exponentiation. EFFECT rounds
    /// the way this does, not the way <see cref="Math.Pow"/> does: EFFECT(0.05,4) is
    /// 0.050945336914062445 and EFFECT(0.05,1000000) is 0.0512710949426225 in Excel.
    /// </summary>
    private static double PowerBySquaring(double x, double n)
    {
        var result = 1d;
        while (n > 0)
        {
            if (n % 2 == 1)
                result *= x;

            n = Math.Floor(n / 2);
            if (n > 0)
                x *= x;
        }

        return result;
    }

    private static ScalarValue Nominal(CalcContext ctx, double effectiveRate, double periodsPerYear)
    {
        periodsPerYear = Math.Truncate(periodsPerYear);
        if (effectiveRate <= 0 || periodsPerYear < 1)
            return XLError.NumberInvalid;

        return (Math.Pow(effectiveRate + 1, 1 / periodsPerYear) - 1) * periodsPerYear;
    }

    private static ScalarValue Rri(CalcContext ctx, double numberOfPeriods, double presentValue, double futureValue)
    {
        if (numberOfPeriods <= 0 || presentValue <= 0 || futureValue < 0)
            return XLError.NumberInvalid;

        return Math.Pow(futureValue / presentValue, 1 / numberOfPeriods) - 1;
    }

    private static ScalarValue PDuration(CalcContext ctx, double rate, double presentValue, double futureValue)
    {
        if (rate <= 0 || presentValue <= 0 || futureValue <= 0)
            return XLError.NumberInvalid;

        return (Math.Log(futureValue) - Math.Log(presentValue)) / Math.Log(1 + rate);
    }

    private static ScalarValue DollarDe(CalcContext ctx, double fractionalDollar, double fraction)
    {
        fraction = Math.Truncate(fraction);
        if (fraction < 0)
            return XLError.NumberInvalid;
        if (fraction == 0)
            return XLError.DivisionByZero;

        var integerPart = Math.Truncate(fractionalDollar);
        var fractionPart = fractionalDollar - integerPart;

        // The digits after the point are read as a numerator written in `fraction`ths, so they are
        // shifted left by as many decimal places as the denominator occupies.
        var scale = Math.Pow(10, Math.Ceiling(Math.Log10(fraction)));
        return integerPart + fractionPart * scale / fraction;
    }

    private static ScalarValue DollarFr(CalcContext ctx, double decimalDollar, double fraction)
    {
        fraction = Math.Truncate(fraction);
        if (fraction < 0)
            return XLError.NumberInvalid;
        if (fraction == 0)
            return XLError.DivisionByZero;

        var integerPart = Math.Truncate(decimalDollar);
        var fractionPart = decimalDollar - integerPart;

        var scale = Math.Pow(10, Math.Ceiling(Math.Log10(fraction)));
        return integerPart + fractionPart * fraction / scale;
    }

    #endregion

    #region Loan schedules

    private static ScalarValue IsPmt(CalcContext ctx, double rate, double period, double numberOfPayments, double presentValue)
    {
        if (numberOfPayments == 0)
            return XLError.DivisionByZero;

        // The outstanding principal falls linearly to zero, so period `per` still owes
        // (1 - per/nper) of it.
        return presentValue * rate * (period / numberOfPayments - 1);
    }

    private static ScalarValue CumIpmt(CalcContext ctx, double rate, double numberOfPayments, double presentValue, double startPeriod, double endPeriod, double type)
        => CumulativePayment(rate, numberOfPayments, presentValue, startPeriod, endPeriod, type, principal: false);

    private static ScalarValue CumPrinc(CalcContext ctx, double rate, double numberOfPayments, double presentValue, double startPeriod, double endPeriod, double type)
        => CumulativePayment(rate, numberOfPayments, presentValue, startPeriod, endPeriod, type, principal: true);

    private static ScalarValue CumulativePayment(double rate, double numberOfPayments, double presentValue, double startPeriod, double endPeriod, double type, bool principal)
    {
        if (rate <= 0 || numberOfPayments <= 0 || presentValue <= 0)
            return XLError.NumberInvalid;
        if (startPeriod < 1 || endPeriod < 1 || startPeriod > endPeriod)
            return XLError.NumberInvalid;
        if (type != 0 && type != 1)
            return XLError.NumberInvalid;
        if (endPeriod > numberOfPayments)
            return XLError.NumberInvalid;

        var pmt = PmtInternal(rate, numberOfPayments, presentValue, 0, type);

        var total = 0d;
        for (var period = Math.Ceiling(startPeriod); period <= Math.Truncate(endPeriod); period++)
        {
            var interest = InterestOfPeriod(rate, period, presentValue, type, pmt);
            total += principal ? pmt - interest : interest;
        }

        return total;
    }

    /// <summary>Interest portion of a single annuity payment; the IPMT calculation without the argument checks.
    /// Shared by IPMT, PPMT, CUMIPMT and CUMPRINC so they agree on every period.</summary>
    private static double InterestOfPeriod(double rate, double period, double presentValue, double type, double pmt)
    {
        // Payment one of an annuity-due carries no interest — the payment is made before any accrues.
        if (type != 0 && period == 1)
            return 0;

        var interest = FvInternal(rate, period - 1, pmt, presentValue, type) * rate;
        return type != 0 ? interest / (1 + rate) : interest;
    }

    #endregion

    #region Securities

    /// <summary>
    /// Bond-equivalent yield of a Treasury bill: <c>(365 × discount) / (360 − discount × DSM)</c>,
    /// where DSM is the actual number of days between settlement and maturity.
    /// </summary>
    private static ScalarValue TBillEq(CalcContext ctx, double settlement, double maturity, double discount)
    {
        if (!TryTBillDays(settlement, maturity, out var days))
            return XLError.NumberInvalid;
        if (discount <= 0)
            return XLError.NumberInvalid;

        var denominator = 360 - discount * days;
        if (denominator <= 0)
            return XLError.NumberInvalid;

        return 365 * discount / denominator;
    }

    /// <summary>Price per $100 face value of a Treasury bill: <c>100 × (1 − discount × DSM/360)</c>.</summary>
    private static ScalarValue TBillPrice(CalcContext ctx, double settlement, double maturity, double discount)
    {
        if (!TryTBillDays(settlement, maturity, out var days))
            return XLError.NumberInvalid;
        if (discount <= 0)
            return XLError.NumberInvalid;

        return 100 * (1 - discount * days / 360);
    }

    /// <summary>Yield of a Treasury bill: <c>(100 − price)/price × 360/DSM</c>.</summary>
    private static ScalarValue TBillYield(CalcContext ctx, double settlement, double maturity, double price)
    {
        if (!TryTBillDays(settlement, maturity, out var days))
            return XLError.NumberInvalid;
        if (price <= 0)
            return XLError.NumberInvalid;

        return (100 - price) / price * (360 / days);
    }

    /// <summary>
    /// Actual days between settlement and maturity. Treasury bills mature within a year, so a
    /// longer span — or a maturity that does not follow settlement — is rejected.
    /// </summary>
    private static bool TryTBillDays(double settlement, double maturity, out double days)
    {
        var settlementDate = Math.Truncate(settlement);
        var maturityDate = Math.Truncate(maturity);
        days = maturityDate - settlementDate;
        return settlement >= 0 && days > 0 && days <= 365;
    }

    /// <summary>Discount rate of a security: <c>(redemption − price) / redemption / yearFraction</c>.</summary>
    private static ScalarValue Disc(CalcContext ctx, double settlement, double maturity, double price, double redemption, double basis)
    {
        if (price <= 0 || redemption <= 0)
            return XLError.NumberInvalid;
        if (!TrySecurityYearFraction(ctx, settlement, maturity, basis, out var yearFraction, out var error))
            return error;

        return (redemption - price) / redemption / yearFraction;
    }

    /// <summary>Interest rate of a fully invested security: <c>(redemption − investment) / investment / yearFraction</c>.</summary>
    private static ScalarValue IntRate(CalcContext ctx, double settlement, double maturity, double investment, double redemption, double basis)
    {
        if (investment <= 0 || redemption <= 0)
            return XLError.NumberInvalid;
        if (!TrySecurityYearFraction(ctx, settlement, maturity, basis, out var yearFraction, out var error))
            return error;

        return (redemption - investment) / investment / yearFraction;
    }

    /// <summary>Amount received at maturity: <c>investment / (1 − discount × yearFraction)</c>.</summary>
    private static ScalarValue Received(CalcContext ctx, double settlement, double maturity, double investment, double discount, double basis)
    {
        if (investment <= 0 || discount <= 0)
            return XLError.NumberInvalid;
        if (!TrySecurityYearFraction(ctx, settlement, maturity, basis, out var yearFraction, out var error))
            return error;

        var divisor = 1 - discount * yearFraction;
        if (divisor <= 0)
            return XLError.NumberInvalid;

        return investment / divisor;
    }

    /// <summary>
    /// Days between settlement and maturity over days in a year, for one of Excel's five day-count
    /// bases. Shares YEARFRAC's implementation, so basis 1 uses YEARFRAC's average-year-length rule.
    /// </summary>
    private static bool TrySecurityYearFraction(CalcContext ctx, double settlement, double maturity, double basis, out double yearFraction, out XLError error)
    {
        yearFraction = 0;
        error = XLError.NumberInvalid;

        if (Math.Truncate(settlement) >= Math.Truncate(maturity))
            return false;

        var fraction = Functions.DateAndTime.YearFrac(ctx, settlement, maturity, basis);
        if (fraction.TryPickError(out var yearFracError))
        {
            error = yearFracError;
            return false;
        }

        if (!fraction.TryPickNumber(out yearFraction) || yearFraction <= 0)
            return false;

        error = default;
        return true;
    }

    #endregion

    #region Irregular and modified cash flows

    private static AnyValue FvSchedule(CalcContext ctx, Span<AnyValue> args)
    {
        // FVSCHEDULE(principal, schedule) — compounds the principal by each rate in turn.
        if (!args[0].TryReduceToNumber(ctx, out var principal, out var principalError))
            return principalError;

        var value = principal;
        foreach (var scalar in EnumerateScalars(ctx, args[1]))
        {
            if (scalar.IsError)
                return scalar.GetError();
            if (scalar.IsBlank)
                continue;
            if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var rate, out var rateError))
                return rateError;

            value *= 1 + rate;
        }

        return value;
    }

    private static AnyValue Mirr(CalcContext ctx, Span<AnyValue> args)
    {
        // MIRR(values, finance_rate, reinvest_rate) — negative flows are financed at finance_rate,
        // positive flows reinvested at reinvest_rate.
        if (!TryCollectNumbers(ctx, args[..1], out var cashflows, out var valuesError))
            return valuesError;
        if (!args[1].TryReduceToNumber(ctx, out var financeRate, out var financeError))
            return financeError;
        if (!args[2].TryReduceToNumber(ctx, out var reinvestRate, out var reinvestError))
            return reinvestError;

        var count = cashflows.Count;
        if (count < 2)
            return XLError.DivisionByZero;

        if (!TrySplitDiscountedFlows(cashflows, financeRate, reinvestRate, out var positiveNpv, out var negativeNpv))
            return XLError.DivisionByZero;

        var numerator = -positiveNpv * Math.Pow(1 + reinvestRate, count);
        var denominator = negativeNpv * (1 + financeRate);
        if (denominator == 0.0)
            return XLError.DivisionByZero;

        return Math.Pow(numerator / denominator, 1.0 / (count - 1)) - 1;
    }

    /// <summary>
    /// Discount the positive flows at <paramref name="reinvestRate"/> and the negative ones at
    /// <paramref name="financeRate"/>, each by its period. Returns <c>false</c> unless there is at
    /// least one flow of each sign.
    /// </summary>
    private static bool TrySplitDiscountedFlows(List<double> cashflows, double financeRate, double reinvestRate, out double positiveNpv, out double negativeNpv)
    {
        positiveNpv = 0;
        negativeNpv = 0;
        var hasPositive = false;
        var hasNegative = false;
        for (var t = 0; t < cashflows.Count; t++)
        {
            var flow = cashflows[t];
            if (flow > 0)
            {
                hasPositive = true;
                positiveNpv += flow / Math.Pow(1 + reinvestRate, t + 1);
            }
            else if (flow < 0)
            {
                hasNegative = true;
                negativeNpv += flow / Math.Pow(1 + financeRate, t + 1);
            }
        }

        return hasPositive && hasNegative;
    }

    private static AnyValue XNpv(CalcContext ctx, Span<AnyValue> args)
    {
        // XNPV(rate, values, dates) — each flow is discounted by its own actual/365 offset from the
        // first date, rather than by a period index.
        if (!args[0].TryReduceToNumber(ctx, out var rate, out var rateError))
            return rateError;
        if (rate <= -1)
            return XLError.NumberInvalid;
        if (!TryCollectSchedule(ctx, args[1], args[2], out var schedule, out var scheduleError))
            return scheduleError;

        return XNpvOf(rate, schedule);
    }

    private static AnyValue XIrr(CalcContext ctx, Span<AnyValue> args)
    {
        // XIRR(values, dates, [guess]) — the rate at which XNPV of the schedule is zero.
        if (!TryCollectSchedule(ctx, args[0], args[1], out var schedule, out var scheduleError))
            return scheduleError;

        var guess = 0.1;
        if (args.Length > 2 && !args[2].TryReduceToNumber(ctx, out guess, out var guessError))
            return guessError;
        if (guess <= -1)
            return XLError.NumberInvalid;

        if (!HasPositiveAndNegativeAmounts(schedule))
            return XLError.NumberInvalid;

        if (TryXIrrByNewton(schedule, guess, out var rate))
            return rate;

        // Newton wandered off; fall back to bisecting a bracket found by scanning outwards.
        return XIrrByBisection(schedule);
    }

    /// <summary>A rate of return only exists when the schedule has both an inflow and an outflow.</summary>
    private static bool HasPositiveAndNegativeAmounts(List<(double Amount, double Days)> schedule)
    {
        var hasPositive = false;
        var hasNegative = false;
        foreach (var (amount, _) in schedule)
        {
            hasPositive |= amount > 0;
            hasNegative |= amount < 0;
        }

        return hasPositive && hasNegative;
    }

    /// <summary>
    /// Newton's method on XNPV from the <paramref name="guess"/>. Fails when the iteration hits a
    /// non-finite value, a flat derivative, or doesn't converge within the iteration limit.
    /// </summary>
    private static bool TryXIrrByNewton(List<(double Amount, double Days)> schedule, double guess, out double result)
    {
        const int maxIterations = 100;
        const double tolerance = 1e-9;
        const double delta = 1e-7;
        var rate = guess;
        for (var iteration = 0; iteration < maxIterations; iteration++)
        {
            var value = XNpvOf(rate, schedule);
            if (!double.IsFinite(value))
                break;

            var derivative = (XNpvOf(rate + delta, schedule) - value) / delta;
            if (derivative == 0.0 || double.IsNaN(derivative))
                break;

            var nextRate = XIrrNewtonStep(rate, value, derivative);
            if (Math.Abs(nextRate - rate) < tolerance)
            {
                result = nextRate;
                return true;
            }

            rate = nextRate;
        }

        result = default;
        return false;
    }

    /// <summary>
    /// One Newton step, pulled back halfway towards -1 when it would reach a rate of -100% or less,
    /// where XNPV is undefined.
    /// </summary>
    private static double XIrrNewtonStep(double rate, double value, double derivative)
    {
        var nextRate = rate - value / derivative;
        return nextRate <= -1 ? (rate - 1) / 2 : nextRate;
    }

    private static AnyValue XIrrByBisection(List<(double Amount, double Days)> schedule)
    {
        var low = -1 + 1e-9;
        var lowValue = XNpvOf(low, schedule);

        var high = 1e-9;
        double highValue;
        do
        {
            high *= 10;
            highValue = XNpvOf(high, schedule);
        }
        while (Math.Sign(highValue) == Math.Sign(lowValue) && high < 1e9);

        if (Math.Sign(highValue) == Math.Sign(lowValue))
            return XLError.NumberInvalid;

        for (var i = 0; i < 200; i++)
        {
            var middle = (low + high) / 2;
            var value = XNpvOf(middle, schedule);
            if (value == 0.0)
                return middle;

            if (Math.Sign(value) == Math.Sign(lowValue))
            {
                low = middle;
                lowValue = value;
            }
            else
            {
                high = middle;
            }
        }

        return (low + high) / 2;
    }

    private static double XNpvOf(double rate, List<(double Amount, double Days)> schedule)
    {
        var total = 0d;
        foreach (var (amount, days) in schedule)
            total += amount / Math.Pow(1 + rate, days / 365.0);

        return total;
    }

    /// <summary>
    /// Read the paired value/date arguments of XIRR and XNPV into amounts and day offsets from the
    /// first date. The two arguments must hold the same number of cells and no date may fall before
    /// the first one.
    /// </summary>
    private static bool TryCollectSchedule(CalcContext ctx, in AnyValue valuesArg, in AnyValue datesArg, out List<(double Amount, double Days)> schedule, out XLError error)
    {
        schedule = null!;

        if (!TryReadAllNumbers(ctx, valuesArg, truncate: false, out var amounts, out error))
            return false;

        if (!TryReadAllNumbers(ctx, datesArg, truncate: true, out var dates, out error))
            return false;

        error = XLError.NumberInvalid;
        if (amounts.Count != dates.Count || amounts.Count < 2)
            return false;

        var start = dates[0];
        var result = new List<(double Amount, double Days)>(amounts.Count);
        for (var i = 0; i < amounts.Count; i++)
        {
            if (dates[i] < start || dates[i] < 0)
                return false;

            result.Add((amounts[i], dates[i] - start));
        }

        error = default;
        schedule = result;
        return true;
    }

    /// <summary>
    /// Read every cell of <paramref name="arg"/> as a number, whole numbers only when
    /// <paramref name="truncate"/> is set. An error, or a value that isn't a number, stops the walk
    /// with its error.
    /// </summary>
    private static bool TryReadAllNumbers(CalcContext ctx, in AnyValue arg, bool truncate, out List<double> numbers, out XLError error)
    {
        numbers = new List<double>();
        foreach (var scalar in EnumerateScalars(ctx, arg))
        {
            if (scalar.IsError)
            {
                error = scalar.GetError();
                return false;
            }

            if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var number, out error))
                return false;

            numbers.Add(truncate ? Math.Truncate(number) : number);
        }

        error = default;
        return true;
    }

    #endregion

    private static bool TryCollectNumbers(CalcContext ctx, ReadOnlySpan<AnyValue> valueArgs, out List<double> numbers, out XLError error)
    {
        error = default;
        var result = new List<double>();
        foreach (var arg in valueArgs)
        {
            foreach (var scalar in EnumerateScalars(ctx, arg))
            {
                if (scalar.IsError)
                {
                    numbers = null!;
                    error = scalar.GetError();
                    return false;
                }

                if (scalar.IsNumber)
                    result.Add(scalar.GetNumber());
            }
        }

        numbers = result;
        return true;
    }

    /// <summary>
    /// Yield the scalar values of an argument in order, whether it's a single scalar, an array, or a
    /// range reference. Mirrors how <see cref="Statistical"/> reads a data set. A reference yields
    /// only its non-blank cells, in row-major order.
    /// </summary>
    private static IEnumerable<ScalarValue> EnumerateScalars(CalcContext ctx, AnyValue value)
    {
        if (value.TryPickScalar(out var scalar, out var collection))
        {
            yield return scalar;
            yield break;
        }

        if (collection.TryPickT0(out var array, out var reference))
        {
            foreach (var item in array)
                yield return item;
        }
        else
        {
            foreach (var item in ctx.GetNonBlankValues(reference))
                yield return item;
        }
    }
}
