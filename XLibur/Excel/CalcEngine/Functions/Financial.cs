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
        ce.RegisterFunction("EFFECT", 2, 2, Adapt(Effect), FunctionFlags.Scalar); // Returns the effective annual interest rate
        ce.RegisterFunction("FV", 3, 5, AdaptLastTwoOptional(Fv, 0, 0), FunctionFlags.Scalar); // Returns the future value of an investment
        ce.RegisterFunction("FVSCHEDULE", 2, 2, FvSchedule, FunctionFlags.Range, AllowRange.Only, 1); // Returns the future value of an initial principal after applying a series of compound interest rates
        ce.RegisterFunction("INTRATE", 4, 5, AdaptLastOptional(IntRate, 0), FunctionFlags.Scalar); // Returns the interest rate for a fully invested security
        ce.RegisterFunction("IPMT", 4, 6, AdaptLastTwoOptional(Ipmt, 0, 0), FunctionFlags.Scalar); // Returns the interest payment for an investment for a given period
        ce.RegisterFunction("IRR", 1, 2, Irr, FunctionFlags.Range, AllowRange.Only, 0); // Returns the internal rate of return for a series of cash flows
        ce.RegisterFunction("ISPMT", 4, 4, Adapt(IsPmt), FunctionFlags.Scalar); // Calculates the interest paid during a specific period of an investment
        ce.RegisterFunction("MIRR", 3, 3, Mirr, FunctionFlags.Range, AllowRange.Only, 0); // Returns the internal rate of return where positive and negative cash flows are financed at different rates
        ce.RegisterFunction("NOMINAL", 2, 2, Adapt(Nominal), FunctionFlags.Scalar); // Returns the annual nominal interest rate
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
        if (numberOfPayments <= 0 || rate <= -1)
            return XLError.NumberInvalid;

        numberOfPayments = Math.Ceiling(numberOfPayments);

        if (period < 1 || period > numberOfPayments)
            return XLError.NumberInvalid;

        double ipmt = FvInternal(rate, period - 1, PmtInternal(rate, numberOfPayments, presentValue, futureValue, type), presentValue, type) * rate;

        if (type != 0.0)
            ipmt /= (1 + rate);

        return ipmt;
    }

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

        const int paymentAtTheEndOfPeriod = 0;
        const int paymentAtTheBeginningOfPeriod = 1;
        var timingOffset = type != 0.0 ? paymentAtTheBeginningOfPeriod : paymentAtTheEndOfPeriod;

        return (-futureValue - presentValue * Math.Pow(1.0 + rate, numberOfPayments)) /
               (1 + rate * timingOffset) / ((Math.Pow(1.0 + rate, numberOfPayments) - 1) / rate);
    }

    private static AnyValue Pv(double rate, double numberOfPayments, double pmt, double futureValue, double type)
    {
        if (rate == 0.0)
            return -(futureValue + pmt * numberOfPayments);

        var pow = Math.Pow(1 + rate, numberOfPayments);
        return -(futureValue + pmt * (1 + rate * type) * (pow - 1) / rate) / pow;
    }

    private static AnyValue Nper(double rate, double pmt, double presentValue, double futureValue, double type)
    {
        if (rate == 0.0)
        {
            if (pmt == 0.0)
                return XLError.NumberInvalid;

            return -(presentValue + futureValue) / pmt;
        }

        var timing = pmt * (1 + rate * type);
        var numerator = timing - futureValue * rate;
        var denominator = presentValue * rate + timing;
        if (denominator == 0.0 || numerator / denominator <= 0.0)
            return XLError.NumberInvalid;

        return Math.Log(numerator / denominator) / Math.Log(1 + rate);
    }

    private static AnyValue Ppmt(double rate, double period, double numberOfPayments, double presentValue, double futureValue, double type)
    {
        // Principal = total payment - interest payment. Ipmt validates period/nper/rate.
        var ipmt = Ipmt(rate, period, numberOfPayments, presentValue, futureValue, type);
        if (!ipmt.TryPickScalar(out var ipmtScalar, out _) || !ipmtScalar.TryPickNumber(out var ipmtValue))
            return ipmt;

        var pmt = PmtInternal(rate, numberOfPayments, presentValue, futureValue, type);
        return pmt - ipmtValue;
    }

    private static AnyValue Npv(CalcContext ctx, Span<AnyValue> args)
    {
        // NPV(rate, value1, [value2], ...). rate is a scalar (marked param 0), values may be ranges.
        if (!TryScalarNumber(ctx, args[0], out var rate, out var rateError))
            return rateError;
        if (rate <= -1)
            return XLError.NumberInvalid;

        double npv = 0;
        var period = 1;
        for (var i = 1; i < args.Length; i++)
        {
            foreach (var scalar in EnumerateScalars(ctx, args[i]))
            {
                if (scalar.IsError)
                    return scalar.GetError();

                // NPV ignores blanks, text and logicals in references; each number is discounted by
                // its sequential position.
                if (!scalar.IsNumber)
                    continue;

                npv += scalar.GetNumber() / Math.Pow(1 + rate, period);
                period++;
            }
        }

        return npv;
    }

    private static AnyValue Irr(CalcContext ctx, Span<AnyValue> args)
    {
        // IRR(values, [guess]). values is the cash-flow range (marked param 0), starting at period 0.
        if (!TryCollectNumbers(ctx, args[..1], out var cashflows, out var valuesError))
            return valuesError;
        if (cashflows.Count < 2)
            return XLError.NumberInvalid;

        var guess = 0.1;
        if (args.Length > 1 && !TryScalarNumber(ctx, args[1], out guess, out var guessError))
            return guessError;

        const int maxIterations = 50;
        const double tolerance = 1e-7;
        var rate = guess;
        for (var iteration = 0; iteration < maxIterations; iteration++)
        {
            double npv = 0, derivative = 0;
            for (var t = 0; t < cashflows.Count; t++)
            {
                var factor = Math.Pow(1 + rate, t);
                npv += cashflows[t] / factor;
                derivative -= t * cashflows[t] / (factor * (1 + rate));
            }

            if (Math.Abs(npv) < tolerance)
                return rate;
            if (derivative == 0.0)
                break;

            var nextRate = rate - npv / derivative;
            if (Math.Abs(nextRate - rate) < tolerance)
                return nextRate;

            rate = nextRate;
        }

        return XLError.NumberInvalid;
    }

#pragma warning disable S3776 // Six scalar arguments to read before one Newton iteration
    private static AnyValue Rate(CalcContext ctx, Span<AnyValue> args)
    {
        // RATE(nper, pmt, pv, [fv], [type], [guess]) - solved iteratively. All arguments are scalars.
        if (!TryScalarNumber(ctx, args[0], out var nper, out var nperError))
            return nperError;
        if (!TryScalarNumber(ctx, args[1], out var pmt, out var pmtError))
            return pmtError;
        if (!TryScalarNumber(ctx, args[2], out var pv, out var pvError))
            return pvError;

        double fv = 0, type = 0, guess = 0.1;
        if (args.Length > 3 && !TryScalarNumber(ctx, args[3], out fv, out var fvError))
            return fvError;
        if (args.Length > 4 && !TryScalarNumber(ctx, args[4], out type, out var typeError))
            return typeError;
        if (args.Length > 5 && !TryScalarNumber(ctx, args[5], out guess, out var guessError))
            return guessError;

        const int maxIterations = 100;
        const double tolerance = 1e-8;
        const double delta = 1e-6;
        var rate = guess;
        for (var iteration = 0; iteration < maxIterations; iteration++)
        {
            var value = TvmEquation(rate, nper, pmt, pv, fv, type);
            if (Math.Abs(value) < tolerance)
                return rate;

            var derivative = (TvmEquation(rate + delta, nper, pmt, pv, fv, type) - value) / delta;
            if (derivative == 0.0)
                break;

            var nextRate = rate - value / derivative;
            if (Math.Abs(nextRate - rate) < tolerance)
                return nextRate;

            rate = nextRate;
        }

        return XLError.NumberInvalid;
    }
#pragma warning restore S3776

    /// <summary>
    /// The time-value-of-money residual: <c>pv·(1+rate)^nper + pmt·(1+rate·type)·((1+rate)^nper−1)/rate + fv</c>,
    /// which is zero at the solved rate. Used by <see cref="Rate"/>.
    /// </summary>
    private static double TvmEquation(double rate, double nper, double pmt, double pv, double fv, double type)
    {
        if (rate == 0.0)
            return pv + pmt * nper + fv;

        var pow = Math.Pow(1 + rate, nper);
        return pv * pow + pmt * (1 + rate * type) * (pow - 1) / rate + fv;
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
        if (life <= 0)
            return XLError.NumberInvalid;
        if (period <= 0 || period > life)
            return XLError.NumberInvalid;

        return (cost - salvage) * (life - period + 1) * 2 / (life * (life + 1));
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

        var firstPeriod = cost * rate * month / 12;
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
    /// fractional <paramref name="period"/> works the same way Excel's does.
    /// </summary>
    private static double DdbPeriod(double cost, double salvage, double life, double period, double factor)
    {
        var rate = factor / life;
        double openingValue;
        if (rate >= 1)
        {
            rate = 1;
            openingValue = period == 1 ? cost : 0;
        }
        else
        {
            openingValue = cost * Math.Pow(1 - rate, period - 1);
        }

        var closingValue = cost * Math.Pow(1 - rate, period);

        // Depreciation stops once the book value would drop below the salvage value.
        var depreciation = closingValue < salvage ? openingValue - salvage : openingValue - closingValue;
        return Math.Max(depreciation, 0);
    }

    /// <summary>
    /// Depreciation charged between <paramref name="startPeriod"/> and <paramref name="endPeriod"/>,
    /// as the difference of two cumulative runs from period zero. Replaying from the start is what
    /// makes the straight-line cross-over land in the same place regardless of which slice is asked
    /// for, so consecutive slices add up to the whole and a full-life run charges exactly
    /// <c>cost - salvage</c>.
    /// </summary>
    private static ScalarValue Vdb(CalcContext ctx, double cost, double salvage, double life, double startPeriod, double endPeriod, double factor, bool noSwitch)
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
            double charge;
            if (switched)
            {
                charge = straightLine;
            }
            else
            {
                var declining = DdbPeriod(cost, salvage, life, period, factor);
                straightLine = noSwitch ? 0 : remaining / (life - (period - 1));
                if (straightLine > declining)
                {
                    charge = straightLine;
                    switched = true;
                }
                else
                {
                    charge = declining;
                }
            }

            remaining -= charge;

            // The final period is only charged for the fraction of it that the range covers.
            if (period == lastPeriod)
                charge *= upTo - (lastPeriod - 1);

            total += charge;
        }

        return total;
    }

    #endregion

    #region Interest rate conversion and growth

    private static ScalarValue Effect(CalcContext ctx, double nominalRate, double periodsPerYear)
    {
        periodsPerYear = Math.Truncate(periodsPerYear);
        if (nominalRate <= 0 || periodsPerYear < 1)
            return XLError.NumberInvalid;

        return Math.Pow(1 + nominalRate / periodsPerYear, periodsPerYear) - 1;
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

    /// <summary>Interest portion of a single annuity payment; the IPMT calculation without the argument checks.</summary>
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
        if (!TryScalarNumber(ctx, args[0], out var principal, out var principalError))
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
        if (!TryScalarNumber(ctx, args[1], out var financeRate, out var financeError))
            return financeError;
        if (!TryScalarNumber(ctx, args[2], out var reinvestRate, out var reinvestError))
            return reinvestError;

        var count = cashflows.Count;
        if (count < 2)
            return XLError.DivisionByZero;

        double positiveNpv = 0, negativeNpv = 0;
        var hasPositive = false;
        var hasNegative = false;
        for (var t = 0; t < count; t++)
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

        if (!hasPositive || !hasNegative)
            return XLError.DivisionByZero;

        var numerator = -positiveNpv * Math.Pow(1 + reinvestRate, count);
        var denominator = negativeNpv * (1 + financeRate);
        if (denominator == 0.0)
            return XLError.DivisionByZero;

        return Math.Pow(numerator / denominator, 1.0 / (count - 1)) - 1;
    }

    private static AnyValue XNpv(CalcContext ctx, Span<AnyValue> args)
    {
        // XNPV(rate, values, dates) — each flow is discounted by its own actual/365 offset from the
        // first date, rather than by a period index.
        if (!TryScalarNumber(ctx, args[0], out var rate, out var rateError))
            return rateError;
        if (rate <= -1)
            return XLError.NumberInvalid;
        if (!TryCollectSchedule(ctx, args[1], args[2], out var schedule, out var scheduleError))
            return scheduleError;

        return XNpvOf(rate, schedule);
    }

#pragma warning disable S3776 // Newton with a documented bisection fallback; the convergence tests are the algorithm
    private static AnyValue XIrr(CalcContext ctx, Span<AnyValue> args)
    {
        // XIRR(values, dates, [guess]) — the rate at which XNPV of the schedule is zero.
        if (!TryCollectSchedule(ctx, args[0], args[1], out var schedule, out var scheduleError))
            return scheduleError;

        var guess = 0.1;
        if (args.Length > 2 && !TryScalarNumber(ctx, args[2], out guess, out var guessError))
            return guessError;
        if (guess <= -1)
            return XLError.NumberInvalid;

        var hasPositive = false;
        var hasNegative = false;
        foreach (var (amount, _) in schedule)
        {
            hasPositive |= amount > 0;
            hasNegative |= amount < 0;
        }

        if (!hasPositive || !hasNegative)
            return XLError.NumberInvalid;

        const int maxIterations = 100;
        const double tolerance = 1e-9;
        const double delta = 1e-7;
        var rate = guess;
        for (var iteration = 0; iteration < maxIterations; iteration++)
        {
            var value = XNpvOf(rate, schedule);
            if (double.IsNaN(value) || double.IsInfinity(value))
                break;
            if (Math.Abs(value) < tolerance)
                return rate;

            var derivative = (XNpvOf(rate + delta, schedule) - value) / delta;
            if (derivative == 0.0 || double.IsNaN(derivative))
                break;

            var nextRate = rate - value / derivative;
            if (nextRate <= -1)
                nextRate = (rate - 1) / 2;

            if (Math.Abs(nextRate - rate) < tolerance)
                return nextRate;

            rate = nextRate;
        }

        // Newton wandered off; fall back to bisecting a bracket found by scanning outwards.
        return XIrrByBisection(schedule);
    }
#pragma warning restore S3776

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
            if (Math.Abs(value) < 1e-9)
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
#pragma warning disable S3776 // Two symmetric collection walks with error propagation, then one pairing check
    private static bool TryCollectSchedule(CalcContext ctx, in AnyValue valuesArg, in AnyValue datesArg, out List<(double Amount, double Days)> schedule, out XLError error)
    {
        schedule = null!;
        error = default;

        var amounts = new List<double>();
        foreach (var scalar in EnumerateScalars(ctx, valuesArg))
        {
            if (scalar.IsError)
            {
                error = scalar.GetError();
                return false;
            }

            if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var amount, out var amountError))
            {
                error = amountError;
                return false;
            }

            amounts.Add(amount);
        }

        var dates = new List<double>();
        foreach (var scalar in EnumerateScalars(ctx, datesArg))
        {
            if (scalar.IsError)
            {
                error = scalar.GetError();
                return false;
            }

            if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var date, out var dateError))
            {
                error = dateError;
                return false;
            }

            dates.Add(Math.Truncate(date));
        }

        if (amounts.Count != dates.Count || amounts.Count < 2)
        {
            error = XLError.NumberInvalid;
            return false;
        }

        var start = dates[0];
        var result = new List<(double Amount, double Days)>(amounts.Count);
        for (var i = 0; i < amounts.Count; i++)
        {
            if (dates[i] < start || dates[i] < 0)
            {
                error = XLError.NumberInvalid;
                return false;
            }

            result.Add((amounts[i], dates[i] - start));
        }

        schedule = result;
        return true;
    }
#pragma warning restore S3776

    #endregion

    /// <summary>
    /// Read an argument that must be a single number. The functions in this file take ranges for
    /// some parameters and scalars for others, so a scalar parameter still receives whatever the
    /// formula wrote there: a reference to one cell is unwrapped, a larger range goes through
    /// implicit intersection — the same reduction the signature adapters apply.
    /// </summary>
    private static bool TryScalarNumber(CalcContext ctx, in AnyValue value, out double number, out XLError error)
    {
        error = default;
        number = 0;

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
    /// range reference. Mirrors how <see cref="Statistical"/> reads a data set.
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
            foreach (var item in reference.GetCellsValues(ctx))
                yield return item;
        }
    }
}
