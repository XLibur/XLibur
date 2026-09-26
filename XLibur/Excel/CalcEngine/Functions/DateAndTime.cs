using System;
using System.Collections.Generic;
using System.Diagnostics;
using XLibur.Extensions;
using static XLibur.Excel.CalcEngine.Functions.SignatureAdapter;

namespace XLibur.Excel.CalcEngine.Functions;

// S4136 wants every overload group adjacent. Members here are ordered by the Excel function they implement,
// which is the order a reader follows; regrouping by name would break it.
#pragma warning disable S4136
internal static class DateAndTime
{
    /// <summary>
    /// Serial date of 9999-12-31. Date is generally considered invalid, if above that or below 0.
    /// </summary>
    private const int Year10K = 2958465;

    public static void Register(FunctionRegistry ce)
    {
        var dateValue = DateValue;
        var timeValue = TimeValue;
        ce.RegisterFunction("DATE", 3, 3, Adapt(Date), FunctionFlags.Scalar); // Returns the serial number of a particular date
        ce.RegisterFunction("DATEDIF", 3, 3, Adapt(DateDif), FunctionFlags.Scalar); // Calculates the number of days, months, or years between two dates
        ce.RegisterFunction("DATEVALUE", 1, 1, Adapt(dateValue), FunctionFlags.Scalar); // Converts a date in the form of text to a serial number
        ce.RegisterFunction("DAY", 1, 1, Adapt(GetDay), FunctionFlags.Scalar); // Converts a serial number to a day of the month
        ce.RegisterFunction("DAYS", 2, 2, Adapt(Days), FunctionFlags.Scalar | FunctionFlags.Future); // Returns the number of days between two dates.
        ce.RegisterFunction("DAYS360", 2, 3, AdaptLastOptional(Days360, false), FunctionFlags.Scalar); // Calculates the number of days between two dates based on a 360-day year
        ce.RegisterFunction("EDATE", 2, 2, Adapt(EDate), FunctionFlags.Scalar); // Returns the serial number of the date that is the indicated number of months before or after the start date
        ce.RegisterFunction("EOMONTH", 2, 2, Adapt(Eomonth), FunctionFlags.Scalar); // Returns the serial number of the last day of the month before or after a specified number of months
        ce.RegisterFunction("HOUR", 1, 1, Adapt(Hour), FunctionFlags.Scalar); // Converts a serial number to an hour
        ce.RegisterFunction("ISOWEEKNUM", 1, 1, Adapt(IsoWeekNum), FunctionFlags.Scalar | FunctionFlags.Future); // Returns number of the ISO week number of the year for a given date.
        ce.RegisterFunction("MINUTE", 1, 1, Adapt(Minute), FunctionFlags.Scalar); // Converts a serial number to a minute
        ce.RegisterFunction("MONTH", 1, 1, Adapt(GetMonth), FunctionFlags.Scalar); // Converts a serial number to a month
        ce.RegisterFunction("NETWORKDAYS", 2, 3, AdaptLastOptional(NetWorkDays), FunctionFlags.Range, AllowRange.Only, 2); // Returns the number of whole workdays between two dates
        ce.RegisterFunction("NETWORKDAYS.INTL", 2, 4, NetWorkDaysIntl, FunctionFlags.Range | FunctionFlags.Future, AllowRange.Only, 3); // Whole workdays between two dates, with a configurable weekend
        ce.RegisterFunction("NOW", 0, 0, Adapt(Now), FunctionFlags.Scalar | FunctionFlags.Volatile); // Returns the serial number of the current date and time
        ce.RegisterFunction("SECOND", 1, 1, Adapt(Second), FunctionFlags.Scalar); // Converts a serial number to a second
        ce.RegisterFunction("TIME", 3, 3, Adapt(Time), FunctionFlags.Scalar); // Returns the serial number of a particular time
        ce.RegisterFunction("TIMEVALUE", 1, 1, Adapt(timeValue), FunctionFlags.Scalar); // Converts a time in the form of text to a serial number
        ce.RegisterFunction("TODAY", 0, 0, Adapt(Today), FunctionFlags.Scalar | FunctionFlags.Volatile); // Returns the serial number of today's date
        ce.RegisterFunction("WEEKDAY", 1, 2, AdaptLastOptional(Weekday), FunctionFlags.Scalar); // Converts a serial number to a day of the week
        ce.RegisterFunction("WEEKNUM", 1, 2, AdaptLastOptional(WeekNum, 1), FunctionFlags.Scalar); // Converts a serial number to a number representing where the week falls numerically with a year
        ce.RegisterFunction("WORKDAY", 2, 3, AdaptLastOptional(Workday), FunctionFlags.Range, AllowRange.Only, 2); // Returns the serial number of the date before or after a specified number of workdays
        ce.RegisterFunction("WORKDAY.INTL", 2, 4, WorkdayIntl, FunctionFlags.Range | FunctionFlags.Future, AllowRange.Only, 3); // The date a number of workdays away, with a configurable weekend
        ce.RegisterFunction("YEAR", 1, 1, Adapt(GetYear), FunctionFlags.Scalar); // Converts a serial number to a year
        ce.RegisterFunction("YEARFRAC", 2, 3, AdaptLastOptional(YearFrac, 0), FunctionFlags.Scalar); // Returns the year fraction representing the number of whole days between start_date and end_date
    }

    private static ScalarValue Date(CalcContext ctx, double year, double month, double day)
    {
        // Unlike most functions, values are floored - not truncated.
        year = Math.Floor(year);
        month = Math.Floor(month);
        day = Math.Floor(day);

        if (month is < -short.MaxValue or >= short.MaxValue)
            return XLError.NumberInvalid;

        // Excel behaves out of spec. Spec says 0-99 are interpreted as year + 1900,
        // but reality is that anything below 1900 is interpreted as year + 1900.
        // That seems to be true for both 1900 and 1904 date systems.
        if (year < 1900)
            year += 1900;

        // Excel buggy implementation :) Should probably return error,
        // but silently changes the result instead.
        day = Math.Min(day, short.MaxValue);
        if (day < short.MinValue)
            day = short.MaxValue;

        if (year > 10000)
            year = 10000;

        // Excel allows months and days outside the normal range, and adjusts the date
        // accordingly.
        var yearAdjustment = Math.Floor((month - 1d) / 12.0);
        year += yearAdjustment;
        month -= yearAdjustment * 12;

        // Year 1 is earliest allowable in both date system. Also avoid the double
        // to int conversion problems when double is too small.
        // DateTime only accepts years 1-9999, so year 10000+ must return an error.
        if (year < 1 || year > 9999)
            return XLError.NumberInvalid;

        var startOfMonth = new DateTime((int)year, (int)month, 1, 0, 0, 0, DateTimeKind.Unspecified).ToSerialDateTime();
        var serialDate = startOfMonth + day - 1;
        if (serialDate < 0 || serialDate >= ctx.DateSystemUpperLimit)
            return XLError.NumberInvalid;

        return serialDate;
    }

    private static ScalarValue DateDif(CalcContext ctx, double startDateTime, double endDateTime, string unit)
    {
        if (!TryGetDate(ctx, startDateTime, out var startSerialDate))
            return XLError.NumberInvalid;

        if (!TryGetDate(ctx, endDateTime, out var endSerialDate))
            return XLError.NumberInvalid;

        if (startSerialDate > endSerialDate)
            return XLError.NumberInvalid;

        var startDate = DateParts.From(ctx, startSerialDate);
        var endDate = DateParts.From(ctx, endSerialDate);

        return unit.ToUpperInvariant() switch
        {
            "Y" => DateDifYears(startDate, endDate),
            "M" => DateDifMonths(startDate, endDate),
            "D" => endSerialDate - startSerialDate,
            "MD" => DateDifDaysIgnoringMonths(startDate, endDate, endSerialDate),
            "YM" => DateDifMonthsIgnoringYears(startDate, endDate),
            "YD" => DateDifDaysIgnoringYears(startDate, endDate, startSerialDate),
            _ => XLError.NumberInvalid
        };
    }

    private static int DateDifYears(DateParts startDate, DateParts endDate)
    {
        // Calculate number of complete years the end date is from the start date
        var isLastYearComplete = endDate.Month > startDate.Month ||
                                 (endDate.Month == startDate.Month && endDate.Day >= startDate.Day);
        return endDate.Year - startDate.Year - (isLastYearComplete ? 0 : 1);
    }

    private static int DateDifMonths(DateParts startDate, DateParts endDate)
    {
        // Calculate number of complete months the end date is from start date
        var isLastMonthComplete = endDate.Day >= startDate.Day;
        return (endDate.Year - startDate.Year) * 12 + endDate.Month - startDate.Month - (isLastMonthComplete ? 0 : 1);
    }

    private static int DateDifDaysIgnoringMonths(DateParts startDate, DateParts endDate, int endSerialDate)
    {
        // The difference between the days in startDate and endDate, ignore year and month
        // of startDate, only days are used
        if (endDate.Day >= startDate.Day)
            return endDate.Day - startDate.Day;

        var adjacentStartDate = startDate with
        {
            Month = (endDate.Month - 2 + 12) % 12 + 1,
            Year = endDate.Month > 1 ? endDate.Year : endDate.Year - 1
        };
        return endSerialDate - adjacentStartDate.SerialDate;
    }

    private static int DateDifMonthsIgnoringYears(DateParts startDate, DateParts endDate)
    {
        // The difference between the months in start-date and end-date. Add 12 and then
        // modulo, so result is always positive
        var isLastMonthComplete = endDate.Day >= startDate.Day;
        return (endDate.Month + 12 - startDate.Month - (isLastMonthComplete ? 0 : 1)) % 12;
    }

    private static int DateDifDaysIgnoringYears(DateParts startDate, DateParts endDate, int startSerialDate)
    {
        var endFollowsStart = endDate.Month > startDate.Month || (endDate.Month == startDate.Month && endDate.Day >= startDate.Day);
        var newEndYear = startDate.Year + (endFollowsStart ? 0 : 1);
        var newEndDate = endDate with { Year = newEndYear };
        var daysDiff = newEndDate.SerialDate - startSerialDate;

        // If start date is in 1900 jan/feb, there are sometimes errors. I couldn't decipher actual logic,
        // only condition when it happens. Based on the Excel vs XLibur comparisons, it seems to work.
        if (startSerialDate <= 60 && endDate is { Year: > 1900, Month: 3 } && endDate.Day < startDate.Day)
            daysDiff--;

        return daysDiff;
    }

    private static ScalarValue DateValue(CalcContext ctx, ScalarValue value)
    {
        if (!value.TryPickText(out var text, out var error))
            return error;

        if (!ScalarValue.ToSerialDateTime(text!, ctx.Culture, out var serialDateTime))
            return XLError.IncompatibleValue;

        return Math.Truncate(serialDateTime);
    }

    private static ScalarValue GetDay(CalcContext ctx, double serialDateTime)
    {
        if (!TryGetDate(ctx, serialDateTime, out var serialDate))
            return XLError.NumberInvalid;

        return DateParts.From(ctx, serialDate).Day;
    }

    private static ScalarValue Days(CalcContext ctx, double endSerialDate, double startSerialDate)
    {
        if (!TryGetDate(ctx, startSerialDate, out var startDate))
            return XLError.NumberInvalid;

        if (!TryGetDate(ctx, endSerialDate, out var endDate))
            return XLError.NumberInvalid;

        return endDate - startDate;
    }

    private static ScalarValue Days360(CalcContext ctx, double startDateTime, double endDateTime, bool isEuropean)
    {
        if (!TryGetDate(ctx, startDateTime, out var startDate))
            return XLError.NumberInvalid;

        if (!TryGetDate(ctx, endDateTime, out var endDate))
            return XLError.NumberInvalid;

        return Days360(ctx, startDate, endDate, isEuropean);
    }

    private static int Days360(CalcContext ctx, int startSerialDate, int endSerialDate, bool isEuropean)
    {
        var startDate = DateParts.From(ctx, startSerialDate);
        var (startYear, startMonth, startDay) = startDate;
        var (endYear, endMonth, endDay) = DateParts.From(ctx, endSerialDate);

        if (isEuropean)
            AdjustEuropeanDays360(ref startDay, ref endDay);
        else
            AdjustUsDays360(startDate, ref startDay, ref endDay);

        return 360 * (endYear - startYear) + 30 * (endMonth - startMonth) + (endDay - startDay);
    }

    private static void AdjustEuropeanDays360(ref int startDay, ref int endDay)
    {
        if (startDay == 31)
            startDay = 30;

        if (endDay == 31)
            endDay = 30;
    }

    private static void AdjustUsDays360(DateParts startDate, ref int startDay, ref int endDay)
    {
        // There are several descriptions of the US algorithm: spec, wikipedia, function help,
        // ODF. Out of these, only ODF is correct (rest is incomplete/has different results).
        if (startDate.IsLastDayOfMonth())
            startDay = 30;

        if (endDay == 31 && startDay == 30)
            endDay = 30;
    }

    private static ScalarValue EDate(CalcContext ctx, double startSerialDate, double monthOffset)
    {
        if (!TryGetDate(ctx, startSerialDate, out var startSerial))
            return XLError.NumberInvalid;

        if (!TryGetMonthsOffset(monthOffset, out var months))
            return XLError.NumberInvalid;

        var startDate = DateParts.From(ctx, startSerial);
        var endDate = startDate.AddMonths(months);
        return endDate.ToSerialDate()
            .Match<ScalarValue>(d => d, e => e);
    }

    private static ScalarValue Eomonth(CalcContext ctx, double startSerialDate, double monthOffset)
    {
        if (!TryGetDate(ctx, startSerialDate, out var startSerial))
            return XLError.NumberInvalid;

        if (!TryGetMonthsOffset(monthOffset, out var months))
            return XLError.NumberInvalid;

        var startDate = DateParts.From(ctx, startSerial);
        var endDate = startDate.AddMonths(months);
        var endOfMonth = endDate.EndOfMonth();
        return endOfMonth.ToSerialDate()
            .Match<ScalarValue>(d => d, e => e);
    }

    private static ScalarValue Hour(CalcContext ctx, double serialTime)
    {
        return GetTimeComponent(ctx, serialTime, static d => d.Hour);
    }

    private static ScalarValue IsoWeekNum(CalcContext ctx, double serialDateTime)
    {
        // Uses ISO week algorithm from Wikipedia
        if (!TryGetDate(ctx, serialDateTime, out var serialDate))
            return XLError.NumberInvalid;

        var date = DateParts.From(ctx, serialDate);

        // Normalized to Monday = 1, Sunday = 7
        var dayOfWeek = ((int)date.DayOfWeek + 6) % 7 + 1;
        var week = (10 + date.DayOfYear - dayOfWeek) / 7;

        if (week < 1)
            return Weeks(date.Year - 1);

        if (week > Weeks(date.Year))
            return 1;

        return week;

        // Returns day of a week on the last day of a year - Dec 31
        static int DayOfWeekDec31(int year)
        {
            return (year + year / 4 - year / 100 + year / 400) % 7;
        }

        static int Weeks(int year)
        {
            // Year ends on Thursday
            if (DayOfWeekDec31(year) == 4)
                return 53;

            // Previous year ends on Wednesday
            if (DayOfWeekDec31(year - 1) == 3)
                return 53;

            return 52;
        }
    }

    private static ScalarValue Minute(CalcContext ctx, double serialTime)
    {
        return GetTimeComponent(ctx, serialTime, static d => d.Minute);
    }

    private static ScalarValue GetMonth(CalcContext ctx, double serialDateTime)
    {
        if (!TryGetDate(ctx, serialDateTime, out var serialDate))
            return XLError.NumberInvalid;

        return DateParts.From(ctx, serialDate).Month;
    }

    /// <summary>
    /// NETWORKDAYS is NETWORKDAYS.INTL with weekend code 1: the same count over a Saturday and
    /// Sunday weekend.
    /// </summary>
    private static ScalarValue NetWorkDays(CalcContext ctx, ScalarValue startDate, ScalarValue endDate, AnyValue holidays)
    {
        if (!TryGetDate(ctx, startDate, out var startSerialDate, out var startDateError))
            return startDateError;

        if (!TryGetDate(ctx, endDate, out var endSerialDate, out var endDateError))
            return endDateError;

        if (!TryGetHolidays(ctx, holidays, SaturdaySundayWeekend, out var holidayDates, out var holidaysError))
            return holidaysError;

        return CountNetWorkdays(startSerialDate, endSerialDate, SaturdaySundayWeekend, holidayDates);
    }

    #region Configurable weekends

    /// <summary>
    /// Saturday and Sunday, weekend code 1: the default of the .INTL functions and the only weekend
    /// of NETWORKDAYS and WORKDAY.
    /// </summary>
    private const int SaturdaySundayWeekend = (1 << 5) | (1 << 6);

    /// <summary>
    /// Every day of the week is a weekend. NETWORKDAYS.INTL accepts it and counts no working days;
    /// WORKDAY.INTL rejects it with <c>#VALUE!</c>, since it would leave no day to land on.
    /// </summary>
    private const int AllDaysWeekend = 0b111_1111;

    /// <summary>
    /// The numbered weekend codes, as bitmasks over the days of the week with bit 0 = Monday.
    /// 1..7 are the two-day weekends, sliding one day later each time from Saturday+Sunday; 11..17
    /// are the single-day weekends, from Sunday to Saturday.
    /// </summary>
    private static bool TryGetWeekendMaskFromCode(int code, out int mask)
    {
        mask = 0;
        if (code is >= 1 and <= 7)
        {
            // Code 1 is Sat+Sun (bits 5 and 6); each later code slides the pair one day on.
            var first = (5 + code - 1) % 7;
            mask = (1 << first) | (1 << ((first + 1) % 7));
            return true;
        }

        if (code is >= 11 and <= 17)
        {
            // Code 11 is Sunday alone (bit 6), 12 is Monday (bit 0), and so on.
            mask = 1 << ((code - 11 + 6) % 7);
            return true;
        }

        return false;
    }

    /// <summary>
    /// Read the <c>weekend</c> argument of the .INTL functions: either one of the numbered codes, or
    /// a seven-character string of 0s and 1s running Monday to Sunday where 1 marks a weekend day.
    /// </summary>
    /// <remarks>
    /// An error passed as the weekend is not propagated: Excel replaces it with
    /// <paramref name="errorValueBecomes"/>, which is <c>#NUM!</c> for WORKDAY.INTL and
    /// <c>#VALUE!</c> for NETWORKDAYS.INTL. A string that is not a mask is <c>#VALUE!</c>, a number
    /// that is not a code is <c>#NUM!</c>, and so is a reference to an empty cell — only an omitted
    /// weekend means Saturday and Sunday.
    /// </remarks>
    private static bool TryGetWeekendMask(CalcContext ctx, in AnyValue value, XLError errorValueBecomes, bool allowAllDays, out int mask, out XLError error)
    {
        mask = 0;
        var scalar = value.ReduceToScalar(ctx);
        if (scalar.IsError)
        {
            error = errorValueBecomes;
            return false;
        }

        if (scalar.IsBlank)
        {
            mask = SaturdaySundayWeekend;
            error = value.IsReference ? XLError.NumberInvalid : default;
            return !value.IsReference;
        }

        if (scalar.TryPickText(out var pattern, out _))
        {
            error = XLError.IncompatibleValue;
            return TryGetWeekendMaskFromPattern(pattern!, out mask) && (allowAllDays || mask != AllDaysWeekend);
        }

        error = XLError.NumberInvalid;
        return scalar.ToNumber(ctx.Culture).TryPickT0(out var number, out _)
               && TryGetWeekendMaskFromCode((int)Math.Truncate(number), out mask);
    }

    /// <summary>
    /// Read a seven-character string of 0s and 1s running Monday to Sunday where 1 marks a weekend day.
    /// </summary>
    private static bool TryGetWeekendMaskFromPattern(string pattern, out int mask)
    {
        mask = 0;
        if (pattern.Length != 7)
            return false;

        for (var day = 0; day < 7; day++)
        {
            switch (pattern[day])
            {
                case '1':
                    mask |= 1 << day;
                    break;
                case '0':
                    break;
                default:
                    return false;
            }
        }

        return true;
    }

    /// <summary>Day of the week of a serial date as a bit index, 0 = Monday … 6 = Sunday.</summary>
    private static int WeekdayBit(int serialDate) => (WeekdayCalc(serialDate) + 5) % 7;

    private static bool IsWeekend(int serialDate, int mask) => (mask & (1 << WeekdayBit(serialDate))) != 0;

    /// <summary>
    /// Collect the holidays argument into distinct serial dates, dropping the ones that already
    /// fall on a weekend — they are not working days to begin with, so they must not be subtracted
    /// twice.
    /// </summary>
    private static bool TryGetHolidays(CalcContext ctx, in AnyValue holidays, int mask, out HashSet<int> dates, out XLError error)
    {
        dates = [];
        error = default;
        foreach (var holidayValue in ctx.GetNonBlankValues(holidays))
        {
            if (!TryGetDate(ctx, holidayValue, out var holidayDate, out error))
                return false;

            if (!IsWeekend(holidayDate, mask))
                dates.Add(holidayDate);
        }

        return true;
    }

    /// <remarks>
    /// Excel reads the weekend before the dates, so <c>=NETWORKDAYS.INTL(#N/A, 1, 99)</c> is
    /// <c>#NUM!</c>. An error given as the weekend becomes <c>#VALUE!</c>, and a weekend of all
    /// seven days is allowed and counts no working days.
    /// </remarks>
    private static AnyValue NetWorkDaysIntl(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryGetWeekendMask(ctx, OptionalArg(args, 2), XLError.IncompatibleValue, allowAllDays: true, out var mask, out var weekendError))
            return weekendError;

        if (!TryGetDate(ctx, args[0].ReduceToScalar(ctx), out var startDate, out var startError))
            return startError;

        if (!TryGetDate(ctx, args[1].ReduceToScalar(ctx), out var endDate, out var endError))
            return endError;

        if (!TryGetHolidays(ctx, OptionalArg(args, 3), mask, out var holidays, out var holidaysError))
            return holidaysError;

        return CountNetWorkdays(startDate, endDate, mask, holidays);
    }

    private static AnyValue OptionalArg(Span<AnyValue> args, int index)
        => args.Length > index ? args[index] : ScalarValue.Blank.ToAnyValue();

    /// <summary>
    /// Working days from <paramref name="startDate"/> to <paramref name="endDate"/>, both included,
    /// less the <paramref name="holidays"/> between them. The shared core of NETWORKDAYS and
    /// NETWORKDAYS.INTL.
    /// </summary>
    private static int CountNetWorkdays(int startDate, int endDate, int mask, HashSet<int> holidays)
    {
        // Counting backwards gives the same magnitude with the opposite sign.
        var reversed = startDate > endDate;
        if (reversed)
            (startDate, endDate) = (endDate, startDate);

        var total = CountWorkdays(startDate, endDate, mask) - CountHolidaysBetween(holidays, startDate, endDate);
        return reversed ? -total : total;
    }

    private static int CountHolidaysBetween(HashSet<int> holidays, int startDate, int endDate)
    {
        var count = 0;
        foreach (var holiday in holidays)
        {
            if (holiday >= startDate && holiday <= endDate)
                count++;
        }

        return count;
    }

    /// <summary>
    /// Working days in an inclusive date range. Whole weeks contribute a fixed number, so only the
    /// days that spill past the last whole week have to be looked at one by one.
    /// </summary>
    private static int CountWorkdays(int startDate, int endDate, int mask)
    {
        var weekendDaysPerWeek = System.Numerics.BitOperations.PopCount((uint)mask);
        var days = endDate - startDate + 1;
        var wholeWeeks = days / 7;

        var total = days - wholeWeeks * weekendDaysPerWeek;
        for (var day = startDate + wholeWeeks * 7; day <= endDate; day++)
        {
            if (IsWeekend(day, mask))
                total--;
        }

        return total;
    }

    private static AnyValue WorkdayIntl(CalcContext ctx, Span<AnyValue> args)
        => Workday(ctx, args[0].ReduceToScalar(ctx), args[1].ReduceToScalar(ctx), OptionalArg(args, 2), OptionalArg(args, 3)).ToAnyValue();

    /// <summary>
    /// The shared body of WORKDAY and WORKDAY.INTL, in the order Excel checks the arguments.
    /// </summary>
    /// <remarks>
    /// <list type="number">
    /// <item>The start date, then the offset. The offset is rounded down, not toward zero, so -0.5
    /// is one working day back.</item>
    /// <item>A zero offset returns the start date, weekend or not, before the weekend or holidays
    /// are looked at — even when they are errors.</item>
    /// <item>Holidays given as one value (a scalar or a single cell) are converted to a number: an
    /// error propagates, and text or a logical is <c>#VALUE!</c>.</item>
    /// <item>The weekend. An error given as the weekend becomes <c>#NUM!</c>; a weekend of all seven
    /// days is <c>#VALUE!</c>.</item>
    /// <item>Every holiday, in order, as a date.</item>
    /// </list>
    /// </remarks>
    private static ScalarValue Workday(CalcContext ctx, ScalarValue startDateScalar, ScalarValue offsetScalar, AnyValue weekend, AnyValue holidays)
    {
        if (!TryGetDate(ctx, startDateScalar, out var startDate, out var startDateError))
            return startDateError;

        if (!offsetScalar.ToNumber(ctx.Culture).TryPickT0(out var offsetNumber, out var offsetError))
            return offsetError;

        var offset = Math.Floor(offsetNumber);
        if (offset == 0)
            return startDate;

        if (!TryCoerceSingleHoliday(ctx, holidays, out var singleHolidayError))
            return singleHolidayError;

        if (!TryGetWeekendMask(ctx, weekend, XLError.NumberInvalid, allowAllDays: false, out var mask, out var weekendError))
            return weekendError;

        if (!TryGetHolidays(ctx, holidays, mask, out var holidayDates, out var holidaysError))
            return holidaysError;

        return StepWorkdays(startDate, offset, mask, holidayDates);
    }

    /// <summary>
    /// WORKDAY.INTL converts a holidays argument that is a single value before it reads the weekend,
    /// so <c>=WORKDAY.INTL(1, 1, 99, "abc")</c> is <c>#VALUE!</c> rather than the weekend's
    /// <c>#NUM!</c>. Whether the number is a valid date is left for later, and a range or an array
    /// is left whole for later too.
    /// </summary>
    private static bool TryCoerceSingleHoliday(CalcContext ctx, in AnyValue holidays, out XLError error)
    {
        error = default;
        if (!holidays.TryPickScalar(out var holiday, out var collection))
        {
            if (collection.TryPickT0(out _, out var reference) || !reference.TryGetSingleCellValue(out holiday, ctx))
                return true;
        }

        if (holiday.IsBlank)
            return true;

        if (holiday.IsLogical)
        {
            error = XLError.IncompatibleValue;
            return false;
        }

        return holiday.ToNumber(ctx.Culture).TryPickT0(out _, out error);
    }

    /// <summary>
    /// The date <paramref name="offset"/> working days from <paramref name="startDate"/>, or
    /// <c>#NUM!</c> when it falls outside the supported date range. The shared core of WORKDAY and
    /// WORKDAY.INTL; <paramref name="offset"/> is a whole, non-zero number.
    /// </summary>
    /// <remarks>
    /// Any seven consecutive days hold the same number of working days, so whole weeks are jumped
    /// at once, taking off the holidays passed over. Only the last week or so is walked day by day,
    /// which keeps a large offset from costing one step per calendar day.
    /// </remarks>
    private static ScalarValue StepWorkdays(int startDate, double offset, int mask, HashSet<int> holidays)
    {
        // Each working day moves the date on by at least one day, so an offset this large cannot
        // land inside the range; checking here also keeps the cast to int safe.
        if (Math.Abs(offset) > Year10K)
            return XLError.NumberInvalid;

        var step = offset > 0 ? 1 : -1;
        var remaining = (int)Math.Abs(offset);
        var workdaysPerWeek = 7 - System.Numerics.BitOperations.PopCount((uint)mask);
        var holidaysAhead = HolidaysInWalkOrder(holidays, startDate, step);
        var nextHoliday = 0;
        var date = startDate;

        // Leave at least one working day for the walk below, so the answer is always past the jump.
        while (remaining > workdaysPerWeek)
        {
            var weeks = (remaining - 1) / workdaysPerWeek;
            var landing = date + step * 7L * weeks;
            if (landing < 0 || landing > Year10K)
                return XLError.NumberInvalid;

            date = (int)landing;
            var holidaysPassed = 0;
            while (nextHoliday < holidaysAhead.Count && (holidaysAhead[nextHoliday] - date) * step <= 0)
            {
                holidaysPassed++;
                nextHoliday++;
            }

            remaining -= weeks * workdaysPerWeek - holidaysPassed;
        }

        while (remaining > 0)
        {
            date += step;
            if (date < 0 || date > Year10K)
                return XLError.NumberInvalid;

            if (!IsWeekend(date, mask) && !holidays.Contains(date))
                remaining--;
        }

        return date;
    }

    /// <summary>
    /// The holidays strictly past <paramref name="startDate"/> in the direction of
    /// <paramref name="step"/>, in the order a walk in that direction meets them.
    /// </summary>
    private static List<int> HolidaysInWalkOrder(HashSet<int> holidays, int startDate, int step)
    {
        var ahead = new List<int>(holidays.Count);
        foreach (var holiday in holidays)
        {
            if ((holiday - startDate) * step > 0)
                ahead.Add(holiday);
        }

        ahead.Sort();
        if (step < 0)
            ahead.Reverse();

        return ahead;
    }

    #endregion

    private static ScalarValue Now()
    {
        return DateTime.Now.ToSerialDateTime();
    }

    private static ScalarValue Second(CalcContext ctx, double serialDate)
    {
        return GetTimeComponent(ctx, serialDate, static d => d.Second);
    }

    private static ScalarValue Time(CalcContext ctx, double hour, double minute, double second)
    {
        if (!TryGetComponent(hour, out var hourFloored))
            return XLError.NumberInvalid;

        if (!TryGetComponent(minute, out var minuteFloored))
            return XLError.NumberInvalid;

        if (!TryGetComponent(second, out var secondFloored))
            return XLError.NumberInvalid;

        var serialDate = new TimeSpan(hourFloored, minuteFloored, secondFloored).ToSerialDateTime();
        return serialDate % 1.0;

        static bool TryGetComponent(double value, out int truncated)
        {
            value = Math.Floor(value);
            if (value is < 0 or > 32767)
            {
                truncated = default;
                return false;
            }

            truncated = checked((int)value);
            return true;
        }
    }

    private static ScalarValue TimeValue(CalcContext ctx, ScalarValue value)
    {
        if (!value.TryPickText(out var text, out var error))
            return error;

        if (!ScalarValue.ToSerialDateTime(text!, ctx.Culture, out var serialDateTime))
            return XLError.IncompatibleValue;

        return serialDateTime % 1.0;
    }

    private static ScalarValue Today()
    {
        return DateTime.Today.ToSerialDateTime();
    }

    private static ScalarValue Weekday(CalcContext ctx, ScalarValue date, ScalarValue flag)
    {
        if (!TryGetDate(ctx, date, out var serialDate, out var dateError, acceptLogical: true))
            return dateError;

        var flagValue = 1d;
        // Caller provided a value for optional parameter
        if (!flag.IsBlank && !flag.ToNumber(ctx.Culture).TryPickT0(out flagValue, out var flagError))
            return flagError;

        var result = Weekday(serialDate, (int)Math.Truncate(flagValue));

        if (!result.TryPickT0(out var weekday, out var weekdayError))
            return weekdayError;

        return weekday;
    }

    private static OneOf<int, XLError> Weekday(int serialDate, int startFlag)
    {
        // There are two offsets:
        // - what is the starting day
        // - how are days numbered (0-6, 1-7 ...)
        int? weekStartOffset = startFlag switch
        {
            1 => 0, // Sun
            2 => 6, // Mon
            3 => 6, // Mon
            11 => 6, // Mon
            12 => 5, // Tue
            13 => 4, // Wed
            14 => 3, // Thu
            15 => 2, // Fri
            16 => 1, // Sat
            17 => 0, // Sunday
            _ => null,
        };
        if (weekStartOffset is null)
            return XLError.NumberInvalid;

        var numberOffset = startFlag == 3 ? 0 : 1;

        // Because we don't go below 1900, there is no need to deal with UTC vs Gregorian calendar.
        // It is affected by 1900 bug, so no accurate weekdays before 1900-02-29. It was Wednesday BTW :)
        var weekday = WeekdayCalc(serialDate, weekStartOffset.Value, numberOffset);
        return weekday;
    }

    /// <summary>
    /// Calculate week day. No checks. The default form is form 3 (week starts at Sun, range 1..7).
    /// </summary>
    private static int WeekdayCalc(int serialDate, int weekStartOffset = 0, int numberOffset = 1)
    {
        return (serialDate + 6 + weekStartOffset) % 7 + numberOffset;
    }

    private static ScalarValue WeekNum(CalcContext ctx, double serialDateTime, double weekStartFlag = 1)
    {
        if (!TryGetDate(ctx, serialDateTime, out var serialDate))
            return XLError.NumberInvalid;

        var flag = (int)weekStartFlag;
        var firstDayOfWeek = flag switch
        {
            1 => DayOfWeek.Sunday,
            2 => DayOfWeek.Monday,
            11 => DayOfWeek.Monday,
            12 => DayOfWeek.Tuesday,
            13 => DayOfWeek.Wednesday,
            14 => DayOfWeek.Thursday,
            15 => DayOfWeek.Friday,
            16 => DayOfWeek.Saturday,
            17 => DayOfWeek.Sunday,
            21 => DayOfWeek.Monday,
            _ => (DayOfWeek)(-1),
        };

        if (firstDayOfWeek < 0)
            return XLError.NumberInvalid;

        // Use existing function
        if (flag == 21)
            return IsoWeekNum(ctx, serialDateTime);

        // When checking all values against Excel, there were two cases when week is 0
        if (serialDate == 0 && firstDayOfWeek == DayOfWeek.Sunday)
            return 0;

        var date = DateParts.From(ctx, serialDate);
        var startOfYearDate = (int)(new DateTime(date.Year, 1, 1, 0, 0, 0, DateTimeKind.Unspecified).ToSerialDateTime());
        var startOfYearDayOfWeek = (DayOfWeek)((startOfYearDate + 6) % 7);
        var startOfWeekAdjust = firstDayOfWeek - startOfYearDayOfWeek;

        // In 1-17 flags, the start of a week must be at Jan 1st or in few last days of
        // previous year. Otherwise some first days of this year wouldn't belong to first
        // week, but last week of previous year (that is how ISO behaves).
        if (startOfWeekAdjust > 0)
            startOfWeekAdjust -= 7;

        var firstWeekStartDate = startOfYearDate + startOfWeekAdjust;
        var weekNum = (serialDate - firstWeekStartDate) / 7;
        return weekNum + 1;
    }

    /// <summary>
    /// WORKDAY is WORKDAY.INTL with the weekend omitted: the same walk over a Saturday and Sunday
    /// weekend, including its <c>#NUM!</c> when the answer falls outside the supported date range.
    /// </summary>
    private static ScalarValue Workday(CalcContext ctx, ScalarValue startDateScalar, ScalarValue dayOffsetValue, AnyValue holidays)
        => Workday(ctx, startDateScalar, dayOffsetValue, ScalarValue.Blank.ToAnyValue(), holidays);

    private static ScalarValue GetYear(CalcContext ctx, double serialDateTime)
    {
        if (!TryGetDate(ctx, serialDateTime, out var serialDate))
            return XLError.NumberInvalid;

        return DateParts.From(ctx, serialDate).Year;
    }

    /// <summary>
    /// Day-count fraction between two dates for one of Excel's five bases. Also used by the
    /// security functions in <see cref="Financial"/> (DISC, INTRATE, RECEIVED), which are defined
    /// in terms of "days between settlement and maturity / days in a year, per basis".
    /// </summary>
    internal static ScalarValue YearFrac(CalcContext ctx, double startDateTime, double endDateTime, double basis = 0)
    {
        if (!TryGetDate(ctx, startDateTime, out var startDate))
            return XLError.NumberInvalid;

        if (!TryGetDate(ctx, endDateTime, out var endDate))
            return XLError.NumberInvalid;

        if (basis is < 0 or >= 5)
            return XLError.NumberInvalid;

        var option = checked((int)Math.Truncate(basis));
        var yearFrac = option switch
        {
            0 => Days360(ctx, startDate, endDate, false) / 360.0,                 // US 30/360
            1 => (endDate - startDate) / GetYearAverage(ctx, startDate, endDate), // Actual/Actual
            2 => (endDate - startDate) / 360.0,                                   // Actual/360
            3 => (endDate - startDate) / 365.0,                                   // Actual/365
            _ => Days360(ctx, startDate, endDate, true) / 360.0,                  // EU 30/360
        };

        return Math.Abs(yearFrac);

        static double GetYearAverage(CalcContext ctx, int startDate, int endDate)
        {
            var startYear = DateParts.From(ctx, startDate).Year;
            var endYear = DateParts.From(ctx, endDate).Year;
            var totalDays = 0;
            for (var year = startYear; year <= endYear; year++)
            {
                // For purposes of average year, 1900 is not counted as a leap year
                totalDays += DateTime.IsLeapYear(year) ? 366 : 365;
            }

            return totalDays / (double)(endYear - startYear + 1);
        }
    }

    private static bool TryGetDate(CalcContext ctx, ScalarValue value, out int serialDate, out XLError error, bool acceptLogical = false)
    {
        // For some reason, Excel dislikes logical for things that are "date" in functions.
        // Someone likely though it was a good idea 40yrs ago.
        if (value.IsLogical && !acceptLogical)
        {
            serialDate = default;
            error = XLError.IncompatibleValue;
            return false;
        }

        if (!value.ToNumber(ctx.Culture).TryPickT0(out var serialDateTime, out error))
        {
            serialDate = default;
            return false;
        }

        if (serialDateTime is < 0 or > Year10K)
        {
            serialDate = default;
            error = XLError.NumberInvalid;
            return false;
        }

        serialDate = (int)Math.Truncate(serialDateTime);
        error = default;
        return true;
    }

    private static bool TryGetDate(CalcContext ctx, double serialDateTime, out int serialDate)
    {
        if (serialDateTime < 0 || serialDateTime >= ctx.DateSystemUpperLimit)
        {
            serialDate = default;
            return false;
        }

        serialDate = checked((int)Math.Truncate(serialDateTime));
        return true;
    }

    private static bool TryGetMonthsOffset(double monthsOffset, out int months)
    {
        // Limit enough so integer math won't overflow when added to a date
        if (monthsOffset is < -9999 * 12 or > 9999 * 12)
        {
            months = default;
            return false;
        }

        months = checked((int)monthsOffset);
        return true;
    }

    private static ScalarValue GetTimeComponent(CalcContext ctx, double serialTime, Func<DateTime, int> component)
    {
        if (serialTime < 0 || serialTime >= ctx.DateSystemUpperLimit)
            return XLError.NumberInvalid;

        return component(DateTime.FromOADate(serialTime));
    }

    /// <summary>
    /// A date type unconstrained by DateTime limitations (1900-01-00 or 1900-02-29).
    /// Has some similar methods as DateTime, but without limit checks.
    /// </summary>
    private readonly record struct DateParts(int Year, int Month, int Day)
    {
        private static readonly int[] DaysToMonth365 = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365];

        private static readonly int[] DaysToMonth366 = [0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 366];

        private static readonly DateParts Epoch1900 = new(1900, 1, 0);

        private static readonly DateParts Feb29 = new(1900, 2, 29);

        internal int SerialDate => (int)(new DateTime(Year, Month, 1, 0, 0, 0, DateTimeKind.Unspecified).ToSerialDateTime()) + Day - 1;

        public DayOfWeek DayOfWeek => (DayOfWeek)((WeekdayCalc(SerialDate) + 7 - 1) % 7);

        /// <summary>
        /// Return day of year, starting from 1 to 365/366. Counts 1900 as 366 leap year.
        /// </summary>
        public int DayOfYear
        {
            get
            {
                var startOfYear = new DateParts(Year, 1, 1);
                return SerialDate - startOfYear.SerialDate + 1;
            }
        }

        internal static DateParts From(CalcContext ctx, int serialDate)
        {
            if (ctx.Use1904DateSystem)
            {
                var date1904 = DateTime.FromOADate(serialDate + 1462);
                return From(date1904);
            }

            // Return value for 1900-01-00
            if (serialDate == 0)
                return Epoch1900;

            // Everyone loves 29th Feb 1900
            if (serialDate == 60)
                return Feb29;

            // January and February 1900. Because of non-existent feb29, adjust by one day
            if (serialDate < 60)
                serialDate++;

            return From(DateTime.FromOADate(serialDate));
        }

        private static int DaysInMonth(int year, int month)
        {
            Debug.Assert(month is >= 1 and <= 12);
            var daysToMonth = IsLeapYear(year) ? DaysToMonth366 : DaysToMonth365;
            return daysToMonth[month] - daysToMonth[month - 1];
        }

        private static bool IsLeapYear(int year)
        {
            return year % 4 == 0 && (year % 100 != 0 || year % 400 == 0);
        }

        private static DateParts From(DateTime date)
        {
            return new DateParts(date.Year, date.Month, date.Day);
        }

        internal bool IsLastDayOfMonth()
        {
            // 1900-02-29 is last day of a month per Excel, thus we have:
            // * return true for that date
            if (this == Feb29)
                return true;

            // * can't return true for real end of month
            if (Year == 1900 && Month == 2 && Day == 28)
                return false;

            return Day == DateTime.DaysInMonth(Year, Month);
        }

        internal DateParts AddMonths(int months)
        {
            var shiftedMonth = Month + months - 1;
            var adjustYear = (int)Math.Floor(shiftedMonth / 12.0);
            var year = Year + adjustYear;
            var month = shiftedMonth - adjustYear * 12 + 1;
            var day = Math.Min(Day, DaysInMonth(year, month)); // Uses real Feb28
            return new DateParts(year, month, day);
        }

        internal DateParts EndOfMonth()
        {
            return this with { Day = DaysInMonth(Year, Month) };
        }

        internal OneOf<int, XLError> ToSerialDate()
        {
            if (Year is > 9999 or < 1900)
                return XLError.NumberInvalid;

            return SerialDate;
        }
    }
}
