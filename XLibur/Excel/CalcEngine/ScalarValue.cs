using System;
using System.Globalization;
using XLibur.Extensions;

#pragma warning disable S1244 // Intentional exact float comparison for Excel formula compatibility

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// A representation of a value as a discriminated union.
/// </summary>
/// <remarks>
/// A bare bone copy of <c>OneOf</c> that can be more optimized:
/// <list type="bullet">
///   <item>readonly struct to get rid of defensive copies</item>
///   <item>struct can be smaller through offsets (based on NoBox)</item>
///   <item>allows to pass additional arguments to Match function to skip a need to instantiate a new lambda instance on each call and allow easier inlining.</item>
/// </list>
/// </remarks>
internal readonly struct ScalarValue
{
    private const int BlankValue = 0;
    private const int LogicalValue = 1;
    private const int NumberValue = 2;
    private const int TextValue = 3;
    private const int ErrorValue = 4;

    private readonly byte _index;
    private readonly bool _logical;
    private readonly double _number;
    private readonly string? _text;
    private readonly XLError _error;

    private ScalarValue(byte index, bool logical, double number, string? text, XLError error)
    {
        _index = index;
        _logical = logical;
        _number = number;
        _text = text;
        _error = error;
    }

    /// <summary>
    /// Internal accessor for the text payload. Callers must verify that <see cref="_index"/>
    /// is <see cref="TextValue"/> before invoking; this property centralises the
    /// null-forgiving cast in one place instead of sprinkling <c>_text!</c> through every
    /// switch arm.
    /// </summary>
    private string Text => _text!;

    /// <summary>
    /// A blank value of a scalar. It can behave as a 0 or empty string, depending on context.
    /// </summary>
    /// <example><c>A1+5</c> is a number 5, blank behaves as 0, <c>A1 &amp; "text"</c> is a "text", blank behaves as empty string.</example>
    public static readonly ScalarValue Blank = new(BlankValue, default, default, default, default);

    public bool IsBlank => _index == BlankValue;

    public bool IsLogical => _index == LogicalValue;

    public bool IsNumber => _index == NumberValue;

    public bool IsText => _index == TextValue;

    public bool IsError => _index == ErrorValue;

    public static ScalarValue From(bool logical) => new(LogicalValue, logical, default, default, default);

    public static ScalarValue From(double number) => new(NumberValue, default, number, default, default);

    public static ScalarValue From(string text)
    {
        ArgumentNullException.ThrowIfNull(text);

        return new ScalarValue(TextValue, default, default, text, default);
    }

    public static ScalarValue From(XLError error) => new(ErrorValue, default, default, default, error);

    public static implicit operator ScalarValue(bool logical) => From(logical);

    public static implicit operator ScalarValue(double number) => From(number);

    public static implicit operator ScalarValue(string text) => From(text);

    public static implicit operator ScalarValue(XLError error) => From(error);

    public static implicit operator ScalarValue(XLCellValue cellValue)
    {
        return cellValue.Type switch
        {
            XLDataType.Blank => Blank,
            XLDataType.Boolean => cellValue.GetBoolean(),
            XLDataType.Number => cellValue.GetNumber(),
            XLDataType.Text => cellValue.GetText(),
            XLDataType.Error => cellValue.GetError(),
            XLDataType.DateTime => cellValue.GetDateTime().ToSerialDateTime(),
            XLDataType.TimeSpan => cellValue.GetTimeSpan().ToSerialDateTime(),
            _ => throw new InvalidOperationException()
        };
    }

    public bool GetLogical() => IsLogical ? _logical : throw new InvalidCastException();

    public double GetNumber() => IsNumber ? _number : throw new InvalidCastException();

    public string GetText() => IsText ? Text : throw new InvalidCastException();

    public XLError GetError() => IsError ? _error : throw new InvalidCastException();

    internal XLCellValue ToCellValue()
    {
        return _index switch
        {
            BlankValue => 0, // The result value of a formula calculation can be blank, but result of formula in a cell value is never blank, but 0.
            LogicalValue => _logical,
            NumberValue => _number,
            TextValue => Text,
            ErrorValue => _error,
            _ => throw new InvalidOperationException()
        };
    }

    public TResult Match<TResult>(Func<TResult> transformBlank, Func<bool, TResult> transformLogical, Func<double, TResult> transformNumber, Func<string, TResult> transformText, Func<XLError, TResult> transformError)
    {
        return _index switch
        {
            BlankValue => transformBlank(),
            LogicalValue => transformLogical(_logical),
            NumberValue => transformNumber(_number),
            TextValue => transformText(Text),
            ErrorValue => transformError(_error),
            _ => throw new InvalidOperationException()
        };
    }

    public TResult Match<TResult, TParam1>(TParam1 param, Func<TParam1, TResult> transformBlank, Func<bool, TParam1, TResult> transformLogical, Func<double, TParam1, TResult> transformNumber, Func<string, TParam1, TResult> transformText, Func<XLError, TParam1, TResult> transformError)
    {
        return _index switch
        {
            BlankValue => transformBlank(param),
            LogicalValue => transformLogical(_logical, param),
            NumberValue => transformNumber(_number, param),
            TextValue => transformText(Text, param),
            ErrorValue => transformError(_error, param),
            _ => throw new InvalidOperationException()
        };
    }

    public TResult Match<TResult, TParam1, TParam2>(TParam1 param1, TParam2 param2, Func<TParam1, TParam2, TResult> transformBlank, Func<bool, TParam1, TParam2, TResult> transformLogical, Func<double, TParam1, TParam2, TResult> transformNumber, Func<string, TParam1, TParam2, TResult> transformText, Func<XLError, TParam1, TParam2, TResult> transformError)
    {
        return _index switch
        {
            BlankValue => transformBlank(param1, param2),
            LogicalValue => transformLogical(_logical, param1, param2),
            NumberValue => transformNumber(_number, param1, param2),
            TextValue => transformText(Text, param1, param2),
            ErrorValue => transformError(_error, param1, param2),
            _ => throw new InvalidOperationException()
        };
    }

    public AnyValue ToAnyValue()
    {
        return _index switch
        {
            BlankValue => AnyValue.Blank,
            LogicalValue => _logical,
            NumberValue => _number,
            TextValue => Text,
            ErrorValue => _error,
            _ => throw new InvalidOperationException()
        };
    }

    /// <summary>
    /// Convert value to text. Error is not convertible.
    /// </summary>
    public OneOf<string, XLError> ToText(CultureInfo culture)
    {
        return _index switch
        {
            BlankValue => string.Empty,
            LogicalValue => _logical ? "TRUE" : "FALSE",
            NumberValue => _number.ToString(culture),
            TextValue => Text,
            ErrorValue => _error,
            _ => throw new InvalidOperationException()
        };
    }

    /// <summary>
    /// Convert value to number. Error is not convertible.
    /// </summary>
    public OneOf<double, XLError> ToNumber(CultureInfo culture)
    {
        return _index switch
        {
            BlankValue => 0,
            LogicalValue => _logical ? 1.0 : 0.0,
            NumberValue => _number,
            TextValue => TextToNumber(Text, culture),
            ErrorValue => _error,
            _ => throw new InvalidOperationException()
        };
    }

    /// <summary>
    /// Parse text to a scalar value. Generally used in formulas or autofilter.
    /// </summary>
    /// <param name="text">Text to parse.</param>
    /// <param name="culture">Culture used for parsing numbers or dates.</param>
    /// <returns>Parsed scalar value.</returns>
    public static ScalarValue Parse(string text, CultureInfo culture)
    {
        if (text is null)
            return Blank;
        if (text.Length == 0)
            return Blank;
        if (StringComparer.OrdinalIgnoreCase.Equals("TRUE", text))
            return true;
        if (StringComparer.OrdinalIgnoreCase.Equals("FALSE", text))
            return false;
        if (TextToNumber(text, culture).TryPickT0(out var number, out _))
            return number;
        if (XLErrorParser.TryParseError(text, out var error))
            return error;

        return text;
    }

    public static OneOf<double, XLError> TextToNumber(string text, CultureInfo culture)
    {
        if (string.IsNullOrWhiteSpace(text))
            return XLError.IncompatibleValue;

        // Numbers. The parsing method recognizes braces as negative number, includes currency parsing.
        // Format 1 '0'
        //        2 '0.00'
        //        3 '#,##0'
        //        4 '#,##0.00'
        //       11 '0.00E+00'
        //       48 '##0.0E+0'
        if (NumberParser.TryParse(text, culture, out var number))
            return number;

        // Excel allows whitespace between the sign and the number ('- 100 %'), .NET parse methods
        // don't. Only reached once the parse above has already refused the text as it stands.
        var withoutSignWhitespace = RemoveWhitespaceAfterLeadingSign(text, culture);
        if (withoutSignWhitespace is not null)
            return TextToNumber(withoutSignWhitespace, culture);

        // Excel reads a value in braces as negative even when the braces contain whitespace or a
        // percent sign, neither of which NumberStyles.AllowParentheses accepts.
        var bracketedValue = RemoveNegatingBraces(text, culture);
        if (bracketedValue is not null)
        {
            return TextToNumber(bracketedValue, culture)
                .TryPickT0(out var bracketedNumber, out var bracketedError)
                ? -bracketedNumber
                : bracketedError;
        }

        // Percents. Percent sign can be at both sides.
        // Format 9 '0%'
        //       10 '0.00%'
        var textSpan = text.AsSpan(); // Avoid extra allocations for trimming/substrings if not match
        var textSpanTrimmedEnd = textSpan.TrimEnd();
        var percentSymbol = culture.NumberFormat.PercentSymbol.AsSpan();
        if (textSpanTrimmedEnd.EndsWith(percentSymbol))
            return ParsePercent(text, 0, textSpanTrimmedEnd.Length - percentSymbol.Length, culture);

        var textSpanTrimmedStart = textSpan.TrimStart();
        if (textSpanTrimmedStart.StartsWith(percentSymbol))
        {
            var newStart = text.Length - textSpanTrimmedStart.Length + percentSymbol.Length;
            return ParsePercent(text, newStart, text.Length - newStart, culture);
        }

        // Fractions
        // Format 12 '# ?/?'
        //        13 '# ??/??'
        if (FractionParser.TryParse(text, out var fraction))
            return fraction;

        if (ToSerialDateTime(text, culture, out var serialDateTime))
            return serialDateTime;

        return XLError.IncompatibleValue;

        static OneOf<double, XLError> ParsePercent(string text, int start, int length, CultureInfo c)
        {
            text = text.Substring(start, length);
            if (double.TryParse(text, NumberStyles.Float
                                      | NumberStyles.AllowThousands
                                      | NumberStyles.AllowParentheses, c, out var percents))
                return percents / 100;

            // other formats don't use '%' sign, but text has it, so just stop for invalid inputs like 'hundred%'
            return XLError.IncompatibleValue;
        }

        // Returns the text with the whitespace between a leading sign and the value removed, or null
        // when there is no leading sign or no whitespace behind it.
        static string? RemoveWhitespaceAfterLeadingSign(string text, CultureInfo c)
        {
            var signStart = SkipSpaces(text, 0);
            var signLength = MatchSign(text, signStart, c);
            if (signLength == 0)
                return null;

            var valueStart = SkipSpaces(text, signStart + signLength);
            return valueStart == signStart + signLength
                ? null
                : string.Concat(text.AsSpan(0, signStart + signLength), text.AsSpan(valueStart));
        }

        // Returns the content of the braces, or null when the text isn't braced. A sign inside the
        // braces is rejected, because in Excel the braces themselves are the sign ('(-1)' is not a
        // number, neither is '-(1)').
        static string? RemoveNegatingBraces(string text, CultureInfo c)
        {
            var start = SkipSpaces(text, 0);
            var end = text.Length - 1;
            while (end >= 0 && text[end] == ' ')
                end--;

            if (start >= end || text[start] != '(' || text[end] != ')')
                return null;

            var content = text.Substring(start + 1, end - start - 1);
            return MatchSign(content, SkipSpaces(content, 0), c) != 0 ? null : content;
        }

        static int SkipSpaces(string text, int index)
        {
            while (index < text.Length && text[index] == ' ')
                index++;

            return index;
        }

        // Length of the culture's positive or negative sign at the index, or 0 if neither is there.
        static int MatchSign(string text, int index, CultureInfo c)
        {
            var negativeSign = c.NumberFormat.NegativeSign;
            if (MatchesAt(text, index, negativeSign))
                return negativeSign.Length;

            var positiveSign = c.NumberFormat.PositiveSign;
            return MatchesAt(text, index, positiveSign) ? positiveSign.Length : 0;
        }

        static bool MatchesAt(string text, int index, string value)
        {
            return value.Length > 0 &&
                   index + value.Length <= text.Length &&
                   string.CompareOrdinal(text, index, value, 0, value.Length) == 0;
        }
    }

    public static bool ToSerialDateTime(string text, CultureInfo culture, out double serialDateTime)
    {
        if (TryParseDatePatterns(text, culture, out serialDateTime))
            return true;

        // Excel accepts a month name of anything from three letters up to the full name ('1-marc'),
        // while .NET recognizes only the exact abbreviation or the exact full name. Expand the prefix
        // to a name .NET knows and try once more. Done as a second pass so that every input the
        // patterns already handle keeps taking the untouched path above.
        var expandedMonthName = DateTimeParser.ExpandMonthNamePrefix(text, culture);
        return expandedMonthName is not null && TryParseDatePatterns(expandedMonthName, culture, out serialDateTime);
    }

    private static bool TryParseDatePatterns(string text, CultureInfo culture, out double serialDateTime)
    {
        const DateTimeStyles dateStyle = DateTimeStyles.NoCurrentDateDefault | DateTimeStyles.AllowInnerWhite | DateTimeStyles.AllowTrailingWhite;

        // This date varies by the culture. Keep first before other standard patterns. Must be for both yy and yyyy.
        // Format 14 : short date (for en 'm/d/yyyy')
        // Format 22 : short date + hours (for en 'm/d/yyyy h:mm')
        if (DateTimeParser.TryParseCultureDate(text, culture, out var dateFormat14Or22))
        {
            return ToSerialDate(dateFormat14Or22, out serialDateTime);
        }

        // Excel lets one component of the time overflow its normal range and carries the excess into
        // the date, e.g. '11/30/2022 24:59' is one minute to one in the morning of December 1st. No
        // DateTime pattern can express that, so split the text and let TimeSpanParser, which already
        // models Excel's overflow rules, handle the time half.
        if (TryParseDateWithOverflowTime(text, culture, out serialDateTime))
        {
            return true;
        }

        // Whether a leading space is allowed is a property of the individual format, not of the
        // parser: format 16 accepts ' 1 - apr  ', while formats 15 and 17 reject the same space.
        // .NET makes no such distinction and allows it everywhere, so the strict formats are guarded.
        var hasLeadingWhitespace = text.Length > 0 && char.IsWhiteSpace(text[0]);

        // Date with names of months. The names of months differ across cultures.
        // Format 15 'd-mmm-yy'
        if (!hasLeadingWhitespace &&
            DateTime.TryParseExact(text, ["d-MMM-yyyy", "d-MMMM-yyyy", "d-MMM-yy", "d-MMMM-yy",
                                          "d-MMM-yyyy h:m", "d-MMMM-yyyy h:m", "d-MMM-yy h:m", "d-MMMM-yy h:m",
                                          "d-MMM-yyyy h:m:s", "d-MMMM-yyyy h:m:s", "d-MMM-yy h:m:s", "d-MMMM-yy h:m:s"], culture, dateStyle, out var dateFormat15))
        {
            return ToSerialDate(dateFormat15, out serialDateTime);
        }

        // Since format doesn't have a year, it uses current year
        // Format 16 'd-mmm'
        if (DateTime.TryParseExact(text, ["d-MMM", "d-MMMM"], culture, dateStyle, out var dateFormat16))
        {
            return ToSerialDate(dateFormat16, out serialDateTime);
        }

        // Excel has an extra 'mmm-dd' pattern ahead of 'mmm-yy' in cultures that write the month
        // before the day, so under en-US 'jan-02' is the second of January of the current year rather
        // than January 2002. Cultures that write the day first only have the year reading, which is
        // why cs-CZ reads 'led-5' as January 2005. Parsing happens in year 1 (NoCurrentDateDefault),
        // so a number that isn't a valid day falls through to the year reading below, and so does
        // 'feb-29', which no year 1 can hold.
        if (!hasLeadingWhitespace &&
            IsMonthBeforeDay(culture) &&
            DateTime.TryParseExact(text, ["MMM-d", "MMMM-d"], culture, dateStyle, out var dateFormat17AsDay))
        {
            var dayInCurrentYear = dateFormat17AsDay.AddYears(DateTime.Now.Year - dateFormat17AsDay.Year);
            return ToSerialDate(dayInCurrentYear, out serialDateTime);
        }

        // Month and a number. In some cultures, the culture date parsing will interpret this pattern as MMM-dd, but
        // that depends on culture date patterns above. Use MMM and MMMM to encompass both abbreviation and full name.
        // Format 17 'mmm-yy'
        if (!hasLeadingWhitespace &&
            DateTime.TryParseExact(text, ["MMM-y", "MMMM-y"], culture, dateStyle, out var dateFormat17))
        {
            if (dateFormat17.Year != DateTime.Now.Year && dateFormat17.Year >= 2030)
                dateFormat17 = dateFormat17.AddYears(-100);

            return ToSerialDate(dateFormat17, out serialDateTime);
        }

        // Format 18 'h:mm AM/PM', works for both localized and AM/PM literal
        // Format 19 'h:mm:ss AM/PM'
        if (DateTimeParser.TryParseTimeOfDay(text, culture, out var dateFormat18Or19))
        {
            serialDateTime = dateFormat18Or19.ToOADate();
            return true;
        }

        // Time span uses a different parser from time of a day.
        // Format 20 'h:mm'
        //        21 'h:mm:ss'
        //        47 'mm:ss.0'
        if (TimeSpanParser.TryParseTime(text, culture, out var timeSpan))
        {
            serialDateTime = timeSpan.ToSerialDateTime();
            return true;
        }

        serialDateTime = default;
        return false;

#pragma warning disable S3776 // Every split point of 'date time' has to be tried; the guards are the search
        static bool TryParseDateWithOverflowTime(string text, CultureInfo culture, out double serialDateTime)
        {
            serialDateTime = default;

            // The date and the time are separated by a space, but the date itself may contain spaces
            // ('aug 10, 2022 14:10'), so every split point has to be tried. Search from the end,
            // where the time is far more likely to start.
            for (var i = text.Length - 1; i > 0; i--)
            {
                if (text[i] != ' ')
                    continue;

                var datePart = text.Substring(0, i);
                var timePart = text.Substring(i + 1);
                if (timePart.Length == 0)
                    continue;

                if (!DateTimeParser.TryParseCultureDate(datePart, culture, out var date))
                    continue;

                if (!TimeSpanParser.TryParseTime(timePart, culture, out var time))
                    continue;

                if (!ToSerialDate(date, out var serialDate))
                    return false;

                serialDateTime = serialDate + time.ToSerialDateTime();
                return true;
            }

            return false;
        }
#pragma warning restore S3776

        // Whether the culture writes '3/1' as the first of March or the third of January.
        static bool IsMonthBeforeDay(CultureInfo c)
        {
            var shortDatePattern = c.DateTimeFormat.ShortDatePattern;
            var monthIndex = shortDatePattern.IndexOf('M');
            var dayIndex = shortDatePattern.IndexOf('d');
            return monthIndex >= 0 && dayIndex >= 0 && monthIndex < dayIndex;
        }

        static bool ToSerialDate(DateTime dateTime, out double serialDate)
        {
            if (dateTime.Year < 1900)
            {
                serialDate = default;
                return false;
            }

            // Excel says 1900 was a leap year  :( Replicate an incorrect behavior thanks
            // to Lotus 1-2-3 decision from 1983...
            var oDate = dateTime.ToOADate();
            const int nonExistent1900Feb29SerialDate = 60;
            serialDate = oDate <= nonExistent1900Feb29SerialDate ? oDate - 1 : oDate;
            return true;
        }
    }

    public bool TryPickLogical(out bool logical)
    {
        if (_index == LogicalValue)
        {
            logical = _logical;
            return true;
        }

        logical = default;
        return false;
    }

    public bool TryPickNumber(out double number)
    {
        return TryPickNumber(out number, out _);
    }

    public bool TryPickNumber(out double number, out XLError error)
    {
        if (_index == NumberValue)
        {
            number = _number;
            error = default;
            return true;
        }

        number = default;
        error = IsError ? _error : XLError.IncompatibleValue;
        return false;
    }

    /// <summary>
    /// Try to pick a number (interpret blank as number 0).
    /// </summary>
    public bool TryPickNumberOrBlank(out double number, out XLError error)
    {
        if (_index == NumberValue)
        {
            number = _number;
            error = default;
            return true;
        }

        // This is mostly useful for unified approach area + array. Literal array
        // can't contain blanks, but area can. In most cases, blank is interpreted as 0.
        if (_index == BlankValue)
        {
            number = 0;
            error = default;
            return true;
        }

        number = default;
        error = IsError ? _error : XLError.IncompatibleValue;
        return false;
    }

    public bool TryPickText(out string? text, out XLError error)
    {
        if (_index == TextValue)
        {
            text = _text;
            error = default;
            return true;
        }

        text = default;
        error = _index == ErrorValue ? _error : XLError.IncompatibleValue;
        return false;
    }

    public bool TryPickError(out XLError error)
    {
        if (_index == ErrorValue)
        {
            error = _error;
            return true;
        }

        error = default;
        return false;
    }

    /// <summary>
    /// Does this value have same type as the other one?
    /// </summary>
    public bool HaveSameType(ScalarValue other) => _index == other._index;

    /// <summary>
    /// Get the logical value, if it is either blank (false), logical or number (0 = false, otherwise true)a text <c>TRUE</c> or <c>FALSE</c> (case insensitive).
    /// </summary>
    /// <remarks>Used for coercion in functions.</remarks>
    public bool TryCoerceLogicalOrBlankOrNumberOrText(out bool value, out XLError error)
    {
        switch (_index)
        {
            case BlankValue:
            case TextValue when (StringComparer.OrdinalIgnoreCase.Equals(Text, "FALSE")):
                value = false;
                error = default;
                return true;
            case LogicalValue:
                value = _logical;
                error = default;
                return true;
            case NumberValue:
                value = _number != 0;
                error = default;
                return true;
            case TextValue when (StringComparer.OrdinalIgnoreCase.Equals(Text, "TRUE")):
                value = true;
                error = default;
                return true;
            case ErrorValue:
                value = default;
                error = _error;
                return false;
            default:
                value = default;
                error = XLError.IncompatibleValue;
                return false;
        }
    }

    public override string ToString()
    {
        return _index switch
        {
            BlankValue => "Blank",
            LogicalValue => _logical.ToString().ToUpper(),
            NumberValue => _number.ToString(CultureInfo.InvariantCulture),
            TextValue => Text,
            ErrorValue => _error.ToDisplayString(),
            _ => throw new InvalidOperationException("Invalid type of scalar value.")
        };
    }
}
