using System;
using System.Globalization;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// The text Excel makes of a number when a formula needs text, as in <c>=A1&amp;""</c> or
/// <c>LEN(1/3)</c>. It is not the General number format of a cell, which also depends on the
/// column width.
/// </summary>
/// <remarks>
/// The number is rounded to 15 significant digits and written in plain decimal form, so
/// <c>0.1+0.2</c> is <c>0.3</c> and <c>1E15</c> is <c>1000000000000000</c>. Scientific form is used
/// only when the plain form, without its sign, would be longer than 20 characters, so
/// <c>1E-10</c> is <c>0.0000000001</c> and <c>1E21</c> is <c>1E+21</c>. The rule is the one Apache POI's
/// <c>NumberToTextConverter</c> uses, and it agrees with every case in the
/// conformance corpus.
/// </remarks>
internal static class NumberToText
{
    private const int SignificantDigits = 15;
    private const int MaxPlainLength = 20;

    public static string Format(double number, CultureInfo culture)
    {
        if (number == 0)
            return "0";

        if (!double.IsFinite(number))
            return number.ToString(culture);

        Span<char> scientific = stackalloc char[32];
        Math.Abs(number).TryFormat(scientific, out var written, "E14", CultureInfo.InvariantCulture);
        Span<char> digits = stackalloc char[SignificantDigits];
        digits[0] = scientific[0];
        scientific.Slice(2, SignificantDigits - 1).CopyTo(digits[1..]);
        var exponent = int.Parse(scientific[(SignificantDigits + 2)..written], NumberStyles.AllowLeadingSign,
            CultureInfo.InvariantCulture);
        digits = digits[..(digits.TrimEnd('0').Length)];

        var format = culture.NumberFormat;
        var sign = number < 0 ? format.NegativeSign : string.Empty;
        var separator = format.NumberDecimalSeparator;
        var plainLength = exponent < 0
            ? "0.".Length + (-exponent - 1) + digits.Length
            : exponent + 1;
        var useScientific = plainLength > MaxPlainLength;

        if (useScientific)
        {
            var mantissa = digits.Length > 1
                ? string.Concat(digits[..1], separator, digits[1..])
                : digits.ToString();
            var exponentText = Math.Abs(exponent).ToString("00", CultureInfo.InvariantCulture);
            return string.Concat(sign, mantissa, exponent < 0 ? "E-" : "E+", exponentText);
        }

        if (exponent < 0)
            return string.Concat(sign, "0", separator, new string('0', -exponent - 1), digits.ToString());

        var integerLength = exponent + 1;
        if (digits.Length > integerLength)
            return string.Concat(sign, digits[..integerLength].ToString(), separator, digits[integerLength..].ToString());

        return string.Concat(sign, digits.ToString(), new string('0', integerLength - digits.Length));
    }
}
