using System;
using System.Globalization;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// Parse a fraction for text-to-number type coercion.
/// </summary>
internal static class FractionParser
{
    /// <summary>
    /// Parses text of the form <c>[sign] whole numerator/denominator</c>, such as <c>1 1/2</c>
    /// or <c>- 12 05/10</c>. Spaces (U+0020 only) are allowed around the sign and at both ends,
    /// there is exactly one space between the whole number and the numerator, and the numerator
    /// and denominator are 1 to 5 ASCII digits each.
    /// </summary>
    public static bool TryParse(string s, out double result)
    {
        result = 0;

        // A scan, not a regex. This was a regex with a 150 ms timeout, and the timeout is
        // wall-clock time: a thread descheduled or paused by GC mid-match threw
        // RegexMatchTimeoutException whatever the text. TextToNumber reaches this for every text
        // that is not a plain number, so a busy server failed to evaluate formulas, run criteria
        // functions and load autofilters at random (#817).
        var text = s.AsSpan();
        var index = SkipSpaces(text, 0);

        var hasNegativeSign = false;
        if (index < text.Length && text[index] is '+' or '-')
        {
            hasNegativeSign = text[index] == '-';
            index = SkipSpaces(text, index + 1);
        }

        var wholeStart = index;
        index = SkipDigits(text, index);
        if (index == wholeStart)
            return false;

        var whole = text[wholeStart..index];

        // Exactly one space between the whole number and the numerator.
        if (index == text.Length || text[index] != ' ')
            return false;

        index++;
        if (!TryReadFractionPart(text, ref index, out var numerator) || index == text.Length || text[index] != '/')
            return false;

        index++;
        if (!TryReadFractionPart(text, ref index, out var denominator) || SkipSpaces(text, index) != text.Length)
            return false;

        if (denominator == 0 || denominator > short.MaxValue || numerator > short.MaxValue)
            return false;

        // The whole number has no length limit, so it is read as a double and not an int. Excel
        // reads =VALUE("99999999999 1/2") as 99999999999.5. A whole number beyond the range of a
        // double is not a number, and double.Parse gives infinity for it instead of throwing.
        var wholeNumber = double.Parse(whole, NumberStyles.None, CultureInfo.InvariantCulture);
        if (double.IsInfinity(wholeNumber))
            return false;

        var fraction = wholeNumber + numerator / (double)denominator;
        result = hasNegativeSign ? -fraction : fraction;
        return true;
    }

    private static int SkipSpaces(ReadOnlySpan<char> text, int index)
    {
        while (index < text.Length && text[index] == ' ')
            index++;

        return index;
    }

    // Not char.IsDigit, as the regex's '[0-9]' matched only ASCII digits.
    private static int SkipDigits(ReadOnlySpan<char> text, int index)
    {
        while (index < text.Length && char.IsAsciiDigit(text[index]))
            index++;

        return index;
    }

    /// <summary>
    /// Reads the numerator or the denominator: 1 to 5 digits. Leading zeros count towards the
    /// five, so <c>000001</c> is too long.
    /// </summary>
    private static bool TryReadFractionPart(ReadOnlySpan<char> text, ref int index, out int value)
    {
        var start = index;
        value = 0;
        while (index < text.Length && char.IsAsciiDigit(text[index]))
        {
            if (index - start == 5)
                return false;

            value = value * 10 + (text[index] - '0');
            index++;
        }

        return index != start;
    }
}
