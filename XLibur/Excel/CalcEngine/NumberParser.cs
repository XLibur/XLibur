using System;
using System.Globalization;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// A parser of the plain number formats used by Excel during coercion from text to number. It is
/// <see cref="double.TryParse(string, NumberStyles, IFormatProvider, out double)" /> with
/// <see cref="NumberStyles.Any" />, narrowed by two rules that method is deliberately lenient about
/// and Excel is not:
/// <list type="bullet">
///     <item>Group separators must actually group. .NET accepts a separator anywhere in the integer
///           part, so <c>1,00</c> and <c>1,00,000</c> both parse as numbers under en-US even though
///           neither is grouped the way the culture groups.</item>
///     <item>The currency symbol must be where the culture puts it. .NET accepts it at either end,
///           so <c>1$</c> parses under en-US and <c>Kč 1</c> under cs-CZ, while Excel rejects both.</item>
/// </list>
/// </summary>
internal static class NumberParser
{
    public static bool TryParse(string text, CultureInfo culture, out double number)
    {
        if (double.TryParse(text, NumberStyles.Any, culture, out number) &&
            HasValidGroupSeparators(text, culture) &&
            HasValidCurrencyPosition(text, culture))
        {
            return true;
        }

        number = 0;
        return false;
    }

    /// <summary>
    /// Checks that every group separator in the integer part sits on a group boundary. Only the
    /// leftmost group may be shorter than its group size, which is why the groups can't be validated
    /// until they have all been seen.
    /// </summary>
#pragma warning disable S3776 // Group sizes are validated right-to-left with the leftmost group exempt
    private static bool HasValidGroupSeparators(string text, CultureInfo culture)
    {
        var format = culture.NumberFormat;
        var hasCurrency = ContainsCurrencySymbol(text, culture);
        var separator = hasCurrency ? format.CurrencyGroupSeparator : format.NumberGroupSeparator;
        var start = IndexOfFirstDigit(text);
        if (separator.Length == 0 || start < 0)
            return true;

        var end = start;
        while (end < text.Length)
        {
            if (char.IsDigit(text[end]))
                end++;
            else if (MatchesAt(text, end, separator))
                end += separator.Length;
            else
                break;
        }

        var groups = text.Substring(start, end - start).Split(separator);
        if (groups.Length == 1)
            return true;

        var sizes = hasCurrency ? format.CurrencyGroupSizes : format.NumberGroupSizes;
        for (var i = 0; i < groups.Length; i++)
        {
            // A size of zero means the culture stops grouping past this point, so anything to the
            // left of it is one unbounded group.
            var size = GroupSizeFromRight(sizes, groups.Length - 1 - i);
            var valid = i == 0
                ? groups[i].Length >= 1 && (size == 0 || groups[i].Length <= size)
                : size != 0 && groups[i].Length == size;

            if (!valid)
                return false;
        }

        return true;
    }
#pragma warning restore S3776

    /// <summary>
    /// Group sizes are listed from the right, and the last one repeats for everything further left.
    /// </summary>
    private static int GroupSizeFromRight(int[] sizes, int indexFromRight)
    {
        return sizes.Length == 0 ? 0 : sizes[Math.Min(indexFromRight, sizes.Length - 1)];
    }

    private static bool HasValidCurrencyPosition(string text, CultureInfo culture)
    {
        var symbol = culture.NumberFormat.CurrencySymbol;
        if (symbol.Length == 0)
            return true;

        var symbolIndex = text.IndexOf(symbol, StringComparison.Ordinal);
        var digitIndex = IndexOfFirstDigit(text);
        if (symbolIndex < 0 || digitIndex < 0)
            return true;

        // Patterns 0 ('$n') and 2 ('$ n') lead with the symbol, 1 ('n$') and 3 ('n $') trail with it.
        // The negative patterns wrap the same arrangement in a sign or braces, so they don't change
        // which side of the digits the symbol falls on.
        var symbolLeads = culture.NumberFormat.CurrencyPositivePattern is 0 or 2;
        return symbolLeads == (symbolIndex < digitIndex);
    }

    private static bool ContainsCurrencySymbol(string text, CultureInfo culture)
    {
        var symbol = culture.NumberFormat.CurrencySymbol;
        return symbol.Length > 0 && text.Contains(symbol, StringComparison.Ordinal);
    }

    private static int IndexOfFirstDigit(string text)
    {
        for (var i = 0; i < text.Length; i++)
        {
            if (char.IsDigit(text[i]))
                return i;
        }

        return -1;
    }

    private static bool MatchesAt(string text, int index, string value)
    {
        return value.Length > 0 &&
               index + value.Length <= text.Length &&
               string.CompareOrdinal(text, index, value, 0, value.Length) == 0;
    }
}
