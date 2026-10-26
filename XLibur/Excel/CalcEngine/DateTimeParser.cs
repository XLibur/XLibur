using System;
using System.Collections.Concurrent;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;

namespace XLibur.Excel.CalcEngine;

internal static class DateTimeParser
{
    private const DateTimeStyles Style = DateTimeStyles.NoCurrentDateDefault | DateTimeStyles.AllowInnerWhite | DateTimeStyles.AllowTrailingWhite;

    // It's highly likely that Excel has its own database of culture specific patterns for parsing.
    // Excel has it's own parser (that accepts 1900-02-29 ^_^), never seems to parse name of a day,
    // values of hours can be up to 9999 and safely overflow...
    // Although for displaying, Excel takes a cue from region setting pattern, not so for parsing (at least
    // couldn't produce observable difference by changing setting of a culture in region dialogue).
    // .NET Core and .NET Framework also produce different patterns for GetAllDateTimePatterns.
    // This is not a perfect solution by any means, but best we can do in absence of knowledge
    // what patterns Excel uses for which cultures.
    private static readonly ConcurrentDictionary<CultureInfo, string[]> CultureSpecificPatterns = new();

    // 'h tt': Excel reads '10 am' as ten o'clock (#831).
    private static readonly string[] TimeOfDayPatterns = ["h:m tt", "h:m t", "h:m:s tt", "h:m:s t", "h tt", "h t"];

    private static readonly string[] TimePatterns = ["h:m:s tt", "h:m tt", "H:m:s", "H:m", "h:m:s", "h:m"];

    // Keyed by instance, not by CultureInfo.Equals, which compares only the name: two en-US
    // instances can have different date separators or patterns, and each must keep its own. An
    // entry goes away with its culture.
    private static readonly ConditionalWeakTable<CultureInfo, CultureInfo> ExcelTwoDigitYearCultures = new();

    /// <summary>
    /// The culture with Excel's window for a two-digit year, 1930-2029 (Windows' default), where
    /// .NET's is 1950-2049: Excel reads '22/04/30' and '1-jan-30' as 1930 (#831). Only a Gregorian
    /// calendar is changed; in another calendar 2029 is a different year. The clone is taken the
    /// first time a culture is seen, so a writable culture changed after that keeps its old settings.
    /// </summary>
    public static CultureInfo WithExcelTwoDigitYears(CultureInfo culture)
    {
        return ExcelTwoDigitYearCultures.GetValue(culture, static c =>
        {
            if (c.DateTimeFormat.Calendar is not GregorianCalendar)
                return c;

            var clone = (CultureInfo)c.Clone();
            clone.DateTimeFormat.Calendar.TwoDigitYearMax = 2029;
            return CultureInfo.ReadOnly(clone);
        });
    }

    public static bool TryParseCultureDate(string s, CultureInfo culture, out DateTime date)
    {
        var datePatterns = CultureSpecificPatterns.GetOrAdd(culture, static ci =>
        {
            // ISO 8601, and the same with '/', as Excel reads '2000/04/30' in every culture measured (#831).
            // The '/' is quoted, since a bare one stands for the culture's date separator. Be more
            // forgiving, M,d instead of MM,dd.
            string[] yearFirstPatterns = ["yyyy-M-d", "yyyy'/'M'/'d"];
            var shortDatePatterns = ci.DateTimeFormat.GetAllDateTimePatterns('d')
                .Concat(ci.DateTimeFormat.GetAllDateTimePatterns('D'))
                .Where(pattern => !pattern.Contains("dddd")) // It doesn't seem that Excel parser is capable of parsing day names in any culture
                .Select(AllowSingleDigitMonthAndDay)
                .Concat(yearFirstPatterns)
                .Distinct().ToArray();

            // Not sure about this, but reasonably close. Hours pattern is probably generated (e.g. 'as-IN' culture
            // has AM designator before hours in patterns, but Excel requires it to be at the end). There most likely
            // isn't a pattern to just use. Example: for en-US, Excel type coercion can transform "aug 10, 2022 14:10",
            // but every single format from CultureInfo.DateTimeFormat requires AM/PM. and two digits for minutes (thus
            // the input couldn't match in any format => excel has likely it's own logic, independent of region setting).
            var timePatterns = TimePatterns;
            var longDatePatterns = shortDatePatterns
                .SelectMany(datePattern => timePatterns.Select(timePattern => string.Create(CultureInfo.InvariantCulture, $"{datePattern} {timePattern}")));

            return shortDatePatterns.Concat(longDatePatterns).Distinct().ToArray();
        });

        return DateTime.TryParseExact(s, datePatterns, WithExcelTwoDigitYears(culture), Style, out date);
    }

    /// <summary>
    /// Shortens each run of exactly two <c>M</c> or two <c>d</c> to one, so the pattern also reads a
    /// month or day without a leading zero. <c>MM</c>/<c>dd</c> match only two digits, and many
    /// cultures give only those (not <c>M</c>/<c>d</c>) from <c>GetAllDateTimePatterns</c>. Longer
    /// runs, such as <c>MMM</c> or <c>ddd</c>, are names and stay as they are.
    /// </summary>
    /// <remarks>
    /// This was two regex replaces with a wall-clock timeout, which a busy server could trip (#825).
    /// </remarks>
    internal static string AllowSingleDigitMonthAndDay(string pattern)
    {
        var result = new StringBuilder(pattern.Length);
        var i = 0;
        while (i < pattern.Length)
        {
            var c = pattern[i];
            var end = i + 1;
            while (end < pattern.Length && pattern[end] == c)
                end++;

            var runLength = end - i;
            if (runLength == 2 && (c is 'M' or 'd'))
                result.Append(c);
            else
                result.Append(pattern, i, runLength);

            i = end;
        }

        return result.ToString();
    }

    /// <summary>
    /// Whether a digit is directly followed by letters that start an AM/PM designator, as in
    /// <c>10:30am</c> or <c>10a</c>. Excel gives <c>#VALUE!</c> for those, while .NET reads the space
    /// in <c>h:m tt</c> as optional (#831). Letters that only share a first letter with a designator,
    /// such as the month in <c>1apr</c>, don't count.
    /// </summary>
    public static bool HasDesignatorAgainstDigit(string s, CultureInfo c)
    {
        for (var i = 1; i < s.Length; i++)
        {
            if (!char.IsAsciiDigit(s[i - 1]) || !char.IsLetter(s[i]))
                continue;

            var end = i;
            while (end < s.Length && char.IsLetter(s[end]))
                end++;

            if (IsDesignatorPrefix(s.AsSpan(i, end - i), c))
                return true;
        }

        return false;
    }

    private static bool IsDesignatorPrefix(ReadOnlySpan<char> letters, CultureInfo c)
    {
        ReadOnlySpan<string> designators =
        [
            c.DateTimeFormat.AMDesignator,
            c.DateTimeFormat.PMDesignator,
            CultureInfo.InvariantCulture.DateTimeFormat.AMDesignator,
            CultureInfo.InvariantCulture.DateTimeFormat.PMDesignator
        ];

        foreach (var designator in designators)
        {
            if (designator.Length > 0 && designator.AsSpan().StartsWith(letters, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return false;
    }

    public static bool TryParseTimeOfDay(string s, CultureInfo c, out DateTime timeOfDay)
    {
        if (TryParseTimeOfDayExact(s, c, out timeOfDay))
            return true;

        // Excel accepts a shortened AM/PM designator ('12:0:18 odp' for the cs-CZ 'odp.') and tolerates
        // a trailing dot the designator doesn't have ('12:0:18 PM.'). Neither is a pattern .NET can
        // express, so canonicalize the designator and try once more.
        var canonicalDesignator = CanonicalizeDesignator(s, c);
        return canonicalDesignator is not null && TryParseTimeOfDayExact(canonicalDesignator, c, out timeOfDay);
    }

    private static bool TryParseTimeOfDayExact(string s, CultureInfo c, out DateTime timeOfDay)
    {
        if (DateTime.TryParseExact(s, TimeOfDayPatterns, c, Style, out timeOfDay))
            return true;

        if (DateTime.TryParseExact(s, TimeOfDayPatterns, CultureInfo.InvariantCulture, Style, out timeOfDay))
            return true;

        return false;
    }

    /// <summary>
    /// Replaces a trailing AM/PM designator that is a prefix of a real designator, optionally with a
    /// superfluous trailing dot, by the designator itself. Returns <c>null</c> when the last token
    /// isn't such a designator, in which case there is nothing to retry.
    /// </summary>
    private static string? CanonicalizeDesignator(string s, CultureInfo c)
    {
        var end = s.Length;
        while (end > 0 && s[end - 1] == ' ')
            end--;

        var start = end;
        while (start > 0 && s[start - 1] != ' ')
            start--;

        if (start == end)
            return null;

        var token = s.Substring(start, end - start);
        var withoutTrailingDot = token.TrimEnd('.');
        if (withoutTrailingDot.Length == 0)
            return null;

        var designator = FindDesignatorExtending(token, withoutTrailingDot, c);
        return designator is null
            ? null
            : string.Concat(s.AsSpan(0, start), designator, s.AsSpan(end));
    }

    /// <summary>
    /// Finds the first AM/PM designator (culture, then invariant) that starts with
    /// <paramref name="withoutTrailingDot"/> but isn't exactly <paramref name="token"/>.
    /// </summary>
    private static string? FindDesignatorExtending(string token, string withoutTrailingDot, CultureInfo c)
    {
        ReadOnlySpan<string> designators =
        [
            c.DateTimeFormat.AMDesignator,
            c.DateTimeFormat.PMDesignator,
            CultureInfo.InvariantCulture.DateTimeFormat.AMDesignator,
            CultureInfo.InvariantCulture.DateTimeFormat.PMDesignator
        ];

        foreach (var designator in designators)
        {
            // An exact match already had its chance in the pass before this one.
            if (designator.Length == 0 || string.Equals(token, designator, StringComparison.OrdinalIgnoreCase))
                continue;

            if (designator.StartsWith(withoutTrailingDot, StringComparison.OrdinalIgnoreCase))
                return designator;
        }

        return null;
    }

    /// <summary>
    /// Excel matches a month by any prefix of its name at least three letters long, .NET only by the
    /// exact abbreviation or full name. Replaces the first such prefix in <paramref name="s"/> with
    /// the culture's abbreviated month name, or returns <c>null</c> when there is nothing to expand
    /// or the prefix is ambiguous between months.
    /// </summary>
    public static string? ExpandMonthNamePrefix(string s, CultureInfo c)
    {
        var start = 0;
        while (start < s.Length && !char.IsLetter(s[start]))
            start++;

        var end = start;
        while (end < s.Length && char.IsLetter(s[end]))
            end++;

        // Excel requires at least three letters, so 'ma' stays ambiguous between March and May.
        var prefix = s.Substring(start, end - start);
        if (prefix.Length < 3)
            return null;

        var format = c.DateTimeFormat;
        var matchedMonth = FindUniqueMonthByPrefix(prefix, format);
        return matchedMonth is null
            ? null
            : string.Concat(s.AsSpan(0, start), format.GetAbbreviatedMonthName(matchedMonth.Value), s.AsSpan(end));
    }

    /// <summary>
    /// Returns the only month whose abbreviated or full name starts with <paramref name="prefix"/>,
    /// or <c>null</c> when no month matches, more than one does, or a name equals the prefix exactly.
    /// </summary>
    private static int? FindUniqueMonthByPrefix(string prefix, DateTimeFormatInfo format)
    {
        int? matchedMonth = null;
        for (var month = 1; month <= 12; month++)
        {
            var abbreviated = format.GetAbbreviatedMonthName(month);
            var full = format.GetMonthName(month);

            // An exact name needs no expansion and must keep whatever meaning it already has.
            if (EqualsIgnoreCase(abbreviated, prefix) || EqualsIgnoreCase(full, prefix))
                return null;

            if (!StartsWithIgnoreCase(abbreviated, prefix) && !StartsWithIgnoreCase(full, prefix))
                continue;

            // Ambiguous between two months, so Excel couldn't have resolved it either.
            if (matchedMonth is not null)
                return null;

            matchedMonth = month;
        }

        return matchedMonth;
    }

    private static bool EqualsIgnoreCase(string name, string prefix) =>
        string.Equals(name, prefix, StringComparison.OrdinalIgnoreCase);

    private static bool StartsWithIgnoreCase(string name, string prefix) =>
        name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
}
