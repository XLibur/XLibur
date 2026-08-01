using System;
using System.Collections.Concurrent;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

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

    private static readonly TimeSpan RegexTimeout = TimeSpan.FromMilliseconds(150);

    private static readonly string[] TimeOfDayPatterns = ["h:m tt", "h:m t", "h:m:s tt", "h:m:s t"];

    private static readonly string[] TimePatterns = ["h:m:s tt", "h:m tt", "H:m:s", "H:m", "h:m:s", "h:m"];

    public static bool TryParseCultureDate(string s, CultureInfo culture, out DateTime date)
    {
        var datePatterns = CultureSpecificPatterns.GetOrAdd(culture, static ci =>
        {
            // Patterns that look for exactly two MM/dd that aren't part of longer sequence of MMM/ddd.
            // The MM/dd matches only two digit month/day and date recognition should be more fuzzy, it
            // should recognize month/day even without leading zero. Many cultures return only MM/dd
            // (NOT M/d) on GetAllDateTimePatterns.
            const string leadingZeroMonthPattern = "(?<!M)MM(?!M)";
            const string leadingZeroDayPattern = "(?<!d)dd(?!d)";
            var shortDatePatterns = ci.DateTimeFormat.GetAllDateTimePatterns('d')
                .Concat(ci.DateTimeFormat.GetAllDateTimePatterns('D'))
                .Where(pattern => !pattern.Contains("dddd")) // It doesn't seem that Excel parser is capable of parsing day names in any culture
                .Select(pattern => Regex.Replace(pattern, leadingZeroMonthPattern, "M", RegexOptions.None, RegexTimeout)) // Recognize months even without leading zero
                .Select(pattern => Regex.Replace(pattern, leadingZeroDayPattern, "d", RegexOptions.None, RegexTimeout)) // Recognize days even without leading zero
                .Distinct().ToArray();

            // Not sure about this, but reasonably close. Hours pattern is probably generated (e.g. 'as-IN' culture
            // has AM designator before hours in patterns, but Excel requires it to be at the end). There most likely
            // isn't a pattern to just use. Example: for en-US, Excel type coercion can transform "aug 10, 2022 14:10",
            // but every single format from CultureInfo.DateTimeFormat requires AM/PM. and two digits for minutes (thus
            // the input couldn't match in any format => excel has likely it's own logic, independent of region setting).
            var timePatterns = TimePatterns;
            var longDatePatterns = shortDatePatterns
                .SelectMany(datePattern => timePatterns.Select(timePattern => string.Create(CultureInfo.InvariantCulture, $"{datePattern} {timePattern}")));

            // ISO8601 should be parseable in all cultures, not sure if Excel does. Be more forgiving, M,d instead MM,dd.
            return shortDatePatterns.Concat(longDatePatterns).Concat(["yyyy-M-d"]).Distinct().ToArray();
        });

        return DateTime.TryParseExact(s, datePatterns, culture, Style, out date);
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
                return string.Concat(s.AsSpan(0, start), designator, s.AsSpan(end));
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
        int? matchedMonth = null;
        for (var month = 1; month <= 12; month++)
        {
            var abbreviated = format.GetAbbreviatedMonthName(month);
            var full = format.GetMonthName(month);

            // An exact name needs no expansion and must keep whatever meaning it already has.
            if (string.Equals(abbreviated, prefix, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(full, prefix, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            if (!abbreviated.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
                !full.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            // Ambiguous between two months, so Excel couldn't have resolved it either.
            if (matchedMonth is not null)
                return null;

            matchedMonth = month;
        }

        return matchedMonth is null
            ? null
            : string.Concat(s.AsSpan(0, start), format.GetAbbreviatedMonthName(matchedMonth.Value), s.AsSpan(end));
    }
}
