using System.Collections.Concurrent;
using System.Globalization;
using ExcelNumberFormat;

namespace XLibur.Extensions;

internal static class FormatExtensions
{
    // NumberFormat parses the format string in its constructor, so cache one instance per
    // distinct format string instead of re-parsing on every cell formatted for display.
    private static readonly ConcurrentDictionary<string, NumberFormat> FormatCache = new();

    public static string ToExcelFormat(this object o, string format, CultureInfo culture)
    {
        var nf = FormatCache.GetOrAdd(format, static f => new NumberFormat(f));
        return !nf.IsValid ? format : nf.Format(o, culture);
    }
}
