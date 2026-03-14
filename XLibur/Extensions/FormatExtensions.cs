
using ExcelNumberFormat;
using System.Globalization;

namespace XLibur.Extensions;

internal static class FormatExtensions
{
    public static string ToExcelFormat(this object o, string format, CultureInfo culture)
    {
        var nf = new NumberFormat(format);
        if (!nf.IsValid)
            return format;

        return nf.Format(o, culture);
    }
}
