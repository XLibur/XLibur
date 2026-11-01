using System;
using System.Globalization;
using System.Numerics;
using XLibur.Excel;

namespace XLibur.Report.Excel;

/// <summary>
/// Turns the object an expression evaluated to into a cell value.
/// </summary>
/// <remarks>
/// Preserving type here is the whole point of evaluating expressions to objects rather than to
/// strings: a decimal total has to land in the cell as a number so that Excel can sum, chart and
/// format it. Anything with no natural cell representation falls back to its text form.
/// </remarks>
internal static class ReportValueConverter
{
    /// <summary>
    /// Converts <paramref name="value"/> to a cell value, formatting unrecognised types as text
    /// under <paramref name="formatProvider"/>.
    /// </summary>
    public static XLCellValue ToCellValue(object? value, IFormatProvider? formatProvider = null)
    {
        formatProvider ??= CultureInfo.InvariantCulture;

        return value switch
        {
            null => Blank.Value,
            XLCellValue cellValue => cellValue,
            Blank blank => blank,
            XLError error => error,

            string text => text,
            char character => character.ToString(),
            bool logical => logical,

            DateTime dateTime => dateTime,
            DateTimeOffset dateTimeOffset => dateTimeOffset.DateTime,
            TimeSpan timeSpan => timeSpan,

            double number => number,
            float number => number,
            decimal number => number,
            sbyte number => number,
            byte number => number,
            short number => number,
            ushort number => number,
            int number => number,
            uint number => number,
            long number => number,
            ulong number => number,
            BigInteger number => (double)number,

            // An enum's name reads better in a report than its underlying number.
            Enum enumeration => enumeration.ToString(),

            IFormattable formattable => formattable.ToString(null, formatProvider),
            _ => value.ToString() ?? string.Empty,
        };
    }
}
