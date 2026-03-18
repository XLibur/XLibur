using System;
using System.Globalization;
using System.Text;
using XLibur.Excel.CalcEngine;

namespace XLibur.Extensions;

internal static class TimeSpanExtensions
{
    extension(TimeSpan time)
    {
        public double ToSerialDateTime()
        {
            return time.Ticks / (double)TimeSpan.TicksPerDay;
        }

        /// <summary>
        /// Return a string representation of a TimeSpan that can be parsed by an Excel through text-to-number coercion.
        /// </summary>
        /// <remarks>
        /// Excel can convert a time span string back to a number, but only if it doesn't has days in the string, only hours.
        /// It's an opposite of <see cref="TimeSpanParser"/>.
        /// </remarks>
        public string ToExcelString(CultureInfo culture)
        {
            var timeSep = culture.DateTimeFormat.TimeSeparator;
            var sb = new StringBuilder()
                .Append(time.Hours + 24 * time.Days).Append(timeSep)
                .Append($"{time.Minutes:D2}").Append(timeSep)
                .Append($"{time.Seconds:D2}");
            // the ts.Miliseconds property uses whole division, and due to serial datetime conversion, it should be rounded instead
            var ms = (int)Math.Round((time.Ticks % TimeSpan.TicksPerSecond) * 1000.0 / (TimeSpan.TicksPerSecond));
            if (ms != 0)
            {
                sb.Append(culture.NumberFormat.CurrencyDecimalSeparator);
                if (ms % 100 == 0)
                    sb.Append($"{ms / 100:0}");
                else if (ms % 10 == 0)
                    sb.Append($"{ms / 10:00}");
                else
                    sb.Append($"{ms:000}");
            }
            return sb.ToString();
        }
    }
}
