using System;
using XLibur.Excel;

namespace XLibur.Extensions;

internal static class DoubleExtensions
{
    extension(double value)
    {
        public double SaveRound()
        {
            return Math.Round(value, 6);
        }

        public TimeSpan ToSerialTimeSpan()
        {
            return XLHelper.GetTimeSpan(value);
        }

        public DateTime ToSerialDateTime()
        {
            return value switch
            {
                >= 61.0 => DateTime.FromOADate(value),
                <= 60.0 => DateTime.FromOADate(value + 1),
                _ => throw new ArgumentException(
                    "Serial date 60 is on a leap year of 1900 - date that doesn't exist and isn't representable in DateTime.")
            };
        }

        /// <summary>
        /// Round the number to the integer.
        /// </summary>
        /// <remarks>A helper method to avoid needs to specify the midpoint rounding and casting each time.</remarks>
        public int RoundToInt()
        {
            return (int)Math.Round(value, MidpointRounding.AwayFromZero);
        }

        /// <summary>
        /// Round the number to a specified number of digits.
        /// </summary>
        /// <remarks>A helper method to avoid the need to specify the midpoint rounding each time.</remarks>
        public double Round(int digits)
        {
            return Math.Round(value, digits, MidpointRounding.AwayFromZero);
        }
    }
}
