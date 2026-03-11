
using System;

namespace XLibur.Excel;

internal static class DoubleExtensions
{
    public static double SaveRound(this double value)
    {
        return Math.Round(value, 6);
    }

    public static TimeSpan ToSerialTimeSpan(this double value)
    {
        return XLHelper.GetTimeSpan(value);
    }

    public static DateTime ToSerialDateTime(this double value)
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
    public static int RoundToInt(this double value)
    {
        return (int)Math.Round(value, MidpointRounding.AwayFromZero);
    }

    /// <summary>
    /// Round the number to a specified number of digits.
    /// </summary>
    /// <remarks>A helper method to avoid the need to specify the midpoint rounding each time.</remarks>
    public static double Round(this double value, int digits)
    {
        return Math.Round(value, digits, MidpointRounding.AwayFromZero);
    }
}
