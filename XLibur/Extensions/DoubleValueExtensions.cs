#nullable disable


using DocumentFormat.OpenXml;
using System;

namespace XLibur.Excel;

internal static class DoubleValueExtensions
{
    public static DoubleValue SaveRound(this DoubleValue value)
    {
        return value.HasValue ? new DoubleValue(Math.Round(value, 6)) : value;
    }
}
