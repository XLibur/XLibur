using System;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using XLibur.Extensions;

namespace XLibur.Excel;

internal static partial class XLCellValueConverter
{
    private static readonly Regex UtfPattern = UtfPatternGenerated();

    internal static bool TryConvert<T>(XLCellValue currentValue, out T value)
    {
        var targetType = typeof(T);
        var isNullable = targetType.IsNullableType();
        if (isNullable && currentValue.TryConvert(out Blank _))
        {
            value = default!;
            return true;
        }

        var underlyingType = targetType.GetUnderlyingType();
        if (underlyingType == typeof(DateTime) && currentValue.TryConvert(out DateTime dateTime))
        {
            value = (T)(object)dateTime;
            return true;
        }

        var culture = CultureInfo.CurrentCulture;
        if (underlyingType == typeof(TimeSpan) && currentValue.TryConvert(out TimeSpan timeSpan, culture))
        {
            value = (T)(object)timeSpan;
            return true;
        }

        if (underlyingType == typeof(bool) && currentValue.TryConvert(out bool boolean))
        {
            value = (T)(object)boolean;
            return true;
        }

        if (TryGetStringValue(out value, currentValue)) return true;

        if (underlyingType == typeof(XLError))
        {
            if (currentValue.IsError)
            {
                value = (T)(object)currentValue.GetError();
                return true;
            }

            return false;
        }

        if (underlyingType.IsEnum)
            return TryConvertEnum<T>(currentValue, underlyingType, culture, out value);

        var typeCode = Type.GetTypeCode(underlyingType);

        if (typeCode is >= TypeCode.Single and <= TypeCode.Decimal)
            return TryConvertFloatingPoint<T>(currentValue, typeCode, culture, out value);

        if (typeCode is >= TypeCode.SByte and <= TypeCode.UInt64)
            return TryConvertInteger<T>(currentValue, typeCode, culture, out value);

        return false;
    }

    private static bool TryConvertEnum<T>(XLCellValue currentValue, Type underlyingType, CultureInfo culture, out T value)
    {
        var strValue = currentValue.ToString(culture);
        if (Enum.IsDefined(underlyingType, strValue))
        {
            value = (T)Enum.Parse(underlyingType, strValue, ignoreCase: false);
            return true;
        }

        value = default!;
        return false;
    }

    private static bool TryConvertFloatingPoint<T>(XLCellValue currentValue, TypeCode typeCode, CultureInfo culture, out T value)
    {
        if (!currentValue.TryConvert(out double doubleValue, culture))
        {
            value = default!;
            return false;
        }

        if (typeCode == TypeCode.Single && doubleValue is < float.MinValue or > float.MaxValue)
        {
            value = default!;
            return false;
        }

        if (typeCode == TypeCode.Decimal && (double.IsNaN(doubleValue) || double.IsInfinity(doubleValue)
            || doubleValue < (double)decimal.MinValue || doubleValue > (double)decimal.MaxValue))
        {
            value = default!;
            return false;
        }

        value = typeCode switch
        {
            TypeCode.Single => (T)(object)(float)doubleValue,
            TypeCode.Double => (T)(object)doubleValue,
            TypeCode.Decimal => (T)(object)(decimal)doubleValue,
            _ => throw new NotSupportedException()
        };
        return true;
    }

    private static bool TryConvertInteger<T>(XLCellValue currentValue, TypeCode typeCode, CultureInfo culture, out T value)
    {
        if (!currentValue.TryConvert(out double doubleValue, culture))
        {
            value = default!;
            return false;
        }

        if (!doubleValue.Equals(Math.Truncate(doubleValue)))
        {
            value = default!;
            return false;
        }

        var valueIsWithinBounds = typeCode switch
        {
            TypeCode.SByte => doubleValue is >= sbyte.MinValue and <= sbyte.MaxValue,
            TypeCode.Byte => doubleValue is >= byte.MinValue and <= byte.MaxValue,
            TypeCode.Int16 => doubleValue is >= short.MinValue and <= short.MaxValue,
            TypeCode.UInt16 => doubleValue is >= ushort.MinValue and <= ushort.MaxValue,
            TypeCode.Int32 => doubleValue is >= int.MinValue and <= int.MaxValue,
            TypeCode.UInt32 => doubleValue is >= uint.MinValue and <= uint.MaxValue,
            TypeCode.Int64 => doubleValue is >= long.MinValue and <= long.MaxValue,
            TypeCode.UInt64 => doubleValue is >= ulong.MinValue and <= ulong.MaxValue,
            _ => throw new NotSupportedException()
        };
        if (!valueIsWithinBounds)
        {
            value = default!;
            return false;
        }

        try
        {
            value = typeCode switch
            {
                TypeCode.SByte => (T)(object)(sbyte)doubleValue,
                TypeCode.Byte => (T)(object)(byte)doubleValue,
                TypeCode.Int16 => (T)(object)(short)doubleValue,
                TypeCode.UInt16 => (T)(object)(ushort)doubleValue,
                TypeCode.Int32 => (T)(object)(int)doubleValue,
                TypeCode.UInt32 => (T)(object)(uint)doubleValue,
                TypeCode.Int64 => (T)(object)checked((long)doubleValue),
                TypeCode.UInt64 => (T)(object)checked((ulong)doubleValue),
                _ => throw new NotSupportedException()
            };
            return true;
        }
        catch (OverflowException)
        {
            value = default!;
            return false;
        }
    }

    private static bool TryGetStringValue<T>(out T value, XLCellValue currentValue)
    {
        if (typeof(T) == typeof(string))
        {
            var s = currentValue.ToString(CultureInfo.CurrentCulture);
            var matches = UtfPattern.Matches(s);

            if (matches.Count == 0)
            {
                value = (T)Convert.ChangeType(s, typeof(T));
                return true;
            }

            var sb = new StringBuilder();
            var lastIndex = 0;

            foreach (var match in matches.Cast<Match>())
            {
                var matchString = match.Value;
                var matchIndex = match.Index;
                sb.Append(s.Substring(lastIndex, matchIndex - lastIndex));

                sb.Append((char)int.Parse(match.Groups[1].Value, NumberStyles.AllowHexSpecifier));

                lastIndex = matchIndex + matchString.Length;
            }

            if (lastIndex < s.Length)
                sb.Append(s.Substring(lastIndex));

            value = (T)Convert.ChangeType(sb.ToString(), typeof(T));
            return true;
        }

        value = default!;
        return false;
    }

    [GeneratedRegex("(?<!_x005F)_x(?!005F)([0-9A-F]{4})_", RegexOptions.Compiled)]
    private static partial Regex UtfPatternGenerated();
}
