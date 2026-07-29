using System;
using System.Collections.Generic;
using System.Globalization;

namespace XLibur.Report.Tags;

/// <summary>
/// Orders the values an expression produced, keeping blanks together at the end.
/// </summary>
/// <remarks>
/// Sorting and grouping both order rows by an evaluated expression, and neither knows the type in
/// advance — the same key is a decimal in one template and a string in the next. Values of the same
/// comparable type compare directly; numbers of different types compare numerically rather than
/// textually, so <c>9</c> does not sort after <c>10</c>; anything else falls back to its text form.
/// </remarks>
internal sealed class SortKeyComparer : IComparer<object?>
{
    public static readonly SortKeyComparer Instance = new();

    public int Compare(object? x, object? y)
    {
        if (x is null)
        {
            return y is null ? 0 : 1;
        }

        if (y is null)
        {
            return -1;
        }

        if (x is IComparable comparable && x.GetType() == y.GetType())
        {
            return comparable.CompareTo(y);
        }

        // Different types: compare numerically when both are numbers, textually otherwise.
        if (IsNumeric(x) && IsNumeric(y))
        {
            return Convert.ToDouble(x, CultureInfo.InvariantCulture)
                .CompareTo(Convert.ToDouble(y, CultureInfo.InvariantCulture));
        }

        return string.Compare(x.ToString(), y.ToString(), StringComparison.CurrentCulture);
    }

    private static bool IsNumeric(object value) =>
        value is sbyte or byte or short or ushort or int or uint or long or ulong or float or double or decimal;
}
