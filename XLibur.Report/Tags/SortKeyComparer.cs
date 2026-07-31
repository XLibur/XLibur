using System;
using System.Collections.Generic;
using System.Globalization;

namespace XLibur.Report.Tags;

/// <summary>
/// Orders the values an expression produced under the report's culture, keeping blanks together at
/// the end.
/// </summary>
/// <remarks>
/// <para>
/// Sorting and grouping both order rows by an evaluated expression, and neither knows the type in
/// advance — the same key is a decimal in one template and a string in the next. Values of the same
/// comparable type compare directly; numbers of different types compare numerically rather than
/// textually, so <c>9</c> does not sort after <c>10</c>; anything else falls back to its text form.
/// </para>
/// <para>
/// Text collates under the culture the engine was given, never under the machine's. Collation is the
/// one part of ordering a locale genuinely changes — Swedish sorts <c>å</c> after <c>z</c>, Czech
/// treats <c>ch</c> as a letter of its own — so a report meant for those readers should honour it,
/// and every other report should get the same row order wherever it was generated. That is only true
/// if the answer comes from the template rather than from the current thread, which is why this
/// takes a culture instead of reading <see cref="CultureInfo.CurrentCulture"/>.
/// </para>
/// </remarks>
internal sealed class SortKeyComparer : IComparer<object?>
{
    private readonly CultureInfo _culture;
    private readonly CompareInfo _collation;

    public SortKeyComparer(CultureInfo culture)
    {
        _culture = culture ?? throw new ArgumentNullException(nameof(culture));
        _collation = _culture.CompareInfo;
    }

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

        // Strings are taken before the general comparable case, because string.CompareTo collates
        // under CultureInfo.CurrentCulture — the one culture this comparer must not consult.
        if (x is string left && y is string right)
        {
            return _collation.Compare(left, right);
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

        return _collation.Compare(KeyText.For(x, _culture), KeyText.For(y, _culture));
    }

    private static bool IsNumeric(object value) =>
        value is sbyte or byte or short or ushort or int or uint or long or ulong or float or double or decimal;
}
