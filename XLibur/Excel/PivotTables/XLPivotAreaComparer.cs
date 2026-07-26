using System;
using System.Collections.Generic;
using System.Linq;

namespace XLibur.Excel;

internal sealed class XLPivotAreaComparer : IEqualityComparer<XLPivotArea>
{
    private readonly XLPivotReferenceComparer _referenceComparer = new();

    public static readonly XLPivotAreaComparer Instance = new();

    public bool Equals(XLPivotArea? x, XLPivotArea? y)
    {
        if (ReferenceEquals(x, y))
            return true;

        if (x is null)
            return false;

        if (y is null)
            return false;

        return x.References.SequenceEqual(y.References, _referenceComparer) &&
               Nullable.Equals(x.Field, y.Field) &&
               x.Type == y.Type &&
               x.DataOnly == y.DataOnly &&
               x.LabelOnly == y.LabelOnly &&
               x.GrandRow == y.GrandRow &&
               x.GrandCol == y.GrandCol &&
               x.CacheIndex == y.CacheIndex &&
               x.Outline == y.Outline &&
               Nullable.Equals(x.Offset, y.Offset) &&
               x.CollapsedLevelsAreSubtotals == y.CollapsedLevelsAreSubtotals &&
               x.Axis == y.Axis &&
               x.FieldPosition == y.FieldPosition;
    }

    /// <summary>
    /// Whether two areas select the same region of the pivot table, ignoring how much of that region
    /// they take in: <see cref="XLPivotArea.DataOnly"/> and <see cref="XLPivotArea.LabelOnly"/> pick
    /// the data cells, the label cells or both, and <see cref="XLPivotArea.FieldPosition"/> is the
    /// position Excel records next to that choice. Used when reading a style, where a format that
    /// takes in more than the caller asked about still styles what the caller asked about. Writing a
    /// style has to identify one exact area and uses <see cref="Equals(XLPivotArea, XLPivotArea)"/>.
    /// </summary>
    public bool EqualsIgnoringScope(XLPivotArea? x, XLPivotArea? y)
    {
        if (ReferenceEquals(x, y))
            return true;

        if (x is null || y is null)
            return false;

        return x.References.SequenceEqual(y.References, _referenceComparer) &&
               Nullable.Equals(x.Field, y.Field) &&
               x.Type == y.Type &&
               x.GrandRow == y.GrandRow &&
               x.GrandCol == y.GrandCol &&
               x.CacheIndex == y.CacheIndex &&
               x.Outline == y.Outline &&
               Nullable.Equals(x.Offset, y.Offset) &&
               x.CollapsedLevelsAreSubtotals == y.CollapsedLevelsAreSubtotals &&
               x.Axis == y.Axis;
    }

    public int GetHashCode(XLPivotArea obj)
    {
        var hashCode = new HashCode();
        foreach (var reference in obj.References)
            hashCode.Add(reference, _referenceComparer);

        hashCode.Add(obj.Field);
        hashCode.Add(obj.Type);
        hashCode.Add(obj.DataOnly);
        hashCode.Add(obj.LabelOnly);
        hashCode.Add(obj.GrandRow);
        hashCode.Add(obj.GrandCol);
        hashCode.Add(obj.CacheIndex);
        hashCode.Add(obj.Outline);
        hashCode.Add(obj.Offset);
        hashCode.Add(obj.CollapsedLevelsAreSubtotals);
        hashCode.Add(obj.Axis);
        hashCode.Add(obj.FieldPosition);
        return hashCode.ToHashCode();
    }

    private sealed class XLPivotReferenceComparer : IEqualityComparer<XLPivotReference>
    {
        public bool Equals(XLPivotReference? x, XLPivotReference? y)
        {
            if (ReferenceEquals(x, y))
                return true;

            if (x is null)
                return false;

            if (y is null)
                return false;

            return x.FieldItems.SequenceEqual(y.FieldItems) &&
                   x.Field == y.Field &&
                   x.Selected == y.Selected &&
                   x.ByPosition == y.ByPosition &&
                   x.Relative == y.Relative &&
                   x.Subtotals.SetEquals(y.Subtotals);
        }

        public int GetHashCode(XLPivotReference obj)
        {
            var hashCode = new HashCode();
            foreach (var item in obj.FieldItems)
                hashCode.Add(item);

            hashCode.Add(obj.Field);
            hashCode.Add(obj.Selected);
            hashCode.Add(obj.ByPosition);
            hashCode.Add(obj.Relative);

            foreach (var subtotal in obj.Subtotals)
                hashCode.Add(subtotal);

            return hashCode.ToHashCode();
        }
    }
}
