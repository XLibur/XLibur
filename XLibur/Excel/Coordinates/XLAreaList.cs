using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace XLibur.Excel.Coordinates;

/// <summary>
/// An immutable, value-typed list of rectangular sheet areas (the model behind an OOXML
/// <c>sqref</c>). Structural operations — <see cref="InsertAndShiftDown"/>,
/// <see cref="InsertAndShiftRight"/>, <see cref="DeleteAndShiftUp"/>,
/// <see cref="DeleteAndShiftLeft"/> — return a new list, transformed by pure functions on
/// <see cref="XLSheetRange"/>. Because areas are plain structs (not shared, repository-backed
/// range objects), a shift can never alias or double-apply — the failure mode behind
/// ClosedXML issue #2850. Intended to back conditional-format and data-validation coverage.
/// </summary>
internal sealed class XLAreaList : IEnumerable<XLSheetRange>
{
    internal static readonly XLAreaList Empty = new(new List<XLSheetRange>());

    private readonly List<XLSheetRange> _areas;

    internal XLAreaList(XLSheetRange area)
    {
        _areas = new List<XLSheetRange>(1) { area };
    }

    internal XLAreaList(List<XLSheetRange> areas)
    {
        _areas = areas;
    }

    internal int Count => _areas.Count;

    internal XLSheetRange this[int idx] => _areas[idx];

    internal static XLAreaList FromRanges(IEnumerable<IXLRange> ranges)
    {
        var areas = new List<XLSheetRange>();
        foreach (var range in ranges)
            areas.Add(XLSheetRange.FromRangeAddress(range.RangeAddress));

        return new XLAreaList(areas);
    }

    /// <summary>
    /// Return a new list with an additional area appended.
    /// </summary>
    internal XLAreaList With(XLSheetRange area)
    {
        return new XLAreaList(new List<XLSheetRange>(_areas) { area });
    }

    /// <summary>
    /// Return a new list with the first occurrence of <paramref name="area"/> removed.
    /// </summary>
    internal XLAreaList Without(XLSheetRange area)
    {
        var newList = new List<XLSheetRange>(_areas);
        newList.Remove(area);
        return new XLAreaList(newList);
    }

    internal XLAreaList InsertAndShiftDown(XLSheetRange insertedArea)
    {
        // Method is not symmetrical with InsertAndShiftRight, because Excel doesn't produce
        // symmetrical results (e.g. original C3:E5 and insert down at C3 produces asymmetrical
        // results from insert right at E3).
        var result = new List<XLSheetRange>(_areas.Count);
        foreach (var originalArea in _areas)
        {
            if (originalArea.HasFullColumnHeight)
            {
                result.Add(originalArea);
                continue;
            }

            // Skip all cases that don't shift or extend the area in some way.
            if (insertedArea.RightColumn < originalArea.LeftColumn ||
                insertedArea.LeftColumn > originalArea.RightColumn ||
                insertedArea.TopRow > originalArea.BottomRow + 1)
            {
                result.Add(originalArea);
                continue;
            }

            if (originalArea.SplitAbove(insertedArea.TopRow, out var above, out var remaining) &&
                above.Value.LeftColumn >= insertedArea.LeftColumn &&
                above.Value.RightColumn <= insertedArea.RightColumn)
            {
                // Special case: if inserted area covers the full width of the original area and
                // there is something above, the whole area is just extended downwards. The null
                // check handles the inserted area attaching to the bottom of the original.
                var mergedAndExtended = above.Value.ExtendBelow(insertedArea.Height + (remaining?.Height ?? 0));
                result.Add(mergedAndExtended);
                continue;
            }

            XLSheetRange? left = null, right = null;
            if (remaining is not null)
                remaining.Value.SplitBefore(insertedArea.LeftColumn, out left, out remaining);

            if (remaining is not null)
                remaining.Value.SplitAfter(insertedArea.RightColumn, out right, out remaining);

            if (above is not null)
                result.Add(above.Value);

            if (left is not null)
                result.Add(left.Value);

            if (right is not null)
                result.Add(right.Value);

            if (above is not null)
            {
                // There was something above the inserted area, so extend.
                if (remaining is not null)
                {
                    var extended = remaining.Value.ExtendBelow(insertedArea.Height);
                    result.Add(extended);
                }
                else if (insertedArea.TopRow == originalArea.BottomRow + 1)
                {
                    // Partial cover attaching at the bottom of the original area, e.g. insert to
                    // B2 with original A1:C1.
                    var cutToWidth = new XLSheetRange(
                        insertedArea.TopRow,
                        Math.Max(insertedArea.LeftColumn, originalArea.LeftColumn),
                        insertedArea.BottomRow,
                        Math.Min(insertedArea.RightColumn, originalArea.RightColumn));
                    result.Add(cutToWidth);
                }
            }
            else
            {
                // There was nothing above the inserted area, so shift.
                if (remaining is null)
                    throw new UnreachableException();

                if (remaining.Value.ShiftRowsAndClip(insertedArea.Height) is { } shifted)
                    result.Add(shifted);
            }
        }

        return new XLAreaList(result);
    }

    internal XLAreaList InsertAndShiftRight(XLSheetRange insertedArea)
    {
        var result = new List<XLSheetRange>(_areas.Count);
        foreach (var originalArea in _areas)
        {
            if (originalArea.HasFullRowWidth)
            {
                result.Add(originalArea);
                continue;
            }

            // Skip all cases that don't shift or extend the area in some way.
            if (insertedArea.BottomRow < originalArea.TopRow ||
                insertedArea.TopRow > originalArea.BottomRow ||
                insertedArea.LeftColumn > originalArea.RightColumn + 1)
            {
                result.Add(originalArea);
                continue;
            }

            // Deal with the special case of attachment at the right side.
            if (insertedArea.LeftColumn == originalArea.RightColumn + 1)
            {
                if (originalArea.TopRow >= insertedArea.TopRow &&
                    originalArea.BottomRow <= insertedArea.BottomRow)
                {
                    result.Add(originalArea.ExtendRight(insertedArea.Width));
                }
                else
                {
                    // Attaches at the right of the original area, e.g. insert to B2 with original A1:C1.
                    var cutToHeight = new XLSheetRange(
                        Math.Max(insertedArea.TopRow, originalArea.TopRow),
                        insertedArea.LeftColumn,
                        Math.Min(insertedArea.BottomRow, originalArea.BottomRow),
                        insertedArea.RightColumn);
                    result.Add(originalArea);
                    result.Add(cutToHeight);
                }

                continue;
            }

            XLSheetRange? below = null, left = null;
            originalArea.SplitAbove(insertedArea.TopRow, out var above, out var remaining);

            if (remaining is not null)
                remaining.Value.SplitBelow(insertedArea.BottomRow, out below, out remaining);

            if (remaining is not null)
                remaining.Value.SplitBefore(insertedArea.LeftColumn, out left, out remaining);

            // Something must remain: the inserted area intersects the original area (the right-side
            // attachment special case is handled above) and we only cut three times, one per side.
            if (remaining is null)
                throw new UnreachableException();

            if (above is not null)
                result.Add(above.Value);

            if (below is not null)
                result.Add(below.Value);

            if (left is not null)
            {
                // There was something on the left of the inserted area, so extend.
                var mergedAndExtended = left.Value.ExtendRight(insertedArea.Width + remaining.Value.Width);
                result.Add(mergedAndExtended);
            }
            else
            {
                // There is nothing on the left side, so shift.
                if (remaining.Value.ShiftColumnsAndClip(insertedArea.Width) is { } shifted)
                    result.Add(shifted);
            }
        }

        return new XLAreaList(result);
    }

    internal XLAreaList DeleteAndShiftUp(XLSheetRange deletedArea)
    {
        var groove = deletedArea.ExtendBelow(XLHelper.MaxRowNumber);
        var result = new List<XLSheetRange>(_areas.Count);
        foreach (var originalArea in _areas)
        {
            if (originalArea.HasFullColumnHeight)
            {
                result.Add(originalArea);
                continue;
            }

            var deleteWontSplitOriginalArea =
                deletedArea.LeftColumn <= originalArea.LeftColumn && deletedArea.RightColumn >= originalArea.RightColumn;
            if (deleteWontSplitOriginalArea)
            {
                var shiftedArea = originalArea.ShiftOrShrinkUp(deletedArea.TopRow, deletedArea.Height);
                if (shiftedArea is not null)
                    result.Add(shiftedArea.Value);
            }
            else
            {
                var inGrooveArea = originalArea.Exclude(groove, result);
                if (inGrooveArea is not null)
                {
                    // There is something to shift, so shift it upwards.
                    var shiftedArea = inGrooveArea.Value.ShiftOrShrinkUp(deletedArea.TopRow, deletedArea.Height);
                    if (shiftedArea is not null)
                        result.Add(shiftedArea.Value);
                }
            }
        }

        return new XLAreaList(result);
    }

    internal XLAreaList DeleteAndShiftLeft(XLSheetRange deletedArea)
    {
        var groove = deletedArea.ExtendRight(XLHelper.MaxColumnNumber);
        var result = new List<XLSheetRange>(_areas.Count);
        foreach (var originalArea in _areas)
        {
            if (originalArea.HasFullRowWidth)
            {
                result.Add(originalArea);
                continue;
            }

            var deleteWontSplitOriginalArea =
                deletedArea.TopRow <= originalArea.TopRow && deletedArea.BottomRow >= originalArea.BottomRow;
            if (deleteWontSplitOriginalArea)
            {
                var shiftedArea = originalArea.ShiftOrShrinkLeft(deletedArea.LeftColumn, deletedArea.Width);
                if (shiftedArea is not null)
                    result.Add(shiftedArea.Value);
            }
            else
            {
                var inGrooveArea = originalArea.Exclude(groove, result);
                if (inGrooveArea is not null)
                {
                    // There is something to shift, so shift it leftward.
                    var shiftedArea = inGrooveArea.Value.ShiftOrShrinkLeft(deletedArea.LeftColumn, deletedArea.Width);
                    if (shiftedArea is not null)
                        result.Add(shiftedArea.Value);
                }
            }
        }

        return new XLAreaList(result);
    }

    internal XLAreaList DeleteWithoutShift(XLSheetRange deletedArea)
    {
        var result = new List<XLSheetRange>(_areas.Count);
        foreach (var originalArea in _areas)
            originalArea.Exclude(deletedArea, result);

        return new XLAreaList(result);
    }

    internal bool IntersectsWith(XLSheetRange otherArea)
    {
        foreach (var area in _areas)
        {
            if (area.Intersects(otherArea))
                return true;
        }

        return false;
    }

    /// <summary>
    /// Return the areas in the list (at their original size) intersecting <paramref name="otherArea"/>.
    /// </summary>
    internal IEnumerable<XLSheetRange> IntersectingWith(XLSheetRange otherArea)
    {
        foreach (var area in _areas)
        {
            if (area.Intersects(otherArea))
                yield return area;
        }
    }

    /// <summary>
    /// Take the areas, intersect them with <paramref name="areaToCopy"/> and shift the pieces to
    /// <paramref name="target"/>. Used mostly in copy&amp;paste.
    /// </summary>
    internal bool TryCopyAreaTo(XLSheetPoint target, XLSheetRange areaToCopy, [NotNullWhen(true)] out XLAreaList? result)
    {
        var rowShift = target.Row - areaToCopy.FirstPoint.Row;
        var columnShift = target.Column - areaToCopy.FirstPoint.Column;
        List<XLSheetRange>? copyList = null;
        foreach (var area in _areas)
        {
            if (area.Intersect(areaToCopy) is not { } intersection)
                continue;

            // The end can be cut off, but the area always has at least 1x1 so it stays valid.
            if (intersection.ShiftAndClip(rowShift, columnShift) is not { } shiftedArea)
                continue;

            copyList ??= new List<XLSheetRange>();
            copyList.Add(shiftedArea);
        }

        if (copyList is not null)
        {
            result = new XLAreaList(copyList);
            return true;
        }

        result = null;
        return false;
    }

    /// <summary>
    /// Return a new list with <paramref name="excludedArea"/> cut out of every area.
    /// </summary>
    internal XLAreaList Excluding(XLSheetRange excludedArea)
    {
        if (!IntersectsWith(excludedArea))
            return this;

        var list = new List<XLSheetRange>();
        foreach (var area in _areas)
            area.Exclude(excludedArea, list);

        return new XLAreaList(list);
    }

    public IEnumerator<XLSheetRange> GetEnumerator()
    {
        return _areas.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    /// <summary>
    /// Render the areas as a space-separated <c>sqref</c> string (e.g. <c>"A1:B2 D4"</c>).
    /// </summary>
    internal string ToSpaceList()
    {
        return string.Join(" ", _areas.Select(a => a.ToString()));
    }
}
