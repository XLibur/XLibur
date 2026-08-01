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
/// <see cref="Area"/>. Because areas are plain structs (not shared, repository-backed
/// range objects), a shift can never alias or double-apply — the failure mode behind
/// ClosedXML issue #2850. Intended to back conditional-format and data-validation coverage.
/// </summary>
internal sealed class XLAreaList : IEnumerable<Area>
{
    internal static readonly XLAreaList Empty = new(new List<Area>());

    private readonly List<Area> _areas;

    internal XLAreaList(Area area)
    {
        _areas = new List<Area>(1) { area };
    }

    internal XLAreaList(List<Area> areas)
    {
        _areas = areas;
    }

    internal int Count => _areas.Count;

    internal Area this[int idx] => _areas[idx];

    internal static XLAreaList FromRanges(IEnumerable<IXLRange> ranges)
    {
        var areas = new List<Area>();
        foreach (var range in ranges)
            areas.Add(Area.FromRangeAddress(range.RangeAddress));

        return new XLAreaList(areas);
    }

    /// <summary>
    /// Return a new list with an additional area appended.
    /// </summary>
    internal XLAreaList With(Area area)
    {
        return new XLAreaList(new List<Area>(_areas) { area });
    }

    /// <summary>
    /// Return a new list with the first occurrence of <paramref name="area"/> removed.
    /// </summary>
    internal XLAreaList Without(Area area)
    {
        var newList = new List<Area>(_areas);
        newList.Remove(area);
        return new XLAreaList(newList);
    }

    internal XLAreaList InsertAndShiftDown(Area insertedArea)
    {
        // Method is not symmetrical with InsertAndShiftRight, because Excel doesn't produce
        // symmetrical results (e.g. original C3:E5 and insert down at C3 produces asymmetrical
        // results from insert right at E3).
        var result = new List<Area>(_areas.Count);
        foreach (var originalArea in _areas)
            AddShiftedDown(originalArea, insertedArea, result);

        return new XLAreaList(result);
    }

    /// <summary>
    /// Appends the pieces one area breaks into when <paramref name="insertedArea"/> pushes it down.
    /// An area can survive whole, be extended, or be cut into an above/left/right/shifted set.
    /// </summary>
#pragma warning disable S3776 // Splitting one area against an insertion; already reduced from 32
    private static void AddShiftedDown(Area originalArea, Area insertedArea, List<Area> result)
    {
        if (originalArea.HasFullColumnHeight)
        {
            result.Add(originalArea);
            return;
        }

        // Skip all cases that don't shift or extend the area in some way.
        if (insertedArea.RightColumn < originalArea.LeftColumn ||
            insertedArea.LeftColumn > originalArea.RightColumn ||
            insertedArea.TopRow > originalArea.BottomRow + 1)
        {
            result.Add(originalArea);
            return;
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
            return;
        }

        Area? left = null, right = null;
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

        if (above is null)
        {
            // There was nothing above the inserted area, so shift.
            if (remaining is null)
                throw new UnreachableException();

            if (remaining.Value.ShiftRowsAndClip(insertedArea.Height) is { } shifted)
                result.Add(shifted);

            return;
        }

        // There was something above the inserted area, so extend.
        if (remaining is not null)
        {
            result.Add(remaining.Value.ExtendBelow(insertedArea.Height));
        }
        else if (insertedArea.TopRow == originalArea.BottomRow + 1)
        {
            // Partial cover attaching at the bottom of the original area, e.g. insert to
            // B2 with original A1:C1.
            var cutToWidth = new Area(
                insertedArea.TopRow,
                Math.Max(insertedArea.LeftColumn, originalArea.LeftColumn),
                insertedArea.BottomRow,
                Math.Min(insertedArea.RightColumn, originalArea.RightColumn));
            result.Add(cutToWidth);
        }
    }
#pragma warning restore S3776

    internal XLAreaList InsertAndShiftRight(Area insertedArea)
    {
        var result = new List<Area>(_areas.Count);
        foreach (var originalArea in _areas)
            AddShiftedRight(originalArea, insertedArea, result);

        return new XLAreaList(result);
    }

    /// <summary>
    /// Appends the pieces one area breaks into when <paramref name="insertedArea"/> pushes it right.
    /// An area can survive whole, be extended, or be cut into an above/below/left/shifted set.
    /// </summary>
    private static void AddShiftedRight(Area originalArea, Area insertedArea, List<Area> result)
    {
        if (originalArea.HasFullRowWidth)
        {
            result.Add(originalArea);
            return;
        }

        // Skip all cases that don't shift or extend the area in some way.
        if (insertedArea.BottomRow < originalArea.TopRow ||
            insertedArea.TopRow > originalArea.BottomRow ||
            insertedArea.LeftColumn > originalArea.RightColumn + 1)
        {
            result.Add(originalArea);
            return;
        }

        // Deal with the special case of attachment at the right side.
        if (insertedArea.LeftColumn == originalArea.RightColumn + 1)
        {
            AddAttachedAtRight(originalArea, insertedArea, result);
            return;
        }

        Area? below = null, left = null;
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

    /// <summary>
    /// The inserted area starts exactly one column past the original, so nothing moves: the original
    /// either grows to swallow the insertion, or keeps it as a separate piece cut to their overlap.
    /// </summary>
    private static void AddAttachedAtRight(Area originalArea, Area insertedArea, List<Area> result)
    {
        if (originalArea.TopRow >= insertedArea.TopRow &&
            originalArea.BottomRow <= insertedArea.BottomRow)
        {
            result.Add(originalArea.ExtendRight(insertedArea.Width));
            return;
        }

        // Attaches at the right of the original area, e.g. insert to B2 with original A1:C1.
        var cutToHeight = new Area(
            Math.Max(insertedArea.TopRow, originalArea.TopRow),
            insertedArea.LeftColumn,
            Math.Min(insertedArea.BottomRow, originalArea.BottomRow),
            insertedArea.RightColumn);
        result.Add(originalArea);
        result.Add(cutToHeight);
    }

    internal XLAreaList DeleteAndShiftUp(Area deletedArea)
    {
        var groove = deletedArea.ExtendBelow(XLHelper.MaxRowNumber);
        var result = new List<Area>(_areas.Count);
        foreach (var originalArea in _areas)
            AddShiftedUp(originalArea, deletedArea, groove, result);

        return new XLAreaList(result);
    }

    /// <summary>
    /// Appends what survives of one area when <paramref name="deletedArea"/> pulls it up. When the
    /// deletion is narrower than the area, the parts outside <paramref name="groove"/> stay put and
    /// only the part inside it moves.
    /// </summary>
    private static void AddShiftedUp(Area originalArea, Area deletedArea, Area groove, List<Area> result)
    {
        if (originalArea.HasFullColumnHeight)
        {
            result.Add(originalArea);
            return;
        }

        var deleteWontSplitOriginalArea =
            deletedArea.LeftColumn <= originalArea.LeftColumn && deletedArea.RightColumn >= originalArea.RightColumn;

        // Exclude appends the pieces outside the groove to the result itself and hands back the one
        // piece inside it, which is the only part the deletion moves.
        var areaToShift = deleteWontSplitOriginalArea
            ? originalArea
            : originalArea.Exclude(groove, result);

        if (areaToShift?.ShiftOrShrinkUp(deletedArea.TopRow, deletedArea.Height) is { } shiftedArea)
            result.Add(shiftedArea);
    }

    internal XLAreaList DeleteAndShiftLeft(Area deletedArea)
    {
        var groove = deletedArea.ExtendRight(XLHelper.MaxColumnNumber);
        var result = new List<Area>(_areas.Count);
        foreach (var originalArea in _areas)
            AddShiftedLeft(originalArea, deletedArea, groove, result);

        return new XLAreaList(result);
    }

    /// <summary>
    /// Appends what survives of one area when <paramref name="deletedArea"/> pulls it left. When the
    /// deletion is shorter than the area, the parts outside <paramref name="groove"/> stay put and
    /// only the part inside it moves.
    /// </summary>
    private static void AddShiftedLeft(Area originalArea, Area deletedArea, Area groove, List<Area> result)
    {
        if (originalArea.HasFullRowWidth)
        {
            result.Add(originalArea);
            return;
        }

        var deleteWontSplitOriginalArea =
            deletedArea.TopRow <= originalArea.TopRow && deletedArea.BottomRow >= originalArea.BottomRow;

        // Exclude appends the pieces outside the groove to the result itself and hands back the one
        // piece inside it, which is the only part the deletion moves.
        var areaToShift = deleteWontSplitOriginalArea
            ? originalArea
            : originalArea.Exclude(groove, result);

        if (areaToShift?.ShiftOrShrinkLeft(deletedArea.LeftColumn, deletedArea.Width) is { } shiftedArea)
            result.Add(shiftedArea);
    }

    internal XLAreaList DeleteWithoutShift(Area deletedArea)
    {
        var result = new List<Area>(_areas.Count);
        foreach (var originalArea in _areas)
            originalArea.Exclude(deletedArea, result);

        return new XLAreaList(result);
    }

    /// <summary>
    /// Return an equivalent list with overlapping and adjacent areas merged into maximal blocks.
    /// </summary>
    internal XLAreaList GetConsolidated()
    {
        return XLAreaConsolidator.Consolidate(this);
    }

    internal bool IntersectsWith(Area otherArea)
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
    internal IEnumerable<Area> IntersectingWith(Area otherArea)
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
    internal bool TryCopyAreaTo(Point target, Area areaToCopy, [NotNullWhen(true)] out XLAreaList? result)
    {
        var rowShift = target.Row - areaToCopy.FirstPoint.Row;
        var columnShift = target.Column - areaToCopy.FirstPoint.Column;
        List<Area>? copyList = null;
        foreach (var area in _areas)
        {
            if (area.Intersect(areaToCopy) is not { } intersection)
                continue;

            // The end can be cut off, but the area always has at least 1x1 so it stays valid.
            if (intersection.ShiftAndClip(rowShift, columnShift) is not { } shiftedArea)
                continue;

            copyList ??= new List<Area>();
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
    internal XLAreaList Excluding(Area excludedArea)
    {
        if (!IntersectsWith(excludedArea))
            return this;

        var list = new List<Area>();
        foreach (var area in _areas)
            area.Exclude(excludedArea, list);

        return new XLAreaList(list);
    }

    public IEnumerator<Area> GetEnumerator()
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
