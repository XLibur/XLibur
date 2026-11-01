using System;

namespace XLibur.Excel.Coordinates;

/// <summary>
/// Where a line index, a line count or an area lands after a structural edit — for the features that
/// hold a raw position rather than a live range.
/// </summary>
/// <remarks>
/// <para>
/// This is not new arithmetic. It is <see cref="XLRangeShiftHelper"/>'s transform, reduced to the
/// integers, so that a feature holding a bare <c>int</c> or <see cref="Area"/> moves exactly as a
/// feature holding a range in the repository already does. That equivalence is the point: before
/// spec 33 a picture anchor got itself shifted by allocating a one-cell <c>IXLRange</c> purely so the
/// repository would move it (<c>XLMarker</c>'s own comment said so), and the picture was therefore
/// the one drawing anchor that worked. Charts, notes, panes and pivot tables now move through this
/// instead, and the picture moves through it too, so all five agree by construction.
/// </para>
/// <para>
/// The rule was read off the picture, case by case, against the unmodified tree — a two-cell picture
/// anchored <c>C4:J20</c>:
/// </para>
/// <list type="table">
///   <item><description>insert 3 rows at row 1 (above) → <c>C7:J23</c> — both corners move</description></item>
///   <item><description>insert 3 rows at row 10 (inside) → <c>C4:J23</c> — it grows</description></item>
///   <item><description>insert 3 rows at row 30 (below) → unchanged</description></item>
///   <item><description>delete rows 1–2 (above) → <c>C2:J18</c></description></item>
///   <item><description>delete rows 5–6 (inside) → <c>C4:J18</c> — it shrinks</description></item>
///   <item><description>delete rows 1–25 (covering it) → <c>C1:J1</c> — <b>clamped, not deleted</b></description></item>
///   <item><description>a partial insert at <c>B1:B5</c> → unchanged; a drawing outside the edited
///     columns does not move</description></item>
/// </list>
/// <para>
/// One expression reproduces every one of those. The single exception is the case where a delete
/// starts exactly on the leading edge: the repository path leaves the range's first address
/// <em>invalid</em> there, so reading the picture's <c>TopLeftCell</c> throws
/// <see cref="ArgumentOutOfRangeException"/> (recorded as D16 in <c>DEFECTS.md</c>). This clamps
/// instead, which is what the neighbouring case already does.
/// </para>
/// </remarks>
internal static class GridShift
{
    /// <summary>
    /// Where the line at <paramref name="index"/> lands. A line before the edit does not move; a
    /// line at or after it moves by <paramref name="shift"/>, but never past the edit's leading
    /// line — a delete that swallows the line leaves it sitting on the deletion point.
    /// </summary>
    internal static int MoveIndex(int index, int editFirstIndex, int shift)
        => index >= editFirstIndex ? Math.Max(index + shift, editFirstIndex) : index;

    /// <summary>
    /// Where a <em>count</em> of lines from the top or left lands — a frozen pane's
    /// <c>SplitRow</c>/<c>SplitColumn</c>, which name how many lines are frozen rather than which
    /// line. The floor is one lower than <see cref="MoveIndex"/>'s, because the count of lines
    /// surviving above the deletion point is <c>editFirstIndex - 1</c>, and a count of zero is a
    /// pane that is gone.
    /// </summary>
    internal static int MoveCount(int count, int editFirstIndex, int shift)
        => count >= editFirstIndex ? Math.Max(count + shift, editFirstIndex - 1) : count;

    /// <summary>
    /// Where an area lands, both corners moved by <see cref="MoveIndex"/> along
    /// <typeparamref name="TAxis"/>. The area is left alone unless the edited range spans it
    /// completely on the cross axis, matching <see cref="XLRangeShiftHelper"/> — a column shift only
    /// repositions what its rows wholly cover.
    /// </summary>
    internal static Area MoveArea<TAxis>(Area area, XLRange editedRange, int shift)
        where TAxis : struct, IGridAxis
    {
        var axis = default(TAxis);
        var editFirst = editedRange.RangeAddress.FirstAddress;
        var editLast = editedRange.RangeAddress.LastAddress;

        if (axis.CrossOf(area.FirstPoint) < axis.CrossOf(editFirst) ||
            axis.CrossOf(area.LastPoint) > axis.CrossOf(editLast))
            return area;

        var editFirstIndex = axis.IndexOf(editFirst);
        var first = MoveIndex(axis.IndexOf(area.FirstPoint), editFirstIndex, shift);
        var last = Math.Max(MoveIndex(axis.IndexOf(area.LastPoint), editFirstIndex, shift), first);

        return new Area(
            axis.PointAt(first, axis.CrossOf(area.FirstPoint)),
            axis.PointAt(last, axis.CrossOf(area.LastPoint)));
    }
}
