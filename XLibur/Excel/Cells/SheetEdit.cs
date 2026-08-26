using System;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel;

/// <summary>
/// One structural edit, as the shifter sees it. Passed to every <see cref="ISheetListener"/>.
/// </summary>
/// <remarks>
/// <para>
/// <see cref="Area"/> alone is not sufficient: for an insert it is the edited range extended by
/// <c>Shift - 1</c> lines along the shift axis (<see cref="XLWorksheetRangeShifter"/>), so the
/// shift magnitude cannot be recovered from it when the edited range spans more than one line on
/// that axis. <c>SheetEditAreaTests</c> pins the arithmetic: <c>A1:A5</c> with a shift of 3 yields
/// an area seven rows tall, not three. Listeners that only need the area — the calc engine and
/// hyperlinks — read <see cref="Area"/> and ignore the rest.
/// </para>
/// <para>
/// A <c>readonly struct</c> passed <c>in</c>, so it is not copied once per listener per edit, and so
/// that this port cannot become the allocation spec 21 spent its budget removing.
/// </para>
/// </remarks>
internal readonly struct SheetEdit
{
    /// <summary>The sheet the edit happened on. Not necessarily the listener's own sheet.</summary>
    internal required XLWorksheet Sheet { get; init; }

    /// <summary>The area inserted, or the area deleted.</summary>
    internal required Area Area { get; init; }

    /// <summary>The edited range, as the caller passed it. Not extended by <see cref="Shift"/>.</summary>
    internal required XLRange Range { get; init; }

    /// <summary>Lines shifted along the edit's axis. Positive for an insert, negative for a delete.</summary>
    internal required int Shift { get; init; }

    /// <summary>
    /// The region the edit inserts or deletes as coverage sees it: the edited range's extent on the
    /// cross axis, and <c>|Shift|</c> lines from its leading edge on the shift axis, mirroring
    /// <see cref="XLRangeInsertHelper"/>.
    /// </summary>
    /// <remarks>
    /// <b>This is deliberately not <see cref="Area"/>, and substituting one for the other is a
    /// defect.</b> For an insert <see cref="Area"/> extends the <em>whole</em> range by
    /// <c>Shift - 1</c>, so it is <c>Range.Height + Shift - 1</c> lines tall while this is
    /// <c>|Shift|</c> lines tall; the two agree only when the edited range is one line tall on the
    /// shift axis. For a delete they happen to agree, because the delete path always shifts by the
    /// range's own line count — but only by coincidence, and coverage must not depend on it. See
    /// <c>SheetEditAreaTests</c> and D15 in <c>DEFECTS.md</c>.
    /// <para>
    /// The area model handles every insert, including one at the first line — the old range-based
    /// path short-circuited there and let the blanket range shifter move the coverage instead.
    /// </para>
    /// </remarks>
    internal Area CoverageArea<TAxis>()
        where TAxis : struct, IGridAxis
    {
        var axis = default(TAxis);
        var first = Range.RangeAddress.FirstAddress;
        return new Area(
            Point.FromAddress(first),
            axis.PointAt(axis.IndexOf(first) + Math.Abs(Shift) - 1,
                axis.CrossOf(Range.RangeAddress.LastAddress)));
    }
}
