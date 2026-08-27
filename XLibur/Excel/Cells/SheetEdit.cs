using System;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel;

/// <summary>
/// One structural edit, as the shifter sees it. Passed to every <see cref="ISheetListener"/>.
/// </summary>
/// <remarks>
/// <para>
/// <see cref="Area"/> is the region the edit moved the cells by: for an insert the edited range's
/// leading edge extended by <c>Shift - 1</c> lines along the shift axis, for a delete the edited
/// range itself (<see cref="XLWorksheetRangeShifter"/>). It carries no record of how far the edited
/// range reached on the shift axis, which is why <see cref="Range"/> is here as well — a listener
/// that must know what the caller edited, rather than what moved, cannot read that off the area.
/// <c>SheetEditAreaTests</c> pins the arithmetic. <see cref="XLHyperlinks"/> is the one listener
/// that needs only the area and ignores the rest.
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
    /// Derived from <see cref="Range"/> and <see cref="Shift"/> rather than read off
    /// <see cref="Area"/>, so coverage cannot inherit a mistake in how the shifter builds that area.
    /// It inherited one until D15 was fixed: <see cref="Area"/> then extended the <em>whole</em>
    /// range by <c>Shift - 1</c> for an insert, making it <c>Range.Height + Shift - 1</c> lines tall
    /// where this is <c>|Shift|</c>. The two now agree in both directions, and
    /// <c>SheetEditAreaTests</c> is what holds them together.
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
