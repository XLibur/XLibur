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
}
