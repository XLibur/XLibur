using XLibur.Excel.Coordinates;

namespace XLibur.Excel;

/// <summary>
/// Hands a structural edit to every component that must react to it.
/// </summary>
/// <remarks>
/// <para>
/// This class names no feature. It builds one <see cref="SheetEdit"/> and walks
/// <see cref="XLWorksheet.GetSheetListeners"/>, which is the only place a sheet listener is named
/// (spec 33). Adding a sheet feature that must survive an insert or a delete is one
/// <see cref="ISheetListener"/> implementation and one <c>yield return</c> — nothing here changes.
/// </para>
/// <para>
/// It used to do the work itself, in six mirror pairs against three shared methods. Spec 26 bound
/// those to an <see cref="IGridAxis"/> through a generic type argument and collapsed each pair to
/// one method; spec 33 moved every one of them onto the feature that owns it. The axis argument
/// survives here only to choose which of the port's four members to call.
/// </para>
/// <para>
/// <b>The listener order is part of the contract</b> and is pinned by
/// <c>SheetListenerOrderTests</c>. It is not arbitrary: coverage is shifted from the pre-shift
/// address, so the merged-range split must precede it, and a rule whose coverage transforms to
/// nothing is deleted before its criteria formulas would be rewritten.
/// </para>
/// </remarks>
internal sealed class XLWorksheetRangeShifter(XLWorksheet worksheet)
{
    public void ShiftColumns(XLRange range, int columnsShifted) => Notify<ColumnAxis>(range, columnsShifted);

    public void ShiftRows(XLRange range, int rowsShifted) => Notify<RowAxis>(range, rowsShifted);

    /// <remarks>
    /// A shift of zero notifies nobody, which is what the hardcoded block this replaced did — it was
    /// an <c>if</c>/<c>else if</c> with no <c>else</c>. Zero is in any case unreachable: an insert
    /// rejects a count below one, and a delete always shifts by the deleted range's own line count,
    /// which a valid range cannot have zero of.
    /// </remarks>
    private void Notify<TAxis>(XLRange range, int shift)
        where TAxis : struct, IGridAxis
    {
        if (shift == 0)
            return;

        var axis = default(TAxis);
        var edit = new SheetEdit
        {
            Sheet = range.Worksheet,
            Area = shift > 0
                ? axis.ExtendAlongIndex(Area.FromRangeAddress(range.RangeAddress), shift - 1)
                : Area.FromRangeAddress(range.RangeAddress),
            Range = range,
            Shift = shift,
        };

        foreach (var listener in worksheet.GetSheetListeners())
        {
            if (shift > 0)
                axis.OnInsertAreaAndShift(listener, in edit);
            else
                axis.OnDeleteAreaAndShift(listener, in edit);
        }
    }
}
