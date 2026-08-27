using System.Linq;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel;

/// <summary>
/// Unmerges the merged ranges a structural edit would tear.
/// </summary>
/// <remarks>
/// Merged ranges are the one converted feature whose owner is a general-purpose type — they live in
/// a plain <see cref="XLRanges"/> on <c>Worksheet.Internals</c>, shared with every other range
/// collection — so the behaviour is carried by a dedicated adapter rather than by the collection.
/// The body moved here verbatim from <c>XLWorksheetRangeShifter.SplitMergedRangesCrossingTheShift</c>.
/// <para>
/// A merged range extending past the edited range on the cross axis cannot survive the shift with
/// its shape intact, so it is unmerged rather than moved. An entire-line edited range cannot tear
/// anything, so it is left alone. The transform reads only the edited range, not the shift, which is
/// why insert and delete run the same code.
/// </para>
/// </remarks>
internal sealed class MergedRangeSplitListener(XLWorksheet worksheet) : ISheetListener
{
    void ISheetListener.OnInsertAreaAndShiftDown(in SheetEdit edit) => Split<RowAxis>(in edit);

    void ISheetListener.OnInsertAreaAndShiftRight(in SheetEdit edit) => Split<ColumnAxis>(in edit);

    void ISheetListener.OnDeleteAreaAndShiftUp(in SheetEdit edit) => Split<RowAxis>(in edit);

    void ISheetListener.OnDeleteAreaAndShiftLeft(in SheetEdit edit) => Split<ColumnAxis>(in edit);

    private void Split<TAxis>(in SheetEdit edit)
        where TAxis : struct, IGridAxis
    {
        if (edit.Sheet != worksheet)
            return;

        var axis = default(TAxis);
        var range = edit.Range;
        if (axis.IsEntireLine(range))
            return;

        var first = range.RangeAddress.FirstAddress;
        var last = range.RangeAddress.LastAddress;
        var model = new XLRangeAddress((XLAddress)first, axis.AddressAtMaxIndex(last));
        var rangesToSplit = worksheet.MergedRanges
            .GetIntersectedRanges(model)
            .Where(r => axis.CrossOf(r.RangeAddress.FirstAddress) < axis.CrossOf(first) ||
                        axis.CrossOf(r.RangeAddress.LastAddress) > axis.CrossOf(last))
            .ToList();
        foreach (var rangeToSplit in rangesToSplit)
            worksheet.MergedRanges.Remove(rangeToSplit);
    }
}
