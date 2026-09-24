using XLibur.Excel.Coordinates;

namespace XLibur.Excel;

/// <summary>Contains the address-shifting algorithm logic for ranges. <see cref="XLRangeBase"/>
/// delegates the bodies of its protected <c>ShiftColumns</c> and <c>ShiftRows</c> methods here.</summary>
internal static class XLRangeShiftHelper
{
    internal static IXLRangeAddress ShiftColumns(
        XLWorksheet worksheet,
        XLRangeAddress currentRangeAddress,
        IXLRangeAddress thisRangeAddress,
        XLRange shiftedRange,
        int columnsShifted)
        => Shift<ColumnAxis>(worksheet, currentRangeAddress, thisRangeAddress, shiftedRange, columnsShifted);

    internal static IXLRangeAddress ShiftRows(
        XLWorksheet worksheet,
        XLRangeAddress currentRangeAddress,
        IXLRangeAddress thisRangeAddress,
        XLRange shiftedRange,
        int rowsShifted)
        => Shift<RowAxis>(worksheet, currentRangeAddress, thisRangeAddress, shiftedRange, rowsShifted);

    /// <summary>Repositions one range's address for a shift on <typeparamref name="TAxis"/>. The edges
    /// are named for the axis, not a direction — "leading" is the left edge on the column axis and the
    /// top edge on the row axis — because the two copies this replaces were the same algorithm, and
    /// their <c>Left</c>/<c>Top</c> naming is why nobody noticed.</summary>
    private static IXLRangeAddress Shift<TAxis>(
        XLWorksheet worksheet,
        XLRangeAddress currentRangeAddress,
        IXLRangeAddress thisRangeAddress,
        XLRange shiftedRange,
        int shift)
        where TAxis : struct, IGridAxis
    {
        if (!thisRangeAddress.IsValid || !shiftedRange.RangeAddress.IsValid) return thisRangeAddress;

        var axis = default(TAxis);
        var thisFirst = thisRangeAddress.FirstAddress;
        var thisLast = thisRangeAddress.LastAddress;
        var shiftedFirst = shiftedRange.RangeAddress.FirstAddress;

        // The shifted range must span this one completely on the cross axis, or this range does not
        // move at all: a column shift only repositions ranges whose rows it wholly covers.
        if (!IsSpannedOnCrossAxis(axis, thisFirst, thisLast, shiftedFirst, shiftedRange.RangeAddress.LastAddress))
            return thisRangeAddress;

        var shiftedFirstIndex = axis.IndexOf(shiftedFirst);
        var thisFirstIndex = axis.IndexOf(thisFirst);
        var thisLastIndex = axis.IndexOf(thisLast);

        var leadingEdgeMoves = LeadingEdgeMoves(shift, thisFirstIndex, shiftedFirstIndex);
        var trailingEdgeMoves = thisLastIndex >= shiftedFirstIndex;

        var newLeadingEdge = leadingEdgeMoves
            ? ShiftLeadingEdge(thisFirstIndex, shift, shiftedFirstIndex)
            : thisFirstIndex;

        var newTrailingEdge = trailingEdgeMoves ? thisLastIndex + shift : thisLastIndex;

        var destroyedByShift = newTrailingEdge < newLeadingEdge;

        var firstAddress = (XLAddress)thisFirst;
        var lastAddress = (XLAddress)thisLast;

        if (destroyedByShift)
        {
            firstAddress = worksheet.InvalidAddress;
            lastAddress = worksheet.InvalidAddress;
            worksheet.DeleteRange(currentRangeAddress);
        }

        if (leadingEdgeMoves)
            firstAddress = axis.AddressAt(worksheet, newLeadingEdge, axis.CrossOf(thisFirst),
                thisFirst.FixedRow, thisFirst.FixedColumn);

        if (trailingEdgeMoves)
            lastAddress = axis.AddressAt(worksheet, newTrailingEdge, axis.CrossOf(thisLast),
                thisLast.FixedRow, thisLast.FixedColumn);

        return new XLRangeAddress(firstAddress, lastAddress);
    }

    private static bool IsSpannedOnCrossAxis<TAxis>(
        TAxis axis,
        IXLAddress thisFirst,
        IXLAddress thisLast,
        in XLAddress shiftedFirst,
        in XLAddress shiftedLast)
        where TAxis : struct, IGridAxis
        => axis.CrossOf(thisFirst) >= axis.CrossOf(shiftedFirst) &&
           axis.CrossOf(thisLast) <= axis.CrossOf(shiftedLast);

    /// <summary>The leading edge moves when an insert starts at or before it, or a delete starts strictly
    /// before it — a delete that starts exactly on the leading edge eats into the range instead.</summary>
    private static bool LeadingEdgeMoves(int shift, int thisFirstIndex, int shiftedFirstIndex)
        => (shift > 0 && thisFirstIndex >= shiftedFirstIndex) ||
           (shift < 0 && thisFirstIndex > shiftedFirstIndex);

    /// <summary>A delete cannot pull the leading edge back past the start of the deleted span.</summary>
    private static int ShiftLeadingEdge(int thisFirstIndex, int shift, int shiftedFirstIndex)
        => thisFirstIndex + shift > shiftedFirstIndex ? thisFirstIndex + shift : shiftedFirstIndex;
}
