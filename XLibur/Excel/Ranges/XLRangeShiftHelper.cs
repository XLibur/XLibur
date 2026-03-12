namespace XLibur.Excel;

/// <summary>
/// Contains the address-shifting algorithm logic for ranges.
/// <see cref="XLRangeBase"/> delegates the bodies of its protected
/// <c>ShiftColumns</c> and <c>ShiftRows</c> methods here.
/// </summary>
internal static class XLRangeShiftHelper
{
    internal static IXLRangeAddress ShiftColumns(
        XLWorksheet worksheet,
        XLRangeAddress currentRangeAddress,
        IXLRangeAddress thisRangeAddress,
        XLRange shiftedRange,
        int columnsShifted)
    {
        if (!thisRangeAddress.IsValid || !shiftedRange.RangeAddress.IsValid) return thisRangeAddress;

        var allRowsAreCovered = thisRangeAddress.FirstAddress.RowNumber >= shiftedRange.RangeAddress.FirstAddress.RowNumber &&
                                thisRangeAddress.LastAddress.RowNumber <= shiftedRange.RangeAddress.LastAddress.RowNumber;

        if (!allRowsAreCovered)
            return thisRangeAddress;

        var shiftLeftBoundary = (columnsShifted > 0 && thisRangeAddress.FirstAddress.ColumnNumber >= shiftedRange.RangeAddress.FirstAddress.ColumnNumber) ||
                                (columnsShifted < 0 && thisRangeAddress.FirstAddress.ColumnNumber > shiftedRange.RangeAddress.FirstAddress.ColumnNumber);

        var shiftRightBoundary = thisRangeAddress.LastAddress.ColumnNumber >= shiftedRange.RangeAddress.FirstAddress.ColumnNumber;

        var newLeftBoundary = thisRangeAddress.FirstAddress.ColumnNumber;
        if (shiftLeftBoundary)
        {
            if (newLeftBoundary + columnsShifted > shiftedRange.RangeAddress.FirstAddress.ColumnNumber)
                newLeftBoundary = newLeftBoundary + columnsShifted;
            else
                newLeftBoundary = shiftedRange.RangeAddress.FirstAddress.ColumnNumber;
        }

        var newRightBoundary = thisRangeAddress.LastAddress.ColumnNumber;
        if (shiftRightBoundary)
            newRightBoundary += columnsShifted;

        var destroyedByShift = newRightBoundary < newLeftBoundary;

        var firstAddress = (XLAddress)thisRangeAddress.FirstAddress;
        var lastAddress = (XLAddress)thisRangeAddress.LastAddress;

        if (destroyedByShift)
        {
            firstAddress = worksheet.InvalidAddress;
            lastAddress = worksheet.InvalidAddress;
            worksheet.DeleteRange(currentRangeAddress);
        }

        if (shiftLeftBoundary)
            firstAddress = new XLAddress(worksheet,
                thisRangeAddress.FirstAddress.RowNumber,
                newLeftBoundary,
                thisRangeAddress.FirstAddress.FixedRow,
                thisRangeAddress.FirstAddress.FixedColumn);

        if (shiftRightBoundary)
            lastAddress = new XLAddress(worksheet,
                thisRangeAddress.LastAddress.RowNumber,
                newRightBoundary,
                thisRangeAddress.LastAddress.FixedRow,
                thisRangeAddress.LastAddress.FixedColumn);

        return new XLRangeAddress(firstAddress, lastAddress);
    }

    internal static IXLRangeAddress ShiftRows(
        XLWorksheet worksheet,
        XLRangeAddress currentRangeAddress,
        IXLRangeAddress thisRangeAddress,
        XLRange shiftedRange,
        int rowsShifted)
    {
        if (!thisRangeAddress.IsValid || !shiftedRange.RangeAddress.IsValid) return thisRangeAddress;

        var allColumnsAreCovered = thisRangeAddress.FirstAddress.ColumnNumber >= shiftedRange.RangeAddress.FirstAddress.ColumnNumber &&
                                   thisRangeAddress.LastAddress.ColumnNumber <= shiftedRange.RangeAddress.LastAddress.ColumnNumber;

        if (!allColumnsAreCovered)
            return thisRangeAddress;

        var shiftTopBoundary = (rowsShifted > 0 && thisRangeAddress.FirstAddress.RowNumber >= shiftedRange.RangeAddress.FirstAddress.RowNumber) ||
                               (rowsShifted < 0 && thisRangeAddress.FirstAddress.RowNumber > shiftedRange.RangeAddress.FirstAddress.RowNumber);

        var shiftBottomBoundary = thisRangeAddress.LastAddress.RowNumber >= shiftedRange.RangeAddress.FirstAddress.RowNumber;

        var newTopBoundary = thisRangeAddress.FirstAddress.RowNumber;
        if (shiftTopBoundary)
        {
            if (newTopBoundary + rowsShifted > shiftedRange.RangeAddress.FirstAddress.RowNumber)
                newTopBoundary = newTopBoundary + rowsShifted;
            else
                newTopBoundary = shiftedRange.RangeAddress.FirstAddress.RowNumber;
        }

        var newBottomBoundary = thisRangeAddress.LastAddress.RowNumber;
        if (shiftBottomBoundary)
            newBottomBoundary += rowsShifted;

        var destroyedByShift = newBottomBoundary < newTopBoundary;

        var firstAddress = (XLAddress)thisRangeAddress.FirstAddress;
        var lastAddress = (XLAddress)thisRangeAddress.LastAddress;

        if (destroyedByShift)
        {
            firstAddress = worksheet.InvalidAddress;
            lastAddress = worksheet.InvalidAddress;
            worksheet.DeleteRange(currentRangeAddress);
        }

        if (shiftTopBoundary)
            firstAddress = new XLAddress(worksheet,
                newTopBoundary,
                thisRangeAddress.FirstAddress.ColumnNumber,
                thisRangeAddress.FirstAddress.FixedRow,
                thisRangeAddress.FirstAddress.FixedColumn);

        if (shiftBottomBoundary)
            lastAddress = new XLAddress(worksheet,
                newBottomBoundary,
                thisRangeAddress.LastAddress.ColumnNumber,
                thisRangeAddress.LastAddress.FixedRow,
                thisRangeAddress.LastAddress.FixedColumn);

        return new XLRangeAddress(firstAddress, lastAddress);
    }
}
