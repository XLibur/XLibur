using System;
using System.Linq;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel;

/// <summary>
/// Contains the heavy algorithmic logic for inserting rows and columns into a range.
/// <see cref="XLRangeBase"/> delegates to these methods to keep the main class smaller.
/// </summary>
/// <remarks>
/// The row and column halves used to be two 106-line copies of one algorithm. They are now one
/// implementation bound to an <see cref="IGridAxis"/> through a generic type argument, so the JIT
/// still specialises per axis (spec 26). The transposition the two copies hid is explicit here: the
/// shift runs on the index axis, the formatting pass runs on the <em>cross</em> axis.
/// </remarks>
internal static class XLRangeInsertHelper
{
    internal static IXLRangeColumns? InsertColumnsBefore(XLRangeBase range, bool onlyUsedCells, int numberOfColumns, bool formatFromLeft, bool nullReturn)
        => Insert<ColumnAxis>(range, onlyUsedCells, numberOfColumns, formatFromLeft, nullReturn, nameof(numberOfColumns))
            ?.Columns();

    internal static IXLRangeRows? InsertRowsAbove(XLRangeBase range, bool onlyUsedCells, int numberOfRows, bool formatFromAbove, bool nullReturn)
        => Insert<RowAxis>(range, onlyUsedCells, numberOfRows, formatFromAbove, nullReturn, nameof(numberOfRows))
            ?.Rows();

    private static XLRange? Insert<TAxis>(XLRangeBase range, bool onlyUsedCells, int count,
        bool formatFromPrevious, bool nullReturn, string countParamName)
        where TAxis : struct, IGridAxis
    {
        var axis = default(TAxis);
        EnsureInsertFits<TAxis>(range, count, countParamName);

        XLFormulaShiftPass.Run(range.Worksheet.Workbook, range.AsRange(), axis.ShiftsRows, count);

        axis.ShiftSparklines(range.Worksheet.SparklineGroupsInternal, Area.FromRangeAddress(range.RangeAddress), count);

        ShiftLineSizes<TAxis>(range, onlyUsedCells, count);

        var insertedRange = new Area(
            Point.FromAddress(range.RangeAddress.FirstAddress),
            axis.PointAt(
                axis.IndexOf(range.RangeAddress.FirstAddress) + count - 1,
                axis.CrossOf(range.RangeAddress.LastAddress)));

        axis.InsertAreaAndShift(range.Worksheet.Internals.CellsCollection, insertedRange);

        var firstIndexReturn = axis.IndexOf(range.RangeAddress.FirstAddress);
        var lastIndexReturn = firstIndexReturn + count - 1;
        var firstCrossReturn = axis.CrossOf(range.RangeAddress.FirstAddress);
        var lastCrossReturn = axis.CrossOf(range.RangeAddress.LastAddress);

        axis.NotifyRangeShifted(range.Worksheet, range.AsRange(), count);

        var rangeToReturn = axis.RangeFor(range.Worksheet, firstIndexReturn, lastIndexReturn, firstCrossReturn, lastCrossReturn);

        var contentFlags = XLCellsUsedOptions.All
                           & ~XLCellsUsedOptions.ConditionalFormats
                           & ~XLCellsUsedOptions.DataValidation;

        ApplyFormatting<TAxis>(range, rangeToReturn, formatFromPrevious, contentFlags);

        // Skip calling .Rows()/.Columns() for performance reasons if required.
        if (nullReturn)
            return null;

        return rangeToReturn;
    }

    /// <summary>
    /// What makes a cell non-empty for <see cref="EnsureInsertFits{TAxis}"/>. Excel
    /// (checked over COM) refuses for a value, a formula, a comment or a fill.
    /// </summary>
    /// <remarks>
    /// <see cref="XLCellsUsedOptions.NormalFormats"/> counts every format, so this is stricter than
    /// Excel in one case: Excel lets an insert push off a cell whose only format is a number format.
    /// </remarks>
    private const XLCellsUsedOptions PushedOffContent =
        XLCellsUsedOptions.AllContents | XLCellsUsedOptions.NormalFormats;

    /// <summary>
    /// Throws when <paramref name="count"/> lines cannot be inserted at <paramref name="range"/>:
    /// the count is out of range, the inserted lines would not fit on the sheet, or the insert
    /// would push a non-empty cell past the sheet's last row (or column). Excel refuses such an
    /// insert and changes nothing. The shift used to drop those cells, and a table it cut became
    /// <c>#REF!</c> and could not be saved (#773).
    /// </summary>
    /// <remarks>
    /// The insert runs this before it changes anything, so a refused insert leaves the workbook as
    /// it was. A caller that changes other state before the insert, as the worksheet row and column
    /// inserts do with their line collections, runs it first as well. Only the lines the range
    /// spans on the cross axis shift, so only those are checked: a cell insert at <c>A10:B10</c> is
    /// not blocked by a value in column C.
    /// </remarks>
    internal static void EnsureInsertFits<TAxis>(XLRangeBase range, int count, string countParamName)
        where TAxis : struct, IGridAxis
    {
        var axis = default(TAxis);
        if (count <= 0 || count > axis.MaxIndex)
            throw new ArgumentOutOfRangeException(countParamName,
                $"Number of {axis.LineNoun} to insert must be a positive number no more than {axis.MaxIndex}");

        var first = range.RangeAddress.FirstAddress;
        var last = range.RangeAddress.LastAddress;

        // The block must fit before the lines past it are checked: otherwise the lines that would
        // be pushed off start above the insert point, and those do not move at all.
        var firstIndex = axis.IndexOf(first);
        if (firstIndex + count - 1 > axis.MaxIndex)
            throw new ArgumentOutOfRangeException(countParamName,
                $"Cannot insert {count} {axis.LineNoun} starting at {firstIndex}: "
                + "they would run past the end of the worksheet.");

        var pushedOff = axis.RangeFor(range.Worksheet,
            axis.MaxIndex - count + 1, axis.MaxIndex, axis.CrossOf(first), axis.CrossOf(last));

        if (HasContentToPushOff(pushedOff))
            throw new InvalidOperationException(
                $"Cannot insert {count} {axis.LineNoun} because it would push non-empty cells "
                + "off the end of the worksheet.");
    }

    private static bool HasContentToPushOff(XLRange pushedOff)
    {
        if (pushedOff.CellsUsed(PushedOffContent).Any())
            return true;

        // An empty merged range, or a table's empty data rows, have no cells to find, but Excel
        // still refuses to push them off.
        var worksheet = pushedOff.Worksheet;
        if (worksheet.Internals.MergedRanges.GetIntersectedRanges(pushedOff.RangeAddress).Any())
            return true;

        foreach (var table in worksheet.Tables)
        {
            if (table.RangeAddress.Intersects(pushedOff.RangeAddress))
                return true;
        }

        return false;
    }

    /// <summary>Carries each line's height (rows) or width (columns) forward by <paramref name="count"/>.</summary>
    private static void ShiftLineSizes<TAxis>(XLRangeBase range, bool onlyUsedCells, int count)
        where TAxis : struct, IGridAxis
    {
        if (onlyUsedCells)
            return;

        var axis = default(TAxis);
        var lastIndex = axis.MaxUsedIndex(range.Worksheet.Internals.CellsCollection);
        if (lastIndex <= 0)
            return;

        // Both longhand copies tested this inside the loop, where it is invariant; hoisting it is a
        // behaviour-preserving improvement the collapse made obvious (spec 26 task 4).
        if (!axis.IsEntireLine(range))
            return;

        var firstIndex = axis.IndexOf(range.RangeAddress.FirstAddress);

        // A line whose target lands past the sheet's last row (or column) is pushed off the sheet,
        // so its size goes with it. Starting the walk below the edge drops those rather than
        // addressing row 1,048,577, which threw. Both longhand copies were missing the clamp.
        for (var i = Math.Min(lastIndex, axis.MaxIndex - count); i >= firstIndex; i--)
            axis.CopyLineSize(range.Worksheet, i, i + count);
    }

    private static void ApplyFormatting<TAxis>(XLRangeBase range, IXLRange rangeToReturn,
        bool formatFromPrevious, XLCellsUsedOptions contentFlags)
        where TAxis : struct, IGridAxis
    {
        var axis = default(TAxis);
        if (formatFromPrevious && axis.IndexOf(rangeToReturn.RangeAddress.FirstAddress) > 1)
            ApplyFormattingFromPreviousLine<TAxis>(rangeToReturn, contentFlags);
        else
            ApplyFormattingFromExistingCrossLines<TAxis>(range, rangeToReturn, contentFlags);
    }

    /// <summary>Styles the inserted block from the line before it — the column to its left, or the row
    /// above. The loop runs over the <em>cross</em> axis: a column insert styles rows, a row insert
    /// styles columns. That transposition is the most error-prone part of the mirror this replaces.</summary>
    private static void ApplyFormattingFromPreviousLine<TAxis>(IXLRange rangeToReturn, XLCellsUsedOptions contentFlags)
        where TAxis : struct, IGridAxis
    {
        var axis = default(TAxis);
        var model = axis.ModelLineBefore(rangeToReturn);
        var modelFirst = model.FirstCellUsed(contentFlags);
        var modelLast = model.LastCellUsed(contentFlags);
        if (modelFirst == null || modelLast == null)
            return;

        var modelFirstCross = axis.CrossOf(model.RangeAddress.FirstAddress);
        var firstCrossReturned = axis.CrossOf(modelFirst.Address) - modelFirstCross + 1;
        var lastCrossReturned = axis.CrossOf(modelLast.Address) - modelFirstCross + 1;
        for (var cross = firstCrossReturned; cross <= lastCrossReturned; cross++)
            axis.SetCrossLineStyle(rangeToReturn, cross, axis.ModelCellStyle(model, cross));
    }

    /// <summary>Styles the inserted block from the sheet's existing cross-axis lines, falling back to
    /// the worksheet style for a line that carries none.</summary>
    private static void ApplyFormattingFromExistingCrossLines<TAxis>(XLRangeBase range, IXLRange rangeToReturn, XLCellsUsedOptions contentFlags)
        where TAxis : struct, IGridAxis
    {
        var axis = default(TAxis);
        var lastUsedCross = axis.LastUsedCross(rangeToReturn, contentFlags);
        if (lastUsedCross < 0)
            return;

        var firstWsCross = axis.CrossOf(rangeToReturn.RangeAddress.FirstAddress);
        var lastCrossReturned = lastUsedCross - firstWsCross + 1;
        for (var cross = 1; cross <= lastCrossReturned; cross++)
            axis.SetCrossLineStyle(rangeToReturn, cross, axis.CrossLineStyle(range.Worksheet, firstWsCross + cross - 1));
    }
}
