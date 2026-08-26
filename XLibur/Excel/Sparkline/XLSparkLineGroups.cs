using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel.Coordinates;
using XLibur.Extensions;

namespace XLibur.Excel;

internal sealed class XLSparklineGroups : IXLSparklineGroups, ISheetListener
{
    private readonly XLWorksheet _worksheet;

    public XLSparklineGroups(XLWorksheet worksheet)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
    }

    public IXLWorksheet Worksheet => _worksheet;

    #region Public Methods

    /// <summary>
    /// Add empty sparkline group.
    /// </summary>
    internal IXLSparklineGroup Add()
    {
        return Add(new XLSparklineGroup(Worksheet));
    }

    /// <summary>
    /// Add the sparkline group to the collection.
    /// </summary>
    /// <param name="sparklineGroup">The sparkline group to add to the collection</param>
    /// <returns>The same sparkline group</returns>
    public IXLSparklineGroup Add(IXLSparklineGroup sparklineGroup)
    {
        if (sparklineGroup.Worksheet != Worksheet)
            throw new ArgumentException("The specified sparkline group belongs to the different worksheet");

        _sparklineGroups.Add(sparklineGroup);
        return sparklineGroup;
    }

    public IXLSparklineGroup Add(string locationAddress, string sourceDataAddress)
    {
        return Add(new XLSparklineGroup(Worksheet, locationAddress, sourceDataAddress));
    }

    public IXLSparklineGroup Add(IXLCell location, IXLRange sourceData)
    {
        return Add(new XLSparklineGroup(location, sourceData));
    }

    public IXLSparklineGroup Add(IXLRange locationRange, IXLRange sourceDataRange)
    {
        return Add(new XLSparklineGroup(locationRange, sourceDataRange));
    }

    /// <summary>
    /// Add a copy of an existing sparkline group to the specified worksheet
    /// </summary>
    /// <param name="sparklineGroupToCopy">The sparkline group to copy</param>
    /// <param name="targetWorksheet">The worksheet the sparkline group is being added to</param>
    /// <returns>The new sparkline group added</returns>
    public IXLSparklineGroup AddCopy(IXLSparklineGroup sparklineGroupToCopy, IXLWorksheet targetWorksheet)
    {
        var sparklineGroup = new XLSparklineGroup(targetWorksheet, sparklineGroupToCopy);
        _sparklineGroups.Add(sparklineGroup);
        return sparklineGroup;
    }

    /// <summary>
    /// Copy this sparkline group to a different worksheet
    /// </summary>
    /// <param name="targetSheet">The worksheet to copy the sparkline group to</param>
    public void CopyTo(IXLWorksheet targetSheet)
    {
        foreach (var slg in this)
        {
            slg.CopyTo(targetSheet);
        }
    }

    /// <summary>
    /// Search for the first sparkline that is in the specified cell
    /// </summary>
    /// <param name="cell">The cell to find the sparkline for</param>
    /// <returns>The sparkline in the cell or null if no sparklines are found</returns>
    public IXLSparkline GetSparkline(IXLCell cell)
    {
        return _sparklineGroups
            .Select(g => g.GetSparkline(cell))
            .FirstOrDefault(s => s != null)!;
    }

    /// <summary>
    /// Find all sparklines located in a given range
    /// </summary>
    /// <param name="rangeBase">The range to search</param>
    /// <returns>The sparkline in the cell or null if no sparklines are found</returns>
    public IEnumerable<IXLSparkline> GetSparklines(IXLRangeBase rangeBase)
    {
        return _sparklineGroups
            .SelectMany(g => g.GetSparklines(rangeBase));
    }

    public IEnumerator<IXLSparklineGroup> GetEnumerator()
    {
        return _sparklineGroups.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    /// <summary>
    /// Remove all sparklines in the specified cell
    /// </summary>
    /// <param name="cell">The cell to remove sparklines from</param>
    public void Remove(IXLCell cell)
    {
        _sparklineGroups
            .AsParallel()
            .ForEach(g => g.Remove(cell));
    }

    public void Remove(IXLRangeBase range)
    {
        var sparklinesToRemove = _sparklineGroups
            .SelectMany(g => g)
            .Where(sparkline => range.Contains(sparkline.Location))
            .ToList();

        sparklinesToRemove.ForEach(Remove);
    }

    /// <summary>
    /// Remove the sparkline group from the worksheet
    /// </summary>
    /// <param name="sparklineGroup">The sparkline group to remove</param>
    public void Remove(IXLSparklineGroup sparklineGroup)
    {
        _sparklineGroups.Remove(sparklineGroup);
    }

    /// <summary>
    /// Remove the sparkline from the worksheet
    /// </summary>
    /// <param name="sparkline">The sparkline to remove</param>
    private static void Remove(IXLSparkline sparkline)
    {
        sparkline.SparklineGroup.Remove(sparkline);
    }

    /// <summary>
    /// Remove all sparkline groups and their contents from the worksheet.
    /// </summary>
    public void RemoveAll()
    {
        _sparklineGroups.Clear();
    }

    #endregion Public Methods

    #region ISheetListener

    /// <summary>
    /// Drops the sparklines a structural edit pushed off the sheet.
    /// </summary>
    /// <remarks>
    /// <b>This is only half of what a sparkline needs, and the other half is elsewhere.</b> The
    /// sparklines themselves are moved by <see cref="ShiftRows"/> / <see cref="ShiftColumns"/>,
    /// called from <see cref="XLRangeInsertHelper"/> and <c>XLRangeBase.Delete</c> — a dispatch
    /// point upstream of <see cref="XLWorksheetRangeShifter"/>, before the cells have moved, and one
    /// this seam does not reach. So a sparkline group is the one sheet feature still notified twice
    /// from two different layers. Spec 33 left it that way deliberately, scoped out to a follow-on:
    /// folding it in means moving the call site, not writing an adapter.
    /// <para>
    /// What is here is the cleanup pass: an address the move pushed past the grid's edge is invalid,
    /// and the sparkline holding it is removed. The transform reads neither axis nor shift, which is
    /// why all four members run the same code.
    /// </para>
    /// </remarks>
    void ISheetListener.OnInsertAreaAndShiftDown(in SheetEdit edit) => RemoveInvalidSparklines(in edit);

    void ISheetListener.OnInsertAreaAndShiftRight(in SheetEdit edit) => RemoveInvalidSparklines(in edit);

    void ISheetListener.OnDeleteAreaAndShiftUp(in SheetEdit edit) => RemoveInvalidSparklines(in edit);

    void ISheetListener.OnDeleteAreaAndShiftLeft(in SheetEdit edit) => RemoveInvalidSparklines(in edit);

    private void RemoveInvalidSparklines(in SheetEdit edit)
    {
        if (edit.Sheet != _worksheet)
            return;

        var invalidSparklines = _sparklineGroups.SelectMany(g => g)
            .Where(sl => !((XLAddress)sl.Location.Address).IsValid)
            .ToList();

        foreach (var sparkline in invalidSparklines)
        {
            Remove(sparkline.Location);
        }
    }

    #endregion ISheetListener

    #region Private Fields

    private readonly List<IXLSparklineGroup> _sparklineGroups = new List<IXLSparklineGroup>();

    #endregion Private Fields

    /// <summary>
    /// Shift address of all sparklines to reflect inserted columns before a range.
    /// </summary>
    /// <param name="shiftedRange">Range before which will the columns be inserted. Has same worksheet.</param>
    /// <param name="numberOfColumns">How many columns, can be positive or negative number.</param>
    internal void ShiftColumns(Area shiftedRange, int numberOfColumns)
    {
        foreach (var group in _sparklineGroups)
        {
            foreach (var sparkline in group.ToList())
            {
                var originalAddress = Point.FromAddress(sparkline.Location.Address);
                if (!originalAddress.InRangeOrToLeft(shiftedRange))
                    continue;

                var newAddressColumn = originalAddress.Column + numberOfColumns;
                if (newAddressColumn is >= 1 and <= XLHelper.MaxColumnNumber)
                    sparkline.Location = new XLCell(_worksheet, originalAddress.Row, newAddressColumn);
                else
                    group.Remove(sparkline);
            }
        }
    }

    /// <summary>
    /// Shift address of all sparklines to reflect inserted rows before a range.
    /// </summary>
    /// <param name="shiftedRange">Range before which will the rows be inserted. Has same worksheet.</param>
    /// <param name="numberOfRows">How many rows, can be positive or negative number.</param>
    internal void ShiftRows(Area shiftedRange, int numberOfRows)
    {
        foreach (var group in _sparklineGroups)
        {
            foreach (var sparkline in group.ToList())
            {
                var originalAddress = Point.FromAddress(sparkline.Location.Address);
                if (!originalAddress.InRangeOrBelow(shiftedRange))
                    continue;

                var newAddressRow = originalAddress.Row + numberOfRows;
                if (newAddressRow is >= 1 and <= XLHelper.MaxRowNumber)
                    sparkline.Location = new XLCell(_worksheet, newAddressRow, originalAddress.Column);
                else
                    group.Remove(sparkline);
            }
        }
    }
}
