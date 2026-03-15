using System;
using System.Linq;
using XLibur.Extensions;

namespace XLibur.Excel;

using System.Collections.Generic;

internal sealed class XLAutoFilter : IXLAutoFilter
{
    /// <summary>
    /// The key is the column number.
    /// </summary>
    private readonly Dictionary<int, XLFilterColumn> _columns = new();

    /// <summary>
    /// When true, <see cref="Reapply"/> will not expand the range to include new data rows.
    /// Used for table autofilters where the table manages the range.
    /// </summary>
    internal bool IsTableAutoFilter { get; set; }

    internal IReadOnlyDictionary<int, XLFilterColumn> Columns => _columns;

    #region IXLAutoFilter Members

    public IEnumerable<IXLRangeRow> HiddenRows => Range.Rows(r => r.WorksheetRow().IsHidden);

    public bool IsEnabled { get; set; }

    public IXLRange Range { get; set; } = null!;

    public int SortColumn { get; set; }

    public bool Sorted { get; set; }

    public XLSortOrder SortOrder { get; set; }

    public IEnumerable<IXLRangeRow> VisibleRows => Range.Rows(r => !r.WorksheetRow().IsHidden);

    IXLAutoFilter IXLAutoFilter.Clear() => Clear();

    IXLFilterColumn IXLAutoFilter.Column(string columnLetter) => Column(columnLetter);

    IXLFilterColumn IXLAutoFilter.Column(int columnNumber) => Column(columnNumber);

    IXLAutoFilter IXLAutoFilter.Sort(int columnToSortBy, XLSortOrder sortOrder, bool matchCase, bool ignoreBlanks)
    {
        return Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks);
    }

    public IXLAutoFilter Reapply()
    {
        if (!IsEnabled) return this;

        ExpandRangeToData();

        var rows = Range.Rows(2, Range.RowCount());
        rows.ForEach(row => row.WorksheetRow().Unhide());

        foreach (var filterColumn in _columns.Values)
            filterColumn.Refresh();

        foreach (var row in rows)
        {
            if (!MatchesAllFilters(row))
                row.WorksheetRow().Hide();
        }

        return this;
    }

    /// <summary>
    /// Expand the autofilter range downward to include any contiguous data rows
    /// that were added after the filter was originally set.
    /// </summary>
    private void ExpandRangeToData()
    {
        if (IsTableAutoFilter)
            return;

        var rangeAddress = Range.RangeAddress;
        var ws = Range.Worksheet;
        var firstCol = rangeAddress.FirstAddress.ColumnNumber;
        var lastCol = rangeAddress.LastAddress.ColumnNumber;
        var currentLastRow = rangeAddress.LastAddress.RowNumber;

        // Walk downward from the row after the current range to find contiguous data.
        var newLastRow = currentLastRow;
        while (true)
        {
            var nextRow = newLastRow + 1;
            if (nextRow > XLHelper.MaxRowNumber)
                break;

            var hasData = false;
            for (var col = firstCol; col <= lastCol; col++)
            {
                if (!ws.Cell(nextRow, col).IsEmpty())
                {
                    hasData = true;
                    break;
                }
            }

            if (!hasData)
                break;

            newLastRow = nextRow;
        }

        if (newLastRow > currentLastRow)
        {
            Range = ws.Range(
                rangeAddress.FirstAddress.RowNumber, firstCol,
                newLastRow, lastCol);
        }
    }

    private bool MatchesAllFilters(IXLRangeRow row)
    {
        return _columns.All(kvp => kvp.Value.Check(row.Cell(kvp.Key)));
    }

    #endregion IXLAutoFilter Members

    private XLFilterColumn Column(string columnLetter)
    {
        var columnNumber = XLHelper.GetColumnNumberFromLetter(columnLetter);
        if (columnNumber is < 1 or > XLHelper.MaxColumnNumber)
            throw new ArgumentOutOfRangeException(nameof(columnLetter), "Column '" + columnLetter + "' is outside the allowed column range.");

        return Column(columnNumber);
    }

    internal XLFilterColumn Column(int columnNumber)
    {
        if (columnNumber is < 1 or > XLHelper.MaxColumnNumber)
            throw new ArgumentOutOfRangeException(nameof(columnNumber), "Column " + columnNumber + " is outside the allowed column range.");

        if (_columns.TryGetValue(columnNumber, out var filterColumn)) return filterColumn;
        filterColumn = new XLFilterColumn(this, columnNumber);
        _columns.Add(columnNumber, filterColumn);

        return filterColumn;
    }

    internal XLAutoFilter Clear()
    {
        if (!IsEnabled) return this;

        IsEnabled = false;
        foreach (var filterColumn in _columns.Values)
            filterColumn.Clear(false);

        foreach (IXLRangeRow row in Range.Rows().Where(r => r.RowNumber() > 1))
            row.WorksheetRow().Unhide();
        return this;
    }

    internal XLAutoFilter Set(IXLRangeBase range)
    {
        var firstOverlappingTable = range.Worksheet.Tables.FirstOrDefault(t => t.RangeUsed()!.Intersects(range));
        if (firstOverlappingTable != null)
            throw new InvalidOperationException($"The range {range.RangeAddress.ToStringRelative(includeSheet: true)} is already part of table '{firstOverlappingTable.Name}'");

        Range = range.AsRange();
        IsEnabled = true;
        return this;
    }

    private XLAutoFilter Sort(int columnToSortBy, XLSortOrder sortOrder, bool matchCase, bool ignoreBlanks)
    {
        if (!IsEnabled)
            throw new InvalidOperationException("Filter has not been enabled.");

        Range.Range(Range.FirstCell().CellBelow(), Range.LastCell()).Sort(columnToSortBy, sortOrder, matchCase,
            ignoreBlanks);

        Sorted = true;
        SortOrder = sortOrder;
        SortColumn = columnToSortBy;

        Reapply();

        return this;
    }
}
