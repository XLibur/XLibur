using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel.Coordinates;
using XLibur.Extensions;

namespace XLibur.Excel;

internal sealed class XLCells : XLStylizedBase, IXLCells, IXLStylized, IEnumerable<XLCell>
{
    #region Fields

    private readonly List<XLRangeAddress> _rangeAddresses = new List<XLRangeAddress>();
    private readonly bool _usedCellsOnly;
    private readonly Func<IXLCell, bool> _predicate;
    private readonly XLCellsUsedOptions _options;
    private bool _styleInitialized;

    #endregion Fields

    #region Constructor

    public XLCells(bool usedCellsOnly, XLCellsUsedOptions options, Func<IXLCell, bool>? predicate = null)
        : base(XLStyle.Default.Value)
    {
        _usedCellsOnly = usedCellsOnly;
        _options = options;
        _predicate = predicate ?? (_ => true);
    }

    #endregion Constructor

    #region IEnumerable<XLCell> Members

    private IEnumerable<XLCell> GetAllCells()
    {
        var groupedAddresses = _rangeAddresses.GroupBy(addr => addr.Worksheet);
        foreach (var worksheetGroup in groupedAddresses)
        {
            var ws = worksheetGroup.Key!;
            var sheetPoints = worksheetGroup.SelectMany(addr => GetAllCellsInRange(addr))
                .Distinct();
            foreach (var sheetPoint in sheetPoints)
            {
                var c = ws.Cell(sheetPoint.Row, sheetPoint.Column);
                if (_predicate(c))
                    yield return c;
            }
        }
    }

    private static IEnumerable<Point> GetAllCellsInRange(IXLRangeAddress rangeAddress)
    {
        if (!rangeAddress.IsValid)
            yield break;

        var normalizedAddress = ((XLRangeAddress)rangeAddress).Normalize();
        var minRow = normalizedAddress.FirstAddress.RowNumber;
        var maxRow = normalizedAddress.LastAddress.RowNumber;
        var minColumn = normalizedAddress.FirstAddress.ColumnNumber;
        var maxColumn = normalizedAddress.LastAddress.ColumnNumber;

        for (var ro = minRow; ro <= maxRow; ro++)
        {
            for (var co = minColumn; co <= maxColumn; co++)
            {
                yield return new Point(ro, co);
            }
        }
    }

    /// <summary>
    /// The option flags that make <see cref="GetUsedCellsCandidates"/> contribute cells from
    /// somewhere other than the cell slices. Only those cells arrive out of order, so only they
    /// make the sort in <see cref="GetUsedCellsOrdered"/> necessary.
    /// </summary>
    private const XLCellsUsedOptions CandidateOptions =
        XLCellsUsedOptions.MergedRanges |
        XLCellsUsedOptions.ConditionalFormats |
        XLCellsUsedOptions.DataValidation |
        XLCellsUsedOptions.Sparklines;

    private IEnumerable<XLCell> GetUsedCells()
    {
        // One range on one sheet, with nothing outside the slices to contribute: stream it.
        //
        // GetUsedCellsInRange reads through XLCellsCollection.SlicesEnumerator, a k-way merge over
        // the value, formula, style and misc slice enumerators. It picks the smallest Point at each
        // step and advances every enumerator sitting on it, so what comes out is strictly ascending
        // and carries no duplicates. Point packs the row above the column, which makes ascending
        // packed order exactly row-major order - the order OrderBy(row).ThenBy(column) produces.
        //
        // So on this path the sort re-sorts sorted input and the visited set can never reject
        // anything, while between them they cost 88.6 ms and 60 MB of the 101.8 ms and 84.7 MB that
        // enumerating 500,000 used cells took (spec 19, UsedCellEnumerationBenchmarks). Worse, the
        // sort is eager: it buffers every cell in the sheet before yielding the first, so
        // CellsUsed().First() walked and sorted the lot.
        if (_rangeAddresses.Count == 1)
        {
            var rangeAddress = _rangeAddresses[0];
            var ws = rangeAddress.Worksheet;
            if (ws is not null && !HasCandidates(ws))
                return GetUsedCellsInRange(rangeAddress, ws, Enumerable.Empty<Point>());
        }

        return GetUsedCellsOrdered();
    }

    private IEnumerable<XLCell> GetUsedCellsOrdered()
    {
        var groupedAddresses = _rangeAddresses.GroupBy(addr => addr.Worksheet);
        foreach (var worksheetGroup in groupedAddresses)
        {
            var ws = worksheetGroup.Key!;

            var usedCellsCandidates = GetUsedCellsCandidates(ws);

            var cells = worksheetGroup.SelectMany(addr => GetUsedCellsInRange(addr, ws, usedCellsCandidates))
                .OrderBy(cell => cell.Address.RowNumber)
                .ThenBy(cell => cell.Address.ColumnNumber);

            // Duplicates land next to each other once the sequence is sorted by (row, column), so
            // remembering the previous address rejects exactly what a set of every address seen
            // would - at O(1) rather than one entry per used cell. Every cell in a group belongs to
            // the one worksheet, so equal addresses are the same cell.
            var havePrevious = false;
            var previous = default(XLAddress);
            foreach (var cell in cells)
            {
                var address = cell.Address;
                if (havePrevious && previous.Equals(address))
                    continue;

                previous = address;
                havePrevious = true;
                yield return cell;
            }
        }
    }

    /// <summary>
    /// Whether <see cref="GetUsedCellsCandidates"/> could yield anything for this sheet. Checks the
    /// sheet as well as the options, because asking for merged ranges on a sheet that has none still
    /// leaves the candidate sequence empty.
    /// </summary>
    private bool HasCandidates(XLWorksheet worksheet)
    {
        if (_options == XLCellsUsedOptions.AllContents || (_options & CandidateOptions) == 0)
            return false;

        if (_options.HasFlag(XLCellsUsedOptions.MergedRanges) && worksheet.Internals.MergedRanges.Count > 0)
            return true;

        if (_options.HasFlag(XLCellsUsedOptions.ConditionalFormats) && worksheet.ConditionalFormats.Any())
            return true;

        if (_options.HasFlag(XLCellsUsedOptions.DataValidation) && worksheet.DataValidations.Any())
            return true;

        return _options.HasFlag(XLCellsUsedOptions.Sparklines) && worksheet.SparklineGroups.Any(sg => sg.Any());
    }

    private IEnumerable<XLCell> GetUsedCellsInRange(XLRangeAddress rangeAddress, XLWorksheet worksheet, IEnumerable<Point> usedCellsCandidates)
    {
        if (!rangeAddress.IsValid)
            yield break;
        var normalizedAddress = rangeAddress.Normalize();
        var minRow = normalizedAddress.FirstAddress.RowNumber;
        var maxRow = normalizedAddress.LastAddress.RowNumber;
        var minColumn = normalizedAddress.FirstAddress.ColumnNumber;
        var maxColumn = normalizedAddress.LastAddress.ColumnNumber;

        var cellRange = worksheet.Internals.CellsCollection
            .GetCells(minRow, minColumn, maxRow, maxColumn, _predicate);

        foreach (var cell in cellRange)
        {
            if (!cell.IsEmpty(_options) && _predicate(cell))
                yield return cell;
        }

        foreach (var sheetPoint in usedCellsCandidates)
        {
            if (sheetPoint.Row.Between(minRow, maxRow) &&
                sheetPoint.Column.Between(minColumn, maxColumn))
            {
                var cell = worksheet.Cell(sheetPoint.Row, sheetPoint.Column);

                if (_predicate(cell))
                    yield return cell;
            }
        }
    }

    private IEnumerable<Point> GetUsedCellsCandidates(XLWorksheet worksheet)
    {
        var candidates = Enumerable.Empty<Point>();

        if (_options == XLCellsUsedOptions.AllContents)
        {
            return candidates;
        }

        if (_options.HasFlag(XLCellsUsedOptions.MergedRanges))
            candidates = candidates.Union(
                worksheet.Internals.MergedRanges.SelectMany(r => GetAllCellsInRange(r.RangeAddress)));

        if (_options.HasFlag(XLCellsUsedOptions.ConditionalFormats))
            candidates = candidates.Union(
                worksheet.ConditionalFormats.SelectMany(cf => cf.Ranges.SelectMany(r => GetAllCellsInRange(r.RangeAddress))));

        if (_options.HasFlag(XLCellsUsedOptions.DataValidation))
            candidates = candidates.Union(
                worksheet.DataValidations.SelectMany(dv => dv.Ranges.SelectMany(r => GetAllCellsInRange(r.RangeAddress))));

        if (_options.HasFlag(XLCellsUsedOptions.Sparklines))
            candidates = candidates.Union(
                worksheet.SparklineGroups.SelectMany(sg => sg).Select(sl => Point.FromAddress(sl.Location.Address)));

        return candidates.Distinct();
    }

    public IEnumerator<XLCell> GetEnumerator()
    {
        return GetCells().GetEnumerator();
    }

    private IEnumerable<XLCell> GetCells()
    {
        return _usedCellsOnly ? GetUsedCells() : GetAllCells();
    }

    #endregion IEnumerable<XLCell> Members

    #region IXLCells Members

    IEnumerator<IXLCell> IEnumerable<IXLCell>.GetEnumerator()
    {
        return GetCells().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

#pragma warning disable S2376 // Write-only properties: intentional batch-set on collection items
    public XLCellValue Value
    {
        set { this.ForEach<XLCell>(c => c.Value = value); }
    }
#pragma warning restore S2376

    public IXLCells Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        this.ForEach<XLCell>(c => c.Clear(clearOptions));
        return this;
    }

    public void DeleteComments()
    {
        this.ForEach<XLCell>(c => c.DeleteComment());
    }

    public void DeleteSparklines()
    {
        this.ForEach<XLCell>(c => c.DeleteSparkline());
    }

#pragma warning disable S2376 // Write-only properties: intentional batch-set on collection items
    public string FormulaA1
    {
        set { this.ForEach<XLCell>(c => c.FormulaA1 = value); }
    }

    public string FormulaR1C1
    {
        set { this.ForEach<XLCell>(c => c.FormulaR1C1 = value); }
    }
#pragma warning restore S2376

    #endregion IXLCells Members

    #region IXLStylized Members

    protected override IEnumerable<XLStylizedBase> Children
    {
        get
        {
            foreach (XLCell c in this)
                yield return c;
        }
    }

    public override IXLRanges RangesUsed
    {
        get
        {
            var retVal = new XLRanges();
            this.ForEach<XLCell>(c => retVal.Add(c.AsRange()));
            return retVal;
        }
    }

    #endregion IXLStylized Members

    public void Add(XLRangeAddress rangeAddress)
    {
        _rangeAddresses.Add(rangeAddress);

        if (_styleInitialized)
            return;

        var worksheetStyle = rangeAddress.Worksheet?.Style;
        if (worksheetStyle == null)
            return;

        InnerStyle = worksheetStyle;
        _styleInitialized = true;
    }

    public void Add(XLCell cell)
    {
        Add(new XLRangeAddress(cell.Address, cell.Address));
    }

    public void Select()
    {
        foreach (var cell in this)
            cell.Select();
    }
}
