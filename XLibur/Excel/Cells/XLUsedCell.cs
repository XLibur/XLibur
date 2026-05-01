using System.Collections.Generic;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel;

/// <summary>
/// Snapshot of a used cell yielded by <see cref="IXLWorksheet.EnumerateUsedCells"/>.
/// Holds the cell's address (1-based row and column) and its current value.
/// </summary>
/// <remarks>
/// The snapshot is captured at the moment the enumerator yields it; later mutations
/// to the cell's value are not observed by an in-flight iteration.
/// </remarks>
public readonly struct XLUsedCell
{
    private readonly XLSheetPoint _point;

    internal XLUsedCell(XLSheetPoint point, XLCellValue value)
    {
        _point = point;
        Value = value;
    }

    /// <summary>The 1-based row number of the cell.</summary>
    public int Row => _point.Row;

    /// <summary>The 1-based column number of the cell.</summary>
    public int Column => _point.Column;

    /// <summary>The cell's value at the moment it was yielded.</summary>
    public XLCellValue Value { get; }
}

/// <summary>
/// Allocation-free enumerable over the used cells of a worksheet, returned by
/// <see cref="IXLWorksheet.EnumerateUsedCells"/>. Designed to be consumed via
/// <c>foreach</c>; do not store. Iteration yields one <see cref="XLUsedCell"/>
/// per non-blank cell value without allocating a wrapper per cell.
/// </summary>
public ref struct XLUsedCellEnumerable
{
    private readonly ValueSlice _valueSlice;
    private readonly XLSheetRange _range;

    internal XLUsedCellEnumerable(ValueSlice valueSlice, XLSheetRange range)
    {
        _valueSlice = valueSlice;
        _range = range;
    }

    /// <summary>Returns the underlying enumerator for use by <c>foreach</c>.</summary>
    public XLUsedCellEnumerator GetEnumerator() => new(_valueSlice, _range);
}

/// <summary>
/// Enumerator for <see cref="XLUsedCellEnumerable"/>. Reads cell values directly
/// from the worksheet's value slice without allocating a wrapper per cell.
/// </summary>
public ref struct XLUsedCellEnumerator
{
    private readonly ValueSlice _valueSlice;
    private readonly XLSheetRange _range;
    private IEnumerator<XLSheetPoint>? _inner;
    private XLUsedCell _current;

    internal XLUsedCellEnumerator(ValueSlice valueSlice, XLSheetRange range)
    {
        _valueSlice = valueSlice;
        _range = range;
        _inner = null;
        _current = default;
    }

    /// <summary>The cell snapshot at the current position.</summary>
    public XLUsedCell Current => _current;

    /// <summary>
    /// Advances the enumerator to the next used cell. Lazily constructs the
    /// underlying slice enumerator on the first call, so an empty <c>foreach</c>
    /// over a sheet with no used cells does no allocation at all.
    /// </summary>
    public bool MoveNext()
    {
        _inner ??= _valueSlice.GetEnumerator(_range);
        if (!_inner.MoveNext())
        {
            _inner.Dispose();
            return false;
        }

        var point = _inner.Current;
        _current = new XLUsedCell(point, _valueSlice.GetCellValue(point));
        return true;
    }

    /// <summary>Releases the underlying slice enumerator, if one was created.</summary>
    public void Dispose() => _inner?.Dispose();
}
