using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel;

/// <summary>
/// Slice is a sparse array that stores a part of cell information (e.g., only values,
/// only styles ...). Slice has the same size as a worksheet. If some cells are pushed out
/// of the permitted range, they are gone.
/// </summary>
/// <remarks>
/// This is a ref return, so if the underlaying value
/// changes, the returned value also changes. To avoid,
/// just don't use <c>ref</c> and structs will be copied.
/// </remarks>
/// <typeparam name="TElement">The type of data stored in the slice.</typeparam>
internal sealed partial class Slice<TElement> : ISlice
{
    private readonly TElement _defaultValue = default!;

    /// <summary>
    /// The content of the slice. Note that LUT uses an index that starts from 0,
    /// so rows and columns must be adjusted to retrieve the value.
    /// </summary>
    private readonly Lut<RowData> _data;

    /// <summary>
    /// Key is the column number, value is the number of cells in the column that are used.
    /// </summary>
    private readonly Dictionary<int, int> _columnUsage = new();

    private int _version;

    internal Slice()
    {
        _data = new Lut<RowData>();
    }

    /// <inheritdoc />
    public int Version => _version;

    /// <summary>
    /// Get the slice value at the specified point of the sheet.
    /// </summary>
    internal ref readonly TElement this[Point point] => ref this[point.Row, point.Column];

    /// <summary>
    /// Get the slice value at the specified point of the sheet.
    /// </summary>
    internal ref readonly TElement this[int row, int column]
    {
        get
        {
            ref readonly var rowData = ref _data.Get(row - 1);
            return ref rowData.Get(column - 1);
        }
    }

    /// <inheritdoc />
    public bool IsEmpty => MaxRow == 0;

    /// <inheritdoc />
    public int MaxColumn { get; private set; }

    /// <inheritdoc />
    public int MaxRow => _data.MaxUsedIndex + 1;

    /// <inheritdoc />
    public IEnumerable<int> UsedRows
    {
        get
        {
            var rowsEnumerator = new Lut<RowData>.LutEnumerator(_data, XLHelper.MinRowNumber - 1, XLHelper.MaxRowNumber - 1);
            while (rowsEnumerator.MoveNext())
            {
                if (rowsEnumerator.Current.IsNonEmpty)
                    yield return rowsEnumerator.Index + 1;
            }
        }
    }

    /// <inheritdoc />
    public Dictionary<int, int>.KeyCollection UsedColumns => _columnUsage.Keys;

    /// <inheritdoc />
    public void Clear(Area range)
    {
        var enumerator = new Enumerator(this, range);
        while (enumerator.MoveNext())
        {
            Set(enumerator.Point, in _defaultValue);
        }
    }

    /// <inheritdoc />
    public void DeleteAreaAndShiftLeft(Area rangeToDelete)
    {
        Clear(rangeToDelete);

        var noCellsToShift = rangeToDelete.LastPoint.Column == XLHelper.MaxColumnNumber;
        if (noCellsToShift)
            return;

        var shiftDistance = rangeToDelete.Width;
        var shiftRange = rangeToDelete.RightRange();
        var cellEnumerator = new Enumerator(this, shiftRange);
        while (cellEnumerator.MoveNext())
        {
            var srcPoint = cellEnumerator.Point;
            var dstPoint = new Point(srcPoint.Row, srcPoint.Column - shiftDistance);
            Set(dstPoint, in cellEnumerator.Current);
            Set(srcPoint, in _defaultValue);
        }
    }

    /// <inheritdoc />
    public void DeleteAreaAndShiftUp(Area rangeToDelete)
    {
        Clear(rangeToDelete);

        var noCellsToShift = rangeToDelete.LastPoint.Row == XLHelper.MaxRowNumber;
        if (noCellsToShift)
            return;

        var shiftDistance = rangeToDelete.Height;
        var shiftRange = rangeToDelete.BelowRange();
        var cellEnumerator = new Enumerator(this, shiftRange);
        while (cellEnumerator.MoveNext())
        {
            var srcPoint = cellEnumerator.Point;
            var dstPoint = new Point(srcPoint.Row - shiftDistance, srcPoint.Column);
            Set(dstPoint, in cellEnumerator.Current);
            Set(srcPoint, in _defaultValue);
        }
    }

    /// <summary>
    /// Removes a set of whole rows and closes the gaps, in one pass.
    /// </summary>
    /// <remarks>
    /// <see cref="DeleteAreaAndShiftUp"/> works a cell at a time because the area it is given can be
    /// narrower than the sheet, so it has to move cells out from under the ones staying put. A whole
    /// row has no such constraint: the slice stores a row's entire set of columns as one
    /// <c>RowData</c>, so a surviving row moves by assigning that value at its new index. The cost
    /// drops from one operation per cell below the deletion to one per used row, and — since this
    /// takes the whole deletion at once rather than a block at a time — from once per deleted row to
    /// once altogether.
    /// <para>
    /// Rows are walked upward. Every surviving row's destination is at or above where it currently
    /// sits and destinations rise with sources, so a row is never written over one that has not been
    /// read yet, and no temporary copy is needed.
    /// </para>
    /// </remarks>
    public void DeleteRowsAndCompact(XLRowDeletionMap map)
    {
        var lastUsedRow = MaxRow;
        if (lastUsedRow == 0)
            return;

        var rowsEnumerator = new Lut<RowData>.LutEnumerator(_data, XLHelper.MinRowNumber - 1, lastUsedRow - 1);

        // The enumerator walks the same structure being written, and a write can clear the bucket it
        // is standing on, so collect the used rows before touching anything.
        var used = new List<(int Row, RowData Data)>();
        while (rowsEnumerator.MoveNext())
            used.Add((rowsEnumerator.Index + 1, rowsEnumerator.Current));

        // Destinations claimed by a surviving row, ascending — sources ascend and the mapping is
        // strictly increasing across survivors, so this list comes out sorted for free.
        var claimed = new List<int>(used.Count);

        foreach (var (row, data) in used)
        {
            if (map.IsDeleted(row))
            {
                // The row is going away, so every cell in it stops counting toward its column.
                ReleaseColumnUsage(in data);
                continue;
            }

            var target = map.MapFirst(row);
            if (target != row)
                _data.Set(target - 1, data);

            claimed.Add(target);
        }

        // Clear every slot that held something and did not receive a row. Two cases look the same from
        // here: a row that moved up and left its old slot behind, and a slot whose new occupant is an
        // *unused* row, which writes nothing and would otherwise leave the previous contents in place.
        // The second is the one that bites — a row with no explicit style or value is invisible to the
        // walk above, so its destination has to be cleared on its behalf.
        var next = 0;
        foreach (var (row, _) in used)
        {
            while (next < claimed.Count && claimed[next] < row)
                next++;

            if (next < claimed.Count && claimed[next] == row)
                continue;

            _data.Set(row - 1, default);
        }

        if (map.Count > 0)
            _version++;

        if (_columnUsage.Count == 0)
            MaxColumn = 0;
        else if (MaxColumn > 0 && !_columnUsage.ContainsKey(MaxColumn))
            MaxColumn = CalculateMaxColumn();
    }

    /// <summary>
    /// Drops a departing row's cells from the per-column usage counts.
    /// </summary>
    private void ReleaseColumnUsage(in RowData data)
    {
        var columns = data.GetColumnEnumerator(XLHelper.MinColumnNumber - 1, XLHelper.MaxColumnNumber - 1);
        while (columns.MoveNext())
            DecrementColumnUsage(columns.Index + 1);
    }

    /// <summary>
    /// Get enumerator over used values of the range.
    /// </summary>
    public IEnumerator<Point> GetEnumerator(Area range, bool reverse = false)
    {
        return !reverse ? new Enumerator(this, range) : new ReverseEnumerator(this, range);
    }

    /// <inheritdoc />
    public void InsertAreaAndShiftDown(Area range)
    {
        var hasSpaceBelow = range.LastPoint.Row < XLHelper.MaxRowNumber;
        if (!hasSpaceBelow)
        {
            Clear(range);
            return;
        }

        var shiftDistance = range.Height;

        // Purged range might contain some cells that wouldn't be overwritten during shift => clear.
        var purgedRange = new Area(
            new Point(XLHelper.MaxRowNumber - shiftDistance + 1, range.FirstPoint.Column),
            new Point(XLHelper.MaxRowNumber, range.LastPoint.Column));
        Clear(purgedRange);

        var shiftedRange = new Area(
            range.FirstPoint,
            new Point(XLHelper.MaxRowNumber - shiftDistance, range.LastPoint.Column));
        var cellEnumerator = new ReverseEnumerator(this, shiftedRange);
        while (cellEnumerator.MoveNext())
        {
            var srcPoint = cellEnumerator.Point;
            var dstPoint = new Point(srcPoint.Row + shiftDistance, srcPoint.Column);
            Set(dstPoint, in cellEnumerator.Current);
            Set(srcPoint, in _defaultValue);
        }
    }

    /// <inheritdoc />
    public void InsertAreaAndShiftRight(Area range)
    {
        var hasSpaceRight = range.LastPoint.Column < XLHelper.MaxColumnNumber;
        if (!hasSpaceRight)
        {
            Clear(range);
            return;
        }

        var shiftDistance = range.Width;

        // Purged range might contain some cells that wouldn't be overwritten during shift => clear.
        var purgedRange = new Area(
            new Point(range.FirstPoint.Row, XLHelper.MaxColumnNumber - shiftDistance + 1),
            new Point(range.LastPoint.Row, XLHelper.MaxColumnNumber));
        Clear(purgedRange);

        var shiftedRange = new Area(
            range.FirstPoint,
            new Point(range.LastPoint.Row, XLHelper.MaxColumnNumber - shiftDistance));
        var enumerator = new ReverseEnumerator(this, shiftedRange);
        while (enumerator.MoveNext())
        {
            var srcPoint = enumerator.Point;
            var dstPoint = new Point(srcPoint.Row, srcPoint.Column + shiftDistance);
            Set(dstPoint, in enumerator.Current);
            Set(srcPoint, in _defaultValue);
        }
    }

    public bool IsUsed(Point address)
    {
        ref readonly var rowData = ref _data.Get(address.Row - 1);
        return rowData.IsUsed(address.Column - 1);
    }

    public void Swap(Point sp1, Point sp2)
    {
        var value1 = this[sp1];
        var value2 = this[sp2];
        Set(sp1, in value2);
        Set(sp2, in value1);
    }

    internal void Set(Point point, in TElement value)
        => Set(point.Row, point.Column, in value);

    internal void Set(int row, int column, in TElement value)
    {
        ref readonly var existing = ref _data.Get(row - 1);
        if (existing.IsEmpty)
        {
            // Don't allocate a row just to store the default value.
            if (EqualityComparer<TElement>.Default.Equals(value, _defaultValue))
                return;

            var rowData = RowData.CreateForSet(column - 1, value);
            _data.Set(row - 1, rowData);
            IncrementColumnUsage(column);
            if (column > MaxColumn)
                MaxColumn = column;
            _version++;
            return;
        }

        // Copy the struct so we can mutate it.
        var rd = existing;
        var wasUsed = rd.IsUsed(column - 1);
        rd.Set(column - 1, value);
        var isUsed = rd.IsUsed(column - 1);

        // Write back (outer Lut detects default RowData and clears its bitmap).
        _data.Set(row - 1, rd);

        if (wasUsed && !isUsed)
        {
            var newCount = DecrementColumnUsage(column);
            if (newCount == 0 && MaxColumn == column)
            {
                MaxColumn = CalculateMaxColumn();
            }
        }

        if (!wasUsed && isUsed)
        {
            IncrementColumnUsage(column);
            if (column > MaxColumn)
                MaxColumn = column;
        }

        _version++;
    }

    /// <summary>
    /// Fast path for bulk-loading non-default values. The caller guarantees that <paramref name="value"/>
    /// is not <c>default</c>, so we skip <see cref="EqualityComparer{T}"/> checks and the
    /// "was-used-now-unused" bookkeeping that cannot happen during a load of non-default data.
    /// </summary>
    internal void SetNonDefault(Point point, in TElement value)
        => SetNonDefault(point.Row, point.Column, in value);

    /// <inheritdoc cref="SetNonDefault(Point, in TElement)"/>
    internal void SetNonDefault(int row, int column, in TElement value)
    {
        ref readonly var existing = ref _data.Get(row - 1);
        if (existing.IsEmpty)
        {
            var rowData = RowData.CreateForSet(column - 1, value);
            _data.SetNonDefault(row - 1, rowData);
            IncrementColumnUsage(column);
            if (column > MaxColumn)
                MaxColumn = column;
            _version++;
            return;
        }

        // Copy the struct so we can mutate it.
        var rd = existing;
        var wasUsed = rd.IsUsed(column - 1);
        rd.SetNonDefault(column - 1, value);

        // Write back — value is non-default so RowData is always non-empty.
        _data.SetNonDefault(row - 1, rd);

        if (!wasUsed)
        {
            IncrementColumnUsage(column);
            if (column > MaxColumn)
                MaxColumn = column;
        }

        _version++;
    }

    private int CalculateMaxColumn()
    {
        var maxColIdx = -1;
        var rowEnumerator = new Lut<RowData>.LutEnumerator(_data, XLHelper.MinRowNumber - 1, XLHelper.MaxRowNumber - 1);
        while (rowEnumerator.MoveNext())
            maxColIdx = Math.Max(maxColIdx, rowEnumerator.Current.MaxUsedIndex);

        return maxColIdx + 1;
    }

    private int DecrementColumnUsage(int column)
    {
        if (!_columnUsage.TryGetValue(column, out var count))
            return 0;

        if (count > 1)
            return _columnUsage[column] = count - 1;

        _columnUsage.Remove(column);
        return 0;
    }

    private void IncrementColumnUsage(int column)
    {
        if (_columnUsage.TryGetValue(column, out var value))
            _columnUsage[column] = value + 1;
        else
            _columnUsage.Add(column, 1);
    }

    /// <summary>
    /// Enumerator that returns used values from a specified range.
    /// </summary>
    [DebuggerDisplay("{Point}:{Current}")]
    internal sealed class Enumerator : IEnumerator<Point>
    {
        private readonly Area _range;
        private ColumnEnumerator _columnsEnumerator;
        private Lut<RowData>.LutEnumerator _rowsEnumerator;

        internal Enumerator(Slice<TElement> slice, Area range)
        {
            _range = range;

            _columnsEnumerator = default;
            _rowsEnumerator = new Lut<RowData>.LutEnumerator(
                slice._data,
                range.FirstPoint.Row - 1,
                range.LastPoint.Row - 1);
        }

        public ref readonly TElement Current => ref _columnsEnumerator.Current;

        public Point Point => new(_rowsEnumerator.Index + 1, _columnsEnumerator.Index + 1);

        /// <summary>
        /// The movement is columns first, then rows.
        /// </summary>
        public bool MoveNext()
        {
            while (!_columnsEnumerator.MoveNext())
            {
                if (!_rowsEnumerator.MoveNext())
                    return false;

                _columnsEnumerator = _rowsEnumerator.Current.GetColumnEnumerator(
                    _range.FirstPoint.Column - 1,
                    _range.LastPoint.Column - 1);
            }

            return true;
        }

        void IEnumerator.Reset() => throw new NotSupportedException();

        Point IEnumerator<Point>.Current => Point;

        object IEnumerator.Current => Point;

        void IDisposable.Dispose() { }
    }

    [DebuggerDisplay("{Point}:{Current}")]
    private sealed class ReverseEnumerator : IEnumerator<Point>
    {
        private readonly Area _range;
        private ReverseColumnEnumerator _columnsEnumerator;
        private Lut<RowData>.ReverseLutEnumerator _rowsEnumerator;

        internal ReverseEnumerator(Slice<TElement> slice, Area range)
        {
            _range = range;
            _columnsEnumerator = default;
            _rowsEnumerator = new Lut<RowData>.ReverseLutEnumerator(
                slice._data,
                range.FirstPoint.Row - 1,
                range.LastPoint.Row - 1);
        }

        public ref TElement Current => ref _columnsEnumerator.Current;

        public Point Point => new(_rowsEnumerator.Index + 1, _columnsEnumerator.Index + 1);

        public bool MoveNext()
        {
            while (!_columnsEnumerator.MoveNext())
            {
                if (!_rowsEnumerator.MoveNext())
                    return false;

                _columnsEnumerator = _rowsEnumerator.Current.GetReverseColumnEnumerator(
                    _range.FirstPoint.Column - 1,
                    _range.LastPoint.Column - 1);
            }
            return true;
        }

        void IEnumerator.Reset() => throw new NotSupportedException();

        Point IEnumerator<Point>.Current => Point;

        object IEnumerator.Current => Point;

        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }
    }
}
