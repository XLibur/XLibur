using System;
using System.Buffers;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel;

internal sealed class XLCellsCollection : IWorkbookListener
{
    private readonly XLWorksheet _ws;
    private readonly List<ISlice> _slices;

    // Caches for ColumnsUsedKeys / RowsUsedKeys. Each cache is invalidated when the sum of
    // slice versions changes (any slice mutation bumps its own version, which changes the sum).
    // Allocates the result HashSet once per "snapshot" of slice state instead of once per call.
    private HashSet<int>? _cachedColumnsUsedKeys;
    private int _cachedColumnsUsedKeysVersionSum;
    private HashSet<int>? _cachedRowsUsedKeys;
    private int _cachedRowsUsedKeysVersionSum;

    public XLCellsCollection(XLWorksheet ws)
    {
        _ws = ws;
        ValueSlice = new ValueSlice(ws.Workbook.SharedStringTable);
        FormulaSlice = new FormulaSlice(ws);
        _slices = [ValueSlice, FormulaSlice, StyleSlice, MiscSlice];
    }

    internal HashSet<int> ColumnsUsedKeys
    {
        get
        {
            var versionSum = ValueSlice.Version + FormulaSlice.Version + StyleSlice.Version + MiscSlice.Version;
            if (_cachedColumnsUsedKeys is { } cached && _cachedColumnsUsedKeysVersionSum == versionSum)
                return cached;

            var set = _cachedColumnsUsedKeys;
            if (set is null)
                set = new HashSet<int>();
            else
                set.Clear();

            foreach (var slice in _slices)
                set.UnionWith(slice.UsedColumns);

            _cachedColumnsUsedKeys = set;
            _cachedColumnsUsedKeysVersionSum = versionSum;
            return set;
        }
    }

    internal bool IsEmpty
    {
        get
        {
            foreach (var slice in _slices)
            {
                if (!slice.IsEmpty)
                    return false;
            }

            return true;
        }
    }

    internal int MaxColumnUsed
    {
        get
        {
            var max = int.MinValue;
            foreach (var slice in _slices)
                max = Math.Max(max, slice.MaxColumn);

            return Math.Max(1, max);
        }
    }

    internal int MaxRowUsed
    {
        get
        {
            var max = int.MinValue;
            foreach (var slice in _slices)
                max = Math.Max(max, slice.MaxRow);

            return Math.Max(1, max);
        }
    }

    internal HashSet<int> RowsUsedKeys
    {
        get
        {
            var versionSum = ValueSlice.Version + FormulaSlice.Version + StyleSlice.Version + MiscSlice.Version;
            if (_cachedRowsUsedKeys is { } cached && _cachedRowsUsedKeysVersionSum == versionSum)
                return cached;

            var set = _cachedRowsUsedKeys;
            if (set is null)
                set = new HashSet<int>();
            else
                set.Clear();

            foreach (var slice in _slices)
                set.UnionWith(slice.UsedRows);

            _cachedRowsUsedKeys = set;
            _cachedRowsUsedKeysVersionSum = versionSum;
            return set;
        }
    }

    internal ValueSlice ValueSlice { get; }

    internal FormulaSlice FormulaSlice { get; }

    internal Slice<XLStyleValue?> StyleSlice { get; } = new();

    internal Slice<XLMiscSliceContent> MiscSlice { get; } = new();

    internal XLWorksheet Worksheet => _ws;

    internal void Clear()
    {
        Clear(Area.Full);
    }

    internal void Clear(Area clearRange)
    {
        foreach (var slice in _slices)
            slice.Clear(clearRange);
    }

    internal void DeleteAreaAndShiftLeft(Area rangeToDelete)
    {
        foreach (var slice in _slices)
            slice.DeleteAreaAndShiftLeft(rangeToDelete);
    }

    internal void DeleteAreaAndShiftUp(Area rangeToDelete)
    {
        foreach (var slice in _slices)
            slice.DeleteAreaAndShiftUp(rangeToDelete);
    }

    /// <summary>
    /// Remove a set of whole rows from every slice and close the gaps, in one pass each.
    /// </summary>
    internal void DeleteRowsAndCompact(XLRowDeletionMap map)
    {
        foreach (var slice in _slices)
            slice.DeleteRowsAndCompact(map);
    }

    /// <summary>
    /// Direct-mapped cache of recently vended cell wrappers. An <see cref="XLCell"/> is a stateless
    /// handle over (collection, point) — all cell data lives in the slices — so returning the same
    /// instance twice for the same point is equivalent to returning two, and lets the caller reuse
    /// the <c>XLStyle</c> that <see cref="XLStylizedBase.InnerStyle"/> caches on it.
    /// </summary>
    /// <remarks>
    /// Thread-safety: entries are single reference reads/writes, and the point is read back off the
    /// cached cell rather than from a parallel array, so a racing writer can only cause a miss —
    /// never a cell for the wrong point.
    /// </remarks>
    private const int CellCacheSize = 16;

    private const int CellCacheMask = CellCacheSize - 1;

    private readonly XLCell?[] _cellCache = new XLCell?[CellCacheSize];

    internal XLCell GetCell(Point address)
    {
        var slot = (int)address.PackedValue & CellCacheMask;
        var cached = _cellCache[slot];
        if (cached is not null && cached.SheetPoint.Equals(address))
            return cached;

        var cell = new XLCell(_ws, address);
        _cellCache[slot] = cell;
        return cell;
    }

    /// <summary>
    /// Get all used cells in the worksheet.
    /// </summary>
    internal IEnumerable<XLCell> GetCells()
    {
        return GetCells(Area.Full);
    }

    /// <summary>
    /// Get all used cells in the worksheet that satisfy the predicate.
    /// </summary>
    internal IEnumerable<XLCell> GetCells(Func<XLCell, bool> predicate)
    {
        return GetCells(Area.Full, predicate);
    }

    /// <summary>
    /// Get all used cells in the range that satisfy the predicate.
    /// </summary>
    internal IEnumerable<XLCell> GetCells(int rowStart, int columnStart,
        int rowEnd, int columnEnd,
        Func<XLCell, bool>? predicate = null)
    {
        return GetCells(new Area(rowStart, columnStart, rowEnd, columnEnd), predicate);
    }

    /// <summary>
    /// Get all used cells in the range that satisfy the predicate.
    /// </summary>
    internal IEnumerable<XLCell> GetCells(Area range, Func<XLCell, bool>? predicate = null)
    {
        var enumerator = new SlicesEnumerator(range, this);

        while (enumerator.MoveNext())
        {
            var cellAddress = enumerator.Current;
            var cell = GetCell(cellAddress);
            if (predicate == null || predicate(cell))
                yield return cell;
        }
    }

    /// <summary>
    /// Whether any cell in the sheet has a note or a comment thread. Both count, because a thread is
    /// written out with a paired legacy note and so needs the same comments and VML parts. Scans the
    /// misc slice directly, so it costs nothing on sheets that never touched comments, phonetics or
    /// metadata.
    /// </summary>
    internal bool HasAnyComment()
    {
        if (MiscSlice.IsEmpty)
            return false;

        var enumerator = new Slice<XLMiscSliceContent>.Enumerator(MiscSlice, Area.Full);
        while (enumerator.MoveNext())
        {
            if (enumerator.Current.Comment is not null || enumerator.Current.ThreadedComment is not null)
                return true;
        }

        return false;
    }

    /// <summary>
    /// Whether any cell in the sheet has a comment thread, i.e. whether the sheet needs a
    /// <c>threadedComment</c> part.
    /// </summary>
    internal bool HasAnyThreadedComment()
    {
        if (MiscSlice.IsEmpty)
            return false;

        var enumerator = new Slice<XLMiscSliceContent>.Enumerator(MiscSlice, Area.Full);
        while (enumerator.MoveNext())
        {
            if (enumerator.Current.ThreadedComment is not null)
                return true;
        }

        return false;
    }

    /// <summary>
    /// Every comment thread in the sheet, paired with the point it sits on, in row-major order.
    /// Scans the misc slice directly rather than materialising an <see cref="XLCell"/> per used
    /// cell: the save path asks this of every sheet on every save, including the overwhelmingly
    /// common case where no sheet has a thread at all.
    /// </summary>
    internal List<(Point Point, XLThreadedComment Thread)> GetThreadedComments()
    {
        var threads = new List<(Point, XLThreadedComment)>();
        if (MiscSlice.IsEmpty)
            return threads;

        var enumerator = new Slice<XLMiscSliceContent>.Enumerator(MiscSlice, Area.Full);
        while (enumerator.MoveNext())
        {
            if (enumerator.Current.ThreadedComment is { } thread)
                threads.Add((enumerator.Point, thread));
        }

        return threads;
    }

    /// <summary>
    /// The point of the cell that holds <paramref name="note"/>, or null when the note is no longer
    /// on the sheet. Matches by identity rather than by an address the note remembers, because
    /// shifting rows or columns moves the note's misc slice entry without telling the note. Walks
    /// the misc slice rather than materialising an <see cref="XLCell"/> per used cell.
    /// </summary>
    internal Point? FindNote(XLComment note)
    {
        if (MiscSlice.IsEmpty)
            return null;

        var enumerator = new Slice<XLMiscSliceContent>.Enumerator(MiscSlice, Area.Full);
        while (enumerator.MoveNext())
        {
            if (ReferenceEquals(enumerator.Current.Comment, note))
                return enumerator.Point;
        }

        return null;
    }

    /// <summary>
    /// The point of the cell that holds <paramref name="thread"/>, or null when the thread is no
    /// longer on the sheet. Matches by identity, for the same reason as <see cref="FindNote"/>.
    /// </summary>
    internal Point? FindThreadedComment(XLThreadedComment thread)
    {
        if (MiscSlice.IsEmpty)
            return null;

        var enumerator = new Slice<XLMiscSliceContent>.Enumerator(MiscSlice, Area.Full);
        while (enumerator.MoveNext())
        {
            if (ReferenceEquals(enumerator.Current.ThreadedComment, thread))
                return enumerator.Point;
        }

        return null;
    }

    /// <summary>
    /// Points of all cells that carry an in-cell image, in row-major order. Scans the misc slice
    /// directly rather than materialising an <see cref="XLCell"/> per used cell.
    /// </summary>
    internal List<Point> GetCellImagePoints()
    {
        var points = new List<Point>();
        if (MiscSlice.IsEmpty)
            return points;

        var enumerator = new Slice<XLMiscSliceContent>.Enumerator(MiscSlice, Area.Full);
        while (enumerator.MoveNext())
        {
            if (enumerator.Current.CellImage is not null)
                points.Add(enumerator.Point);
        }

        return points;
    }

    internal IEnumerable<XLCell> GetCellsInColumn(int column)
    {
        return GetCells(1, column, XLHelper.MaxRowNumber, column);
    }

    internal IEnumerable<XLCell> GetCellsInRow(int row)
    {
        return GetCells(row, 1, row, XLHelper.MaxColumnNumber);
    }

    /// <summary>
    /// Get cell or null, if cell is not used.
    /// </summary>
    internal XLCell? GetUsedCell(Point address)
    {
        if (!IsUsed(address))
            return null;

        return GetCell(address);
    }

    internal int FirstColumnUsed(Area searchRange, XLCellsUsedOptions options, Func<IXLCell, bool>? predicate = null)
    {
        return FindUsedColumn(searchRange, options, predicate, false);
    }

    internal int FirstRowUsed(Area searchRange, XLCellsUsedOptions options, Func<IXLCell, bool>? predicate = null)
    {
        return FindUsedRow(searchRange, options, predicate, false);
    }

    internal void InsertAreaAndShiftDown(Area insertedRange)
    {
        foreach (var slice in _slices)
            slice.InsertAreaAndShiftDown(insertedRange);
    }

    internal void InsertAreaAndShiftRight(Area insertedRange)
    {
        foreach (var slice in _slices)
            slice.InsertAreaAndShiftRight(insertedRange);
    }

    internal int LastColumnUsed(Area searchRange, XLCellsUsedOptions options, Func<IXLCell, bool>? predicate = null)
    {
        return FindUsedColumn(searchRange, options, predicate, true);
    }

    internal int LastRowUsed(Area searchRange, XLCellsUsedOptions options, Func<IXLCell, bool>? predicate = null)
    {
        return FindUsedRow(searchRange, options, predicate, true);
    }

    /// <summary>
    /// Remap rows of a range.
    /// </summary>
    /// <param name="map">A sorted map of rows. The values must be resorted row numbers from <paramref name="sheetRange"/>.</param>
    /// <param name="sheetRange">Sheet that should have its rows rearranged.</param>
    internal void RemapRows(IList<int> map, Area sheetRange)
    {
        RemapRanges(map, sheetRange.TopRow, SwapRows);

        void SwapRows(int prevRowNumber, int currentRowNumber)
        {
            var prevRowRange = new Area(
                new Point(prevRowNumber, sheetRange.LeftColumn),
                new Point(prevRowNumber, sheetRange.RightColumn));
            var currentRowRange = new Area(
                new Point(currentRowNumber, sheetRange.LeftColumn),
                new Point(currentRowNumber, sheetRange.RightColumn));
            SwapRanges(prevRowRange, currentRowRange);
        }
    }

    /// <summary>
    /// Remap columns of a range.
    /// </summary>
    /// <param name="map">A sorted map of columns. The values must be resorted columns numbers from <paramref name="sheetRange"/>.</param>
    /// <param name="sheetRange">Sheet that should have its columns rearranged.</param>
    internal void RemapColumns(IList<int> map, Area sheetRange)
    {
        RemapRanges(map, sheetRange.LeftColumn, SwapColumns);

        void SwapColumns(int prevColNumber, int currentColNumber)
        {
            var prevRowRange = new Area(
                new Point(sheetRange.TopRow, prevColNumber),
                new Point(sheetRange.BottomRow, prevColNumber));
            var currentRowRange = new Area(
                new Point(sheetRange.TopRow, currentColNumber),
                new Point(sheetRange.BottomRow, currentColNumber));
            SwapRanges(prevRowRange, currentRowRange);
        }
    }

    private static void RemapRanges(IList<int> map, int indexOffset, Action<int, int> swapData)
    {
        for (var i = 0; i < map.Count; ++i)
        {
            var axisNumber = i + indexOffset;
            var dataAxisNumber = map[i];
            if (axisNumber == dataAxisNumber)
                continue;

            // Current row doesn't contain data it should, so it is a part of a permutation
            // loop. Go over each item in a loop and
            // We need to replace
            var prevNumber = axisNumber;
            var currentNumber = dataAxisNumber;
            var startLoopNumber = prevNumber;
            do
            {
                // Current row number contains data that should be on the previous row number,
                // so swap them. That will fix another link in a loop (the previous one), but
                // will keep current inconsistent, but that will be fixed when loop completes.
                swapData(prevNumber, currentNumber);

                // Because previous row number is already fixed and will no longer be touched
                // during loop fix, mark it as a row that contains correct data.
                map[prevNumber - indexOffset] = prevNumber;

                prevNumber = currentNumber;
                currentNumber = map[currentNumber - indexOffset];
            } while (currentNumber != startLoopNumber);

            // Although we don't have to swap the last one (N count loop needs only N-1 swaps),
            // we have to mark the last row mapping for the last link (the one before start).
            map[prevNumber - indexOffset] = prevNumber;
        }
    }

    private void SwapRanges(Area sheetRange1, Area sheetRange2)
    {
        var rowCount = sheetRange1.LastPoint.Row - sheetRange1.FirstPoint.Row + 1;
        var columnCount = sheetRange1.LastPoint.Column - sheetRange1.FirstPoint.Column + 1;
        for (var row = 0; row < rowCount; row++)
        {
            for (var column = 0; column < columnCount; column++)
            {
                var sp1 = new Point(sheetRange1.FirstPoint.Row + row, sheetRange1.FirstPoint.Column + column);
                var sp2 = new Point(sheetRange2.FirstPoint.Row + row, sheetRange2.FirstPoint.Column + column);

                SwapCellsContent(sp1, sp2);
            }
        }
    }

    /// <summary>
    /// The first or last column in <paramref name="range"/> holding a cell that counts as used under
    /// <paramref name="options"/>, or 0 if there is none.
    /// </summary>
    /// <remarks>
    /// The candidate columns come from <see cref="ColumnsUsedKeys"/>, which already merges the four
    /// slices and caches the result until a slice changes. The previous implementation rebuilt that
    /// merge per call as <c>Concat</c> x4 -> <c>Where</c> -> <c>Distinct</c> -> <c>OrderBy</c>, so every
    /// <c>FirstColumnUsed</c>/<c>LastColumnUsed</c> allocated four iterators, a hash set and a sorted
    /// buffer over every used column in the sheet — beside a cache built for exactly this.
    /// </remarks>
    private int FindUsedColumn(Area range, XLCellsUsedOptions options, Func<IXLCell, bool>? predicate, bool descending)
    {
        var firstColumn = range.FirstPoint.Column;
        var lastColumn = range.LastPoint.Column;

        var candidates = ColumnsUsedKeys;
        if (candidates.Count == 0)
            return 0;

        var buffer = ArrayPool<int>.Shared.Rent(candidates.Count);
        int count;
        try
        {
            count = 0;
#pragma warning disable S3267 // Where would allocate the iterator this pooled buffer exists to avoid
            foreach (var column in candidates)
            {
                if (column >= firstColumn && column <= lastColumn)
                    buffer[count++] = column;
            }
#pragma warning restore S3267

            var columns = buffer.AsSpan(0, count);
            columns.Sort();
            if (descending)
                columns.Reverse();

            return ScanColumns(range, columns, options, predicate);
        }
        finally
        {
            ArrayPool<int>.Shared.Return(buffer);
        }
    }

    private int ScanColumns(Area range, ReadOnlySpan<int> usedColumns, XLCellsUsedOptions options,
        Func<IXLCell, bool>? predicate)
    {
        foreach (var columnNumber in usedColumns)
        {
            var enumerator = new SlicesEnumerator(new Area(range.FirstPoint.Row, columnNumber, range.LastPoint.Row, columnNumber), this);
            while (enumerator.MoveNext())
            {
                var cell = new XLCell(_ws, enumerator.Current);
                if (!cell.IsEmpty(options) &&
                    (predicate == null || predicate(cell)))
                {
                    return enumerator.Current.Column;
                }
            }
        }

        return 0;
    }

    private int FindUsedRow(Area searchRange, XLCellsUsedOptions options, Func<IXLCell, bool>? predicate, bool reverse)
    {
        var enumerator = new SlicesEnumerator(searchRange, this, reverse);

        while (enumerator.MoveNext())
        {
            var cellAddress = enumerator.Current;
            var cell = GetCell(cellAddress);
            if (!cell.IsEmpty(options)
                && (predicate == null || predicate(cell)))
                return cellAddress.Row;
        }

        return 0;
    }

    private bool IsUsed(Point address)
    {
        // This is different from XLCellUsedOptions, which uses a business logic (e.g. empty string is considered not-used).
        // Here, we ask whether any slice contains a used elements which might differ from cell used logic.
#pragma warning disable S3267 // Four slices, asked per cell: Where would allocate an iterator and a closure per call
        foreach (var slice in _slices)
        {
            if (slice.IsUsed(address))
                return true;
        }
#pragma warning restore S3267

        return false;
    }

    internal void SwapCellsContent(Point sp1, Point sp2)
    {
        ValueSlice.Swap(sp1, sp2);
        FormulaSlice.Swap(sp1, sp2);
        StyleSlice.Swap(sp1, sp2);
        MiscSlice.Swap(sp1, sp2);
    }

    /// <summary>
    /// Gets used points in the range.
    /// </summary>
    internal SlicesEnumerator ForValuesAndFormulas(Area range)
    {
        var valueEnumerator = ValueSlice.GetEnumerator(range);
        var formulaEnumerator = FormulaSlice.GetEnumerator(range);
        return new SlicesEnumerator(false, valueEnumerator, formulaEnumerator);
    }

    /// <summary>
    /// Enumerator that combines several other slice enumerators and enumerates
    /// <see cref="Point"/> in any of them.
    /// </summary>
    internal struct SlicesEnumerator
    {
        // Fixed-size array to avoid List overhead; max 4 slices (value, formula, style, misc).
        private readonly IEnumerator<Point>[] _enumerators;
        private int _count;
        private readonly bool _reverse;

        public SlicesEnumerator(Area range, XLCellsCollection cellsCollection, bool reverse = false)
            : this(
                reverse,
                cellsCollection.ValueSlice.GetEnumerator(range, reverse),
                cellsCollection.FormulaSlice.GetEnumerator(range, reverse),
                cellsCollection.StyleSlice.GetEnumerator(range, reverse),
                cellsCollection.MiscSlice.GetEnumerator(range, reverse))
        {
        }

        public SlicesEnumerator(bool reverse, params IEnumerator<Point>[] enumerators)
        {
            Current = new Point(1, 1);
            _reverse = reverse;
            _enumerators = new IEnumerator<Point>[enumerators.Length];
            _count = 0;

            // MoveNext both tests and advances, so this reads as a filter but is not one: expressing it
            // as Where would hide a side effect inside a predicate, which is worse than the loop.
#pragma warning disable S3267
            foreach (var enumerator in enumerators)
            {
                if (enumerator.MoveNext())
                    _enumerators[_count++] = enumerator;
            }
#pragma warning restore S3267
        }

        public Point Current { get; private set; }

        public bool MoveNext()
        {
            if (_count == 0)
                return false;

            var current = FindNextPoint();
            Current = current;
            AdvanceMatchingEnumerators(current);
            return true;
        }

        private Point FindNextPoint()
        {
            var current = _enumerators[0].Current;
            for (var i = 1; i < _count; i++)
            {
                var candidate = _enumerators[i].Current;
                var comparison = candidate.CompareTo(current);
                if (_reverse ? comparison > 0 : comparison < 0)
                    current = candidate;
            }

            return current;
        }

        private void AdvanceMatchingEnumerators(Point current)
        {
            for (var i = _count - 1; i >= 0; --i)
            {
                var enumerator = _enumerators[i];
                if (enumerator.Current == current && !enumerator.MoveNext())
                {
                    _count--;
                    if (i < _count)
                        _enumerators[i] = _enumerators[_count];
                }
            }
        }
    }

    void IWorkbookListener.OnSheetRenamed(string oldSheetName, string newSheetName)
    {
        RenameSheetInFormulas(oldSheetName, newSheetName);
    }

    /// <summary>
    /// Rewrites every reference to <paramref name="oldSheetName"/> in the formulas of this collection
    /// to <paramref name="newSheetName"/>. Used both when a sheet is renamed and when one is copied,
    /// where the copy's references to the original must follow the copy.
    /// </summary>
    internal void RenameSheetInFormulas(string oldSheetName, string newSheetName)
    {
        if (XLHelper.SheetComparer.Equals(oldSheetName, newSheetName))
            return;

        using var enumerator = FormulaSlice.GetForwardEnumerator(Area.Full);
        while (enumerator.MoveNext())
        {
            ref readonly var cellFormula = ref enumerator.Current;
            var currentPoint = enumerator.Point;
            if (cellFormula.Type != FormulaType.Normal)
            {
                // Array or data formula. Only change the name once, on the master cell.
                var isMasterCell = cellFormula.Range.FirstPoint == currentPoint;
                if (!isMasterCell)
                {
                    continue;
                }
            }

            cellFormula.RenameSheet(currentPoint, oldSheetName, newSheetName);
        }
    }
}
