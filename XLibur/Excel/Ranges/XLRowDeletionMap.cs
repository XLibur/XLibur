using System;
using System.Collections.Generic;

namespace XLibur.Excel;

/// <summary>
/// Where every row of a sheet ends up after a set of whole rows is deleted in one go.
/// </summary>
/// <remarks>
/// Deleting scattered rows one at a time makes every consumer of a shift — most expensively the
/// formula pass — run once per row. This collapses the whole set into a single mapping so they can
/// run once instead.
/// <para>
/// A reference spanning rows <c>[a, b]</c> becomes <c>[MapFirst(a), MapLast(b)]</c>. The two ends are
/// counted differently on purpose: the top of a range slides up by the rows deleted strictly above it,
/// while the bottom also loses the rows deleted inside it. When the deletion swallows the range
/// entirely the mapped bottom lands above the mapped top, which is how a destroyed reference is
/// detected.
/// </para>
/// </remarks>
internal sealed class XLRowDeletionMap
{
    /// <summary>Deleted row numbers, ascending and distinct.</summary>
    private readonly int[] _deleted;

    private XLRowDeletionMap(int[] deleted)
    {
        _deleted = deleted;
    }

    /// <summary>The number of rows this map deletes.</summary>
    internal int Count => _deleted.Length;

    /// <summary>The first deleted row; the lowest row number any reference can be affected from.</summary>
    internal int FirstDeletedRow => _deleted[0];

    /// <summary>
    /// Builds a map from an arbitrary collection of row numbers. Duplicates collapse and order does
    /// not matter, so callers may hand over rows in whatever order they gathered them.
    /// </summary>
    /// <returns><c>null</c> when nothing would be deleted.</returns>
    internal static XLRowDeletionMap? Create(IEnumerable<int> rows)
    {
        var sorted = new SortedSet<int>(rows);
        if (sorted.Count == 0)
            return null;

        var deleted = new int[sorted.Count];
        sorted.CopyTo(deleted);
        return new XLRowDeletionMap(deleted);
    }

    /// <summary>
    /// Splits the deleted rows into contiguous runs, furthest down the sheet first. Consumers that
    /// still work a block at a time need this order: a run must be removed before the runs above it
    /// move, or their row numbers stop meaning what the map says they mean.
    /// </summary>
    internal List<(int FirstRow, int LastRow)> GetRunsBottomUp()
    {
        var runs = new List<(int, int)>();
        for (var i = _deleted.Length - 1; i >= 0; i--)
        {
            var last = _deleted[i];
            while (i > 0 && _deleted[i - 1] == _deleted[i] - 1)
                i--;

            runs.Add((_deleted[i], last));
        }

        return runs;
    }

    /// <summary>Whether this row is one of the deleted ones.</summary>
    internal bool IsDeleted(int row) => Array.BinarySearch(_deleted, row) >= 0;

    /// <summary>
    /// Where the top of a range lands: its own row less the rows deleted strictly above it. A row that
    /// is itself deleted maps to the position its first surviving successor takes.
    /// </summary>
    internal int MapFirst(int row) => row - CountBelow(row);

    /// <summary>
    /// Where the bottom of a range lands: its own row less every deleted row at or above it, so rows
    /// deleted from inside the range shorten it.
    /// </summary>
    internal int MapLast(int row) => row - CountAtOrBelow(row);

    /// <summary>Deleted rows strictly less than <paramref name="row"/>.</summary>
    private int CountBelow(int row)
    {
        var index = Array.BinarySearch(_deleted, row);
        return index >= 0 ? index : ~index;
    }

    /// <summary>Deleted rows less than or equal to <paramref name="row"/>.</summary>
    private int CountAtOrBelow(int row)
    {
        var index = Array.BinarySearch(_deleted, row);
        return index >= 0 ? index + 1 : ~index;
    }
}
