using System.Collections.Generic;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel;

/// <summary>
/// An interface for methods of <see cref="Slice{TElement}"/> without specified type of element.
/// </summary>
internal interface ISlice
{
    /// <summary>
    /// Is at least one cell in the slice used?
    /// </summary>
    bool IsEmpty { get; }

    /// <summary>
    /// A monotonically-increasing counter that increments on every write to the slice.
    /// Used by callers that cache derived values (e.g., an aggregated set of used columns
    /// across multiple slices) to detect when their cache is stale. The absolute value is
    /// not meaningful; only inequality between two snapshots indicates that the slice
    /// changed in between.
    /// </summary>
    int Version { get; }

    /// <summary>
    /// Get the maximum used column in the slice or 0, if no column is used.
    /// </summary>
    int MaxColumn { get; }

    /// <summary>
    /// Get the maximum used row in the slice or 0, if no row is used.
    /// </summary>
    int MaxRow { get; }

    /// <summary>
    /// Remove a set of whole rows and close the gaps, in one pass over the slice.
    /// </summary>
    /// <remarks>
    /// The whole-row, whole-deletion counterpart of <see cref="DeleteAreaAndShiftUp"/>. Because a row
    /// is one stored value rather than a run of cells, and because the entire deletion is applied at
    /// once, this costs one operation per used row instead of one per cell below each deleted row.
    /// </remarks>
    void DeleteRowsAndCompact(XLRowDeletionMap map);

    /// <summary>
    /// A set of columns that have at least one used cell. The order of columns is non-deterministic.
    /// </summary>
    Dictionary<int, int>.KeyCollection UsedColumns { get; }

    /// <summary>
    /// A set of rows that have at least one used cell. The order of rows is non-deterministic.
    /// </summary>
    IEnumerable<int> UsedRows { get; }

    /// <summary>
    /// Clear all values in the range and mark them as unused.
    /// </summary>
    void Clear(Area range);

    /// <summary>
    /// Clear all values in the <paramref name="rangeToDelete"/> and shift all values right of the deleted area to the deleted place.
    /// </summary>
    void DeleteAreaAndShiftLeft(Area rangeToDelete);

    /// <summary>
    /// Clear all values in the <paramref name="rangeToDelete"/> and shift all values below the deleted area to the deleted place.
    /// </summary>
    void DeleteAreaAndShiftUp(Area rangeToDelete);

    /// <summary>
    /// Get all used points in a slice.
    /// </summary>
    /// <param name="range">Range to iterate over.</param>
    /// <param name="reverse"><c>false</c> = left to right, top to bottom. <c>true</c> = right to left, bottom to top.</param>
    IEnumerator<Point> GetEnumerator(Area range, bool reverse = false);

    /// <summary>
    /// Shift all values at the <paramref name="range"/> and all cells below it
    /// down by <see cref="Area.Height"/> of the <paramref name="range"/>.
    /// The insert area is cleared.
    /// </summary>
    void InsertAreaAndShiftDown(Area range);

    /// <summary>
    /// Shift all values at the <paramref name="range"/> and all cells right of it
    /// to the right by <see cref="Area.Width"/> of the <paramref name="range"/>.
    /// The insert area is cleared.
    /// </summary>
    void InsertAreaAndShiftRight(Area range);

    /// <summary>
    /// Does a slice contain a non-default value at a specified point?
    /// </summary>
    bool IsUsed(Point address);

    /// <summary>
    /// Swap content of two points.
    /// </summary>
    void Swap(Point sp1, Point sp2);
}
