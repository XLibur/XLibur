using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace XLibur.Excel;

/// <summary>
/// The slicers drawn on a worksheet.
/// </summary>
/// <remarks>
/// The worksheet owns its slicers: this is the collection a slicer is added to and removed from.
/// What a slicer filters is a separate relationship held by its cache, exposed as a view on
/// <see cref="IXLPivotTable.Slicers"/>.
/// </remarks>
public interface IXLSlicers : IEnumerable<IXLSlicer>
{
    /// <summary>
    /// The number of slicers on the worksheet.
    /// </summary>
    int Count { get; }

    /// <summary>
    /// The slicer with the given <see cref="IXLSlicer.Name"/>.
    /// </summary>
    /// <param name="name">The slicer's internal name, which is not its caption.</param>
    /// <exception cref="System.Collections.Generic.KeyNotFoundException">
    /// The worksheet has no slicer with that name.
    /// </exception>
    IXLSlicer Slicer(string name);

    /// <summary>
    /// Finds the slicer with the given <see cref="IXLSlicer.Name"/>.
    /// </summary>
    /// <param name="name">The slicer's internal name, which is not its caption.</param>
    /// <param name="slicer">The slicer, when one was found.</param>
    /// <returns>Whether a slicer with that name is on the worksheet.</returns>
    bool TryGetSlicer(string name, [NotNullWhen(true)] out IXLSlicer? slicer);

    /// <summary>
    /// Adds a slicer that filters a pivot table on one of its cache fields.
    /// </summary>
    /// <param name="pivotTable">The pivot table to filter. It need not be on this worksheet.</param>
    /// <param name="fieldName">The pivot cache field the buttons are drawn from.</param>
    /// <returns>The new slicer, starting with every item selected and so filtering nothing.</returns>
    /// <exception cref="System.ArgumentException">
    /// The pivot table's cache has no field of that name.
    /// </exception>
    /// <remarks>
    /// The slicer is placed to the right of the pivot table. Use <see cref="IXLSlicer.Position"/> to
    /// move it.
    /// </remarks>
    IXLSlicer Add(IXLPivotTable pivotTable, string fieldName);

    /// <summary>
    /// Adds a slicer that filters a table on one of its columns.
    /// </summary>
    /// <param name="table">The table to filter. It need not be on this worksheet.</param>
    /// <param name="columnName">The column the buttons are drawn from.</param>
    /// <returns>The new slicer, filtering nothing until the table's auto filter says otherwise.</returns>
    /// <exception cref="System.ArgumentException">The table has no column of that name.</exception>
    /// <remarks>
    /// The slicer is placed to the right of the table. Use <see cref="IXLSlicer.Position"/> to move
    /// it.
    /// </remarks>
    IXLSlicer Add(IXLTable table, string columnName);
}
