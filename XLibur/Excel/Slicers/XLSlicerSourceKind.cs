namespace XLibur.Excel;

/// <summary>
/// What a slicer filters.
/// </summary>
/// <remarks>
/// Excel binds a slicer through its cache, and the two bindings are unrelated to each other: a pivot
/// slicer's cache carries a <c>tabular</c> element naming the pivot cache and the pivot tables it
/// drives, while a table slicer's cache carries an <c>x15:tableSlicerCache</c> extension naming a
/// table and one of its columns.
/// </remarks>
public enum XLSlicerSourceKind
{
    /// <summary>
    /// The slicer filters one or more pivot tables that share a pivot cache.
    /// </summary>
    PivotTable,

    /// <summary>
    /// The slicer filters a single column of a table, through that table's auto filter.
    /// </summary>
    Table,
}
