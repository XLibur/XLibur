using System.Collections.Generic;
using System.Diagnostics;
using XLibur.Excel.Tables;

namespace XLibur.Excel;

/// <summary>
/// One entry of a slicer cache's item list: an index into the pivot cache field's shared items and
/// whether that item is selected.
/// </summary>
/// <remarks>
/// <c>s</c> is absent on an unselected item, which is the opposite of what the attribute's name
/// suggests to a reader of the schema alone. The fixture settles it:
/// <c>Resource/TryToLoad/SlicersOnPivotAndTable.xlsx</c> marks exactly one item <c>s="1"</c>, and it
/// is exactly the one pivot field item that does <em>not</em> carry <c>h="1"</c>.
/// </remarks>
/// <param name="Index">The index into the pivot cache field's shared items.</param>
/// <param name="Selected">Whether the item is selected in the slicer.</param>
internal readonly record struct XLSlicerCacheItem(uint Index, bool Selected);

/// <summary>
/// A slicer cache: the workbook-level part that binds a slicer to what it filters and remembers the
/// selection.
/// </summary>
/// <remarks>
/// <para>
/// Kept internal. The cache is a level of indirection Excel's file format needs and a caller does
/// not: everything it holds is reached through <see cref="IXLSlicer"/>. It is also the piece that
/// makes a slicer's relationships many-to-many — several pivot tables may share one cache — so
/// exposing it before the write path exists would fix a shape that has not been designed yet.
/// </para>
/// <para>
/// Caches are workbook-scoped and registered twice in <c>xl/workbook.xml</c>'s <c>extLst</c>:
/// <c>x14:slicerCaches</c> for pivot caches, <c>x15:slicerCaches</c> for table caches. Excel also
/// emits a <c>#N/A</c> defined name per cache, named after <see cref="Name"/>.
/// </para>
/// </remarks>
[DebuggerDisplay("{Name} ({SourceKind})")]
internal sealed class XLSlicerCache
{
    internal XLSlicerCache(string name, string sourceName, XLSlicerSourceKind sourceKind)
    {
        Name = name;
        SourceName = sourceName;
        SourceKind = sourceKind;
    }

    /// <summary>
    /// The cache name, for example <c>Slicer_Region</c>. A slicer refers to its cache by this name,
    /// and so does the <c>#N/A</c> defined name Excel writes for it.
    /// </summary>
    internal string Name { get; }

    /// <summary>
    /// The pivot cache field or table column the cache draws its items from.
    /// </summary>
    internal string SourceName { get; }

    internal XLSlicerSourceKind SourceKind { get; }

    /// <summary>
    /// Whether the cache was created through the API rather than read from a package.
    /// </summary>
    internal bool IsNew { get; set; }

    /// <summary>
    /// The id of the cache's part relationship on the workbook part, for finding the part again on
    /// save. Null for a cache that has never been in a package.
    /// </summary>
    internal string? WorkbookRelId { get; set; }

    /// <summary>
    /// The <c>x14:pivotCacheDefinition/@pivotCacheId</c> of the pivot cache this cache reads its
    /// items from, when it is a pivot slicer cache.
    /// </summary>
    /// <remarks>
    /// XLibur does not model this identifier, so it is not what binding resolves on — the pivot
    /// table names in <see cref="PivotTableNames"/> are. It is kept because the write path will
    /// have to reproduce it, and because it disambiguates two pivot tables of the same name.
    /// </remarks>
    internal uint? PivotCacheId { get; set; }

    /// <summary>
    /// The names of the pivot tables the cache drives, as written in the part. A name with no
    /// matching pivot table in the workbook is left out of <see cref="PivotTables"/>.
    /// </summary>
    internal List<string> PivotTableNames { get; } = [];

    /// <summary>
    /// The <c>table/@id</c> of the table this cache filters, when it is a table slicer cache.
    /// </summary>
    internal uint? TableId { get; set; }

    /// <summary>
    /// The <c>tableColumn/@id</c> of the filtered column, which is not its position in the table.
    /// </summary>
    internal uint? TableColumnId { get; set; }

    /// <summary>
    /// The cache's item list, present only when the file records one.
    /// </summary>
    internal List<XLSlicerCacheItem> Items { get; } = [];

    /// <summary>
    /// The pivot tables resolved from <see cref="PivotTableNames"/>.
    /// </summary>
    internal List<XLPivotTable> PivotTables { get; } = [];

    /// <summary>
    /// The pivot cache the item indices point into, resolved from the bound pivot tables.
    /// </summary>
    internal XLPivotCache? PivotCache { get; set; }

    /// <summary>
    /// The table resolved from <see cref="TableId"/>.
    /// </summary>
    internal XLTable? Table { get; set; }

    /// <summary>
    /// The 1-based position of the filtered column within <see cref="Table"/>, resolved from
    /// <see cref="TableColumnId"/> while reading the table part.
    /// </summary>
    /// <remarks>
    /// A column's id is stable across column moves and is therefore not its position, but auto
    /// filter columns are addressed by position — and XLibur does not model column ids, so the
    /// translation has to happen while the table part is still in hand.
    /// </remarks>
    internal int? TableColumnPosition { get; set; }
}
