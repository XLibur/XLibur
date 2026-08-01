using System.Collections.Generic;
using XLibur.Excel.Exceptions;

namespace XLibur.Excel;

/// <summary>
/// A cache of pivot data - essentially a collection of fields and their values that can be
/// displayed by a <see cref="IXLPivotTable"/>. Data for the cache are retrieved from
/// an area (a table or a range). The pivot cache data are <strong>cached</strong>, i.e.
/// the data in the source are not immediately updated once the data in a worksheet change.
/// </summary>
public interface IXLPivotCache
{
    /// <summary>
    /// Get names of all fields in the source, in left to right order. Every field name is unique.
    /// </summary>
    /// <remarks>
    /// The field names are case insensitive. The field names of the cached
    /// source might differ from actual names of the columns
    /// in the data cells.
    /// </remarks>
    IReadOnlyList<string> FieldNames { get; }

    /// <summary>
    /// Gets the number of unused items in shared items to allow before discarding unused items.
    /// </summary>
    /// <remarks>
    /// Shared items are distinct values of a source field values. Updating them can be expensive
    /// and this controls, when should the cache be updated. Application-dependent attribute.
    /// </remarks>
    /// <value>Default value is <see cref="XLItemsToRetain.Automatic"/>.</value>
    XLItemsToRetain ItemsToRetainPerField { get; set; }

    /// <summary>
    /// Will Excel refresh the cache when it opens the workbook.
    /// </summary>
    /// <value>Default value is <c>false</c>.</value>
    bool RefreshDataOnOpen { get; set; }

    /// <summary>
    /// Should the cached values of the pivot source be saved into the workbook file?
    /// If source data are not saved, they will have to be refreshed from the source
    /// reference which might cause a change in the table values.
    /// </summary>
    /// <value>Default value is <c>true</c>.</value>
    bool SaveSourceData { get; set; }

    /// <summary>
    /// What kind of source this cache reads from.
    /// </summary>
    /// <remarks>
    /// Check this before <see cref="SourceRange"/> or <see cref="SourceWorksheet"/>: those are
    /// null both for a source XLibur cannot read at all and for one it can read but that no longer
    /// resolves, and only this tells the two apart.
    /// </remarks>
    XLPivotSourceKind SourceKind { get; }

    /// <summary>
    /// The range this cache reads from. Non-null only when <see cref="SourceKind"/> is
    /// <see cref="XLPivotSourceKind.Range"/> and that range's sheet still exists.
    /// </summary>
    IXLRange? SourceRange { get; }

    /// <summary>
    /// The table or defined name this cache reads from. Non-null exactly when
    /// <see cref="SourceKind"/> is <see cref="XLPivotSourceKind.Name"/> — this is the name the
    /// file recorded, whether or not it still resolves to anything.
    /// </summary>
    string? SourceName { get; }

    /// <summary>
    /// The worksheet this cache reads from, resolved through the table or defined name when
    /// <see cref="SourceKind"/> is <see cref="XLPivotSourceKind.Name"/>. Null when the source does
    /// not resolve — a deleted name, a missing sheet — or when the kind is one XLibur cannot read.
    /// </summary>
    IXLWorksheet? SourceWorksheet { get; }

    /// <summary>
    /// Refresh data in the pivot source from the source reference data.
    /// </summary>
    /// <exception cref="InvalidReferenceException">The data source for the pivot table can't be found.</exception>
    IXLPivotCache Refresh();

    /// <summary>
    /// Re-points the cache at <paramref name="range"/>, making <see cref="SourceKind"/>
    /// <see cref="XLPivotSourceKind.Range"/> whatever it was before. Does not refresh — call
    /// <see cref="Refresh"/> to re-read the records.
    /// </summary>
    /// <param name="range">The range to read from. Must belong to a worksheet.</param>
    /// <exception cref="System.ArgumentNullException"><paramref name="range"/> is null.</exception>
    /// <exception cref="System.ArgumentException"><paramref name="range"/> has no worksheet.</exception>
    IXLPivotCache SetSourceRange(IXLRange range);

    /// <inheritdoc cref="ItemsToRetainPerField"/>
    IXLPivotCache SetItemsToRetainPerField(XLItemsToRetain value);

    /// <inheritdoc cref="RefreshDataOnOpen"/>
    /// <remarks>Sets the value to <c>true</c>.</remarks>
    IXLPivotCache SetRefreshDataOnOpen();

    /// <inheritdoc cref="RefreshDataOnOpen"/>
    IXLPivotCache SetRefreshDataOnOpen(bool value);

    /// <inheritdoc cref="SaveSourceData"/>
    /// <remarks>Sets the value to <c>true</c>.</remarks>
    IXLPivotCache SetSaveSourceData();

    /// <inheritdoc cref="SaveSourceData"/>
    IXLPivotCache SetSaveSourceData(bool value);
}
