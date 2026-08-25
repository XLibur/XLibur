using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace XLibur.Excel;

/// <summary>
/// A timeline cache: the workbook-level part that binds a timeline to the pivot tables it filters
/// and remembers its date range.
/// </summary>
/// <remarks>
/// <para>
/// Kept internal, as <see cref="XLSlicerCache"/> is. The cache is a level of indirection Excel's
/// file format needs and a caller does not; everything it holds is reached through
/// <see cref="IXLTimeline"/>.
/// </para>
/// <para>
/// The binding is <c>x15:state/@pivotCacheId</c>, which points at the identifier the pivot cache
/// carries in its own <c>x14:pivotCacheDefinition</c> extension — the same identifier slicer caches
/// quote. Caches are registered in <c>xl/workbook.xml</c>'s <c>extLst</c> under
/// <c>x15:timelineCacheRefs</c>, and Excel writes a <c>#N/A</c> defined name per cache.
/// </para>
/// </remarks>
[DebuggerDisplay("{Name} ({SourceName})")]
internal sealed class XLTimelineCache
{
    internal XLTimelineCache(string name, string sourceName)
    {
        Name = name;
        SourceName = sourceName;
    }

    /// <summary>
    /// The cache name, for example <c>NativeTimeline_Date</c>. A timeline refers to its cache by
    /// this name, and so does the <c>#N/A</c> defined name written for it.
    /// </summary>
    internal string Name { get; }

    /// <summary>The pivot cache field the timeline scrubs.</summary>
    internal string SourceName { get; }

    /// <summary>Whether the cache was created through the API rather than read from a package.</summary>
    internal bool IsNew { get; set; }

    /// <summary>
    /// The id of the cache's part relationship on the workbook part, for finding the part again on
    /// save. Null for a cache that has never been in a package.
    /// </summary>
    internal string? WorkbookRelId { get; set; }

    /// <summary>
    /// The <c>x14:pivotCacheDefinition/@pivotCacheId</c> of the pivot cache this cache scrubs.
    /// </summary>
    internal uint? PivotCacheId { get; set; }

    /// <summary>
    /// The names of the pivot tables the cache drives, as written in the part. A name with no
    /// matching pivot table in the workbook is left out of <see cref="PivotTables"/>.
    /// </summary>
    internal List<string> PivotTableNames { get; } = [];

    /// <summary>The pivot tables resolved from <see cref="PivotTableNames"/>.</summary>
    internal List<XLPivotTable> PivotTables { get; } = [];

    /// <summary>The pivot cache behind those pivot tables.</summary>
    internal XLPivotCache? PivotCache { get; set; }

    /// <summary>
    /// The extent of the scrubber. Nullable because <c>x15:bounds</c> is an optional child of
    /// <c>x15:state</c>; a file that omits it reports nothing rather than a fabricated date.
    /// </summary>
    internal DateTime? BoundsStart { get; set; }

    /// <inheritdoc cref="BoundsStart"/>
    internal DateTime? BoundsEnd { get; set; }

    /// <summary>
    /// The selected range, when the file records one. Read-only in the model: changing it has to
    /// move a <c>dateBetween</c> pivot filter and the pivot field's item visibility with it, and a
    /// timeline whose range disagrees with its pivot table is a broken workbook.
    /// </summary>
    internal DateTime? SelectionStart { get; set; }

    /// <inheritdoc cref="SelectionStart"/>
    internal DateTime? SelectionEnd { get; set; }

    /// <summary>
    /// The raw <c>x15:state/@filterType</c> token, held as a string rather than as the SDK's
    /// <c>EnumValue&lt;PivotFilterValues&gt;</c>.
    /// </summary>
    /// <remarks>
    /// The SDK types this attribute as an enumeration, but an <c>EnumValue</c> preserves an
    /// unrecognised token in its <c>InnerText</c>. Carrying the string is what lets a file written
    /// by a newer Excel round-trip a filter type this build has never heard of. A timeline with no
    /// selection says <c>unknown</c>, which is what a created one starts at.
    /// </remarks>
    internal string FilterType { get; set; } = "unknown";

    /// <summary>
    /// The timeline feature's own version stamps. 6 is what Excel 2013 and later write, and it is
    /// unrelated to the pivot table's <c>createdVersion</c>.
    /// </summary>
    internal uint MinimalRefreshVersion { get; set; } = 6;

    /// <inheritdoc cref="MinimalRefreshVersion"/>
    internal uint LastRefreshVersion { get; set; } = 6;
}
