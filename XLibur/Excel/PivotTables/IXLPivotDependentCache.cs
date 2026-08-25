using System.Collections.Generic;

namespace XLibur.Excel;

/// <summary>
/// The part of a slicer cache or a timeline cache that binds it to pivot tables.
/// </summary>
/// <remarks>
/// <para>
/// Slicers and timelines are unrelated types with unrelated caches, and they stay that way — this
/// exists only so <see cref="XLPivotDependentCascade"/> can express "drop this pivot table from the
/// cache, and say whether the cache is now serving nothing" once instead of twice.
/// </para>
/// <para>
/// Writing it twice is what let the two copies be wrong in the same way: both removed only the
/// control whose turn it was when a shared cache ran out, and left every other control bound to that
/// same cache pointing at a part the save then deleted.
/// </para>
/// </remarks>
internal interface IXLPivotDependentCache
{
    /// <summary>The pivot tables this cache drives, resolved against the workbook.</summary>
    List<XLPivotTable> PivotTables { get; }

    /// <summary>Their names as written in the part, which is what the cache XML carries.</summary>
    List<string> PivotTableNames { get; }
}
