using System;
using System.Collections.Generic;
using System.Linq;

namespace XLibur.Excel;

/// <summary>
/// Keeps the controls that filter a pivot table consistent with it.
/// </summary>
/// <remarks>
/// <para>
/// A slicer cache and a timeline cache both name the pivot tables they drive. Deleting one of them
/// and leaving a cache pointing at a pivot table that is no longer there produces a file Excel
/// offers to repair, so the reference has to go with the pivot table. When the deleted pivot table
/// was the last one a cache served, the control has nothing left to filter and goes too — along with
/// its cache part, the workbook's registration of that cache, the <c>#N/A</c> defined name written
/// for it and its drawing anchor.
/// </para>
/// <para>
/// This closes a gap that predates either being modelled: before, deleting a pivot table left the
/// parts untouched in the package, because nothing knew they were connected.
/// </para>
/// </remarks>
internal static class XLPivotDependentCascade
{
    /// <summary>
    /// Drops the deleted pivot table from every slicer and timeline cache that named it, and removes
    /// every control left with nothing to filter.
    /// </summary>
    internal static void OnPivotTableDeleted(XLWorkbook workbook, XLPivotTable pivotTable)
    {
        foreach (var worksheet in workbook.WorksheetsInternal)
        {
            var slicers = worksheet.SlicersInternal;
            RemoveOrphaned(slicers.Items, slicer => slicer.Cache, slicers.Remove, pivotTable);

            var timelines = worksheet.TimelinesInternal;
            RemoveOrphaned(timelines.Items, timeline => timeline.Cache, timelines.Remove, pivotTable);
        }
    }

    /// <summary>
    /// Unbinds the deleted pivot table from each control's cache, then removes every control whose
    /// cache is left serving nothing.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>It takes two passes, because a cache may be shared.</b> More than one slicer may name the
    /// same cache — that is how one set of buttons drives a dashboard from several sheets — and the
    /// reader produces exactly that whenever two <c>x14:slicer</c> elements name one cache.
    /// </para>
    /// <para>
    /// A single pass gets it wrong in a way that is easy to miss. The first control bound to a shared
    /// cache removes the pivot table and sees the cache empty out; the second finds
    /// <see cref="List{T}.Remove"/> returning <c>false</c>, because the pivot table has already gone,
    /// and skips itself. The shared cache part is then deleted once while the second control is still
    /// pointing at it — a dangling reference, which is the failure this whole class exists to
    /// prevent.
    /// </para>
    /// </remarks>
    private static void RemoveOrphaned<TControl>(
        IReadOnlyList<TControl> controls,
        Func<TControl, IXLPivotDependentCache> cacheOf,
        Action<TControl> remove,
        XLPivotTable pivotTable)
    {
        HashSet<IXLPivotDependentCache>? emptied = null;

        foreach (var control in controls)
        {
            var cache = cacheOf(control);
            if (!cache.PivotTables.Remove(pivotTable))
                continue;

            cache.PivotTableNames.RemoveAll(name => XLHelper.NameComparer.Equals(name, pivotTable.Name));

            // Other pivot tables still share this cache, so its controls keep working and only lose
            // one of their connections.
            if (cache.PivotTables.Count == 0)
                (emptied ??= []).Add(cache);
        }

        if (emptied is null)
            return;

        // Every control bound to an emptied cache goes, not only the one whose turn it was when the
        // cache ran out. Materialised first, because removing mutates the collection being walked.
        foreach (var orphaned in controls.Where(control => emptied.Contains(cacheOf(control))).ToList())
            remove(orphaned);
    }
}
