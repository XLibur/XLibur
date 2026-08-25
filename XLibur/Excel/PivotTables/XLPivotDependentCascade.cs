using System.Collections.Generic;

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
/// parts untouched in the package, because nothing knew they were connected. Timelines were the last
/// instance of that hazard named in <c>docs/round-trip-fidelity.md</c>.
/// </para>
/// </remarks>
internal static class XLPivotDependentCascade
{
    /// <summary>
    /// Drops the deleted pivot table from every slicer and timeline cache that named it, and removes
    /// any control left with nothing to filter.
    /// </summary>
    internal static void OnPivotTableDeleted(XLWorkbook workbook, XLPivotTable pivotTable)
    {
        foreach (var worksheet in workbook.WorksheetsInternal)
        {
            RemoveOrphanedSlicers(worksheet, pivotTable);
            RemoveOrphanedTimelines(worksheet, pivotTable);
        }
    }

    private static void RemoveOrphanedSlicers(XLWorksheet worksheet, XLPivotTable pivotTable)
    {
        List<XLSlicer>? orphaned = null;

        foreach (var slicer in worksheet.SlicersInternal.Items)
        {
            var cache = slicer.Cache;
            if (!cache.PivotTables.Remove(pivotTable))
                continue;

            cache.PivotTableNames.RemoveAll(name => XLHelper.NameComparer.Equals(name, pivotTable.Name));

            // Other pivot tables still share this cache, so the slicer keeps working and only loses
            // one of its connections.
            if (cache.PivotTables.Count > 0)
                continue;

            (orphaned ??= []).Add(slicer);
        }

        if (orphaned is null)
            return;

        foreach (var slicer in orphaned)
            worksheet.SlicersInternal.Remove(slicer);
    }

    private static void RemoveOrphanedTimelines(XLWorksheet worksheet, XLPivotTable pivotTable)
    {
        List<XLTimeline>? orphaned = null;

        foreach (var timeline in worksheet.TimelinesInternal.Items)
        {
            var cache = timeline.Cache;
            if (!cache.PivotTables.Remove(pivotTable))
                continue;

            cache.PivotTableNames.RemoveAll(name => XLHelper.NameComparer.Equals(name, pivotTable.Name));

            if (cache.PivotTables.Count > 0)
                continue;

            (orphaned ??= []).Add(timeline);
        }

        if (orphaned is null)
            return;

        foreach (var timeline in orphaned)
            worksheet.TimelinesInternal.Remove(timeline);
    }
}
