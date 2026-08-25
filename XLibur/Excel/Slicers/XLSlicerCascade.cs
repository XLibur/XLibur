using System.Collections.Generic;

namespace XLibur.Excel;

/// <summary>
/// Keeps slicers consistent with the pivot tables they filter.
/// </summary>
/// <remarks>
/// <para>
/// A slicer cache names the pivot tables it drives. Deleting one of them and leaving the cache
/// pointing at a pivot table that is no longer there produces a file Excel offers to repair, so the
/// reference has to go with the pivot table. When the deleted pivot table was the last one a cache
/// served, the slicer has nothing left to filter and goes too — along with its cache part, the
/// workbook's registration of that cache and the <c>#N/A</c> defined name written for it.
/// </para>
/// <para>
/// This closes a gap that predates slicers being modelled at all: before this, deleting a pivot
/// table left the slicer parts untouched in the package, because nothing knew they were connected.
/// </para>
/// </remarks>
internal static class XLSlicerCascade
{
    /// <summary>
    /// Drops the deleted pivot table from every slicer cache that named it, and removes any slicer
    /// left with nothing to filter.
    /// </summary>
    internal static void OnPivotTableDeleted(XLWorkbook workbook, XLPivotTable pivotTable)
    {
        foreach (var worksheet in workbook.WorksheetsInternal)
        {
            List<XLSlicer>? orphaned = null;

            foreach (var slicer in worksheet.SlicersInternal.Items)
            {
                var cache = slicer.Cache;
                if (!cache.PivotTables.Remove(pivotTable))
                    continue;

                cache.PivotTableNames.RemoveAll(
                    name => XLHelper.NameComparer.Equals(name, pivotTable.Name));

                // Other pivot tables still share this cache, so the slicer keeps working and only
                // loses one of its connections.
                if (cache.PivotTables.Count > 0)
                    continue;

                (orphaned ??= []).Add(slicer);
            }

            if (orphaned is null)
                continue;

            foreach (var slicer in orphaned)
                worksheet.SlicersInternal.Remove(slicer);
        }
    }
}
