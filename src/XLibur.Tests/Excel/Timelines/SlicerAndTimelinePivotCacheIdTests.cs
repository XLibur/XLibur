using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Timelines;

/// <summary>
/// The pivot cache identifier slicer caches and timeline caches both quote, allocated by one
/// allocator for both.
/// </summary>
public class SlicerAndTimelinePivotCacheIdTests
{
    [Test]
    public async Task A_slicer_and_a_timeline_on_different_pivot_caches_get_distinct_pivot_cache_ids()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var data = AddData(wb);
            var pivotSheet = wb.AddWorksheet("Pivot");

            // Two source ranges, so two pivot caches.
            var first = pivotSheet.PivotTables.Add("First", pivotSheet.Cell("A3"), data.Range("A1:C25"));
            first.RowLabels.Add("Region");
            var second = pivotSheet.PivotTables.Add("Second", pivotSheet.Cell("A30"), data.Range("A1:C13"));
            second.RowLabels.Add("Region");

            pivotSheet.Slicers.Add(first, "Region");
            pivotSheet.Timelines.Add(second, "Date");
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var ids = reloaded.PivotCachesInternal.Cast<XLPivotCache>().Select(c => c.PivotCacheId).ToList();

        await Assert.That(ids.Count).IsEqualTo(2);
        await Assert.That(ids).DoesNotContain((uint?)null);
        await Assert.That(ids.Distinct().Count()).IsEqualTo(2);

        // And each control's cache still quotes the pivot cache of the table it filters.
        var sheet = (XLWorksheet)reloaded.Worksheet("Pivot");
        var slicerCache = sheet.SlicersInternal.Items.Single().Cache;
        var timelineCache = sheet.TimelinesInternal.Items.Single().Cache;

        await Assert.That(slicerCache.PivotCacheId).IsEqualTo(slicerCache.PivotCache!.PivotCacheId);
        await Assert.That(timelineCache.PivotCacheId).IsEqualTo(timelineCache.PivotCache!.PivotCacheId);
        await Assert.That(slicerCache.PivotCacheId).IsNotEqualTo(timelineCache.PivotCacheId);
    }

    [Test]
    public async Task A_slicer_and_a_timeline_on_one_pivot_cache_share_its_id()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var data = AddData(wb);
            var pivotSheet = wb.AddWorksheet("Pivot");
            var pivotTable = pivotSheet.PivotTables.Add("Sales", pivotSheet.Cell("A3"), data.Range("A1:C25"));
            pivotTable.RowLabels.Add("Region");

            pivotSheet.Slicers.Add(pivotTable, "Region");
            pivotSheet.Timelines.Add(pivotTable, "Date");
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var sheet = (XLWorksheet)reloaded.Worksheet("Pivot");
        var slicerCache = sheet.SlicersInternal.Items.Single().Cache;
        var timelineCache = sheet.TimelinesInternal.Items.Single().Cache;

        await Assert.That(slicerCache.PivotCacheId).IsNotNull();
        await Assert.That(timelineCache.PivotCacheId).IsEqualTo(slicerCache.PivotCacheId);
    }

    private static IXLWorksheet AddData(XLWorkbook wb)
    {
        var data = wb.AddWorksheet("Data");
        data.Cell("A1").Value = "Date";
        data.Cell("B1").Value = "Region";
        data.Cell("C1").Value = "Amount";

        var start = new DateTime(2024, 1, 15);
        for (var i = 0; i < 24; i++)
        {
            data.Cell(i + 2, 1).Value = start.AddDays(i * 11);
            data.Cell(i + 2, 2).Value = i % 2 == 0 ? "North" : "South";
            data.Cell(i + 2, 3).Value = 100 + (i * 7);
        }

        return data;
    }
}
