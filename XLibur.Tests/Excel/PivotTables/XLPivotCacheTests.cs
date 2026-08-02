using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TUnit.Assertions.Enums;
using XLibur.Excel;

namespace XLibur.Tests.Excel.PivotTables;

public class XLPivotCacheTests
{
    private static readonly string[] PivotCacheFieldNamePie = ["Name", "Pie"];
    private static readonly string[] PivotCacheFieldNameOnly = ["Name"];
    private static readonly string[] PivotCacheFieldPastry = ["Pastry"];

    [Test]
    public async Task FieldNames_KeepNamesEvenWhenSourceChange()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.FirstCell().InsertData(PivotCacheFieldNamePie);

        var pivotCache = wb.PivotCaches.Add(range!);
        ws.Cell("A1").Value = "Pastry";

        await Assert.That(pivotCache.FieldNames).IsEquivalentTo(PivotCacheFieldNameOnly, CollectionOrdering.Matching);
    }

    [Test]
    public async Task Refresh_UpdatesFieldNames()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.FirstCell().InsertData(PivotCacheFieldNamePie);

        var pivotCache = wb.PivotCaches.Add(range!);
        ws.Cell("A1").Value = "Pastry";
        pivotCache.Refresh();

        await Assert.That(pivotCache.FieldNames).IsEquivalentTo(PivotCacheFieldPastry, CollectionOrdering.Matching);
    }

    [Test]
    public async Task Refresh_RetainsSetOptions()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.FirstCell().InsertData(PivotCacheFieldNamePie);

        var pivotCache = wb.PivotCaches.Add(range!);

        pivotCache.ItemsToRetainPerField = XLItemsToRetain.None;
        pivotCache.SaveSourceData = false;
        pivotCache.RefreshDataOnOpen = true;

        pivotCache.Refresh();

        await Assert.That(pivotCache.ItemsToRetainPerField).IsEqualTo(XLItemsToRetain.None);
        await Assert.That(pivotCache.SaveSourceData).IsFalse();
        await Assert.That(pivotCache.RefreshDataOnOpen).IsTrue();
    }

    /// <summary>
    /// The <c>cacheId</c> a pivot table part writes is a position in the workbook's
    /// <c>pivotCaches</c> element, not a property of the cache, so the two have to agree.
    /// With more than one cache in play a mix-up points each pivot table at the wrong source
    /// instead of merely writing an odd number.
    /// </summary>
    [Test]
    public async Task SavedPivotTables_ReferenceTheCacheIdTheWorkbookListsForTheirOwnCache()
    {
        using var wb = BuildWorkbookWithTwoCaches();
        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        await AssertCacheIdsAgreeWithWorkbook(ms);
    }

    /// <summary>
    /// Cache ids are assigned per save. Saving twice has to renumber from scratch rather than
    /// continue where the previous save left off.
    /// </summary>
    [Test]
    public async Task SavingTwice_AssignsTheSameCacheIdsAgain()
    {
        using var wb = BuildWorkbookWithTwoCaches();

        using var first = new MemoryStream();
        wb.SaveAs(first);

        using var second = new MemoryStream();
        wb.SaveAs(second);

        await AssertCacheIdsAgreeWithWorkbook(first);
        await AssertCacheIdsAgreeWithWorkbook(second);
        await Assert.That(ReadCacheIds(second)).IsEquivalentTo(ReadCacheIds(first), CollectionOrdering.Matching);
    }

    private static XLWorkbook BuildWorkbookWithTwoCaches()
    {
        var wb = new XLWorkbook();
        var data = wb.AddWorksheet("Data");
        var pastries = data.FirstCell().InsertData(new object[]
        {
            ("Pastry", "Sold"),
            ("Waffle", 3),
            ("Donut", 5),
        });
        var doughs = data.Cell("D1").InsertData(new object[]
        {
            ("Dough", "Batches"),
            ("Puff", 2),
            ("Choux", 7),
        });

        var pivots = wb.AddWorksheet("Pivots");
        var byPastry = pivots.PivotTables.Add("byPastry", pivots.Cell("A1"), pastries!);
        byPastry.RowLabels.Add("Pastry");
        byPastry.Values.Add("Sold");

        var byDough = pivots.PivotTables.Add("byDough", pivots.Cell("F1"), doughs!);
        byDough.RowLabels.Add("Dough");
        byDough.Values.Add("Batches");

        return wb;
    }

    /// <summary>
    /// Every <c>cacheId</c> written by a pivot table part, in part order.
    /// </summary>
    private static List<uint> ReadCacheIds(MemoryStream saved)
    {
        saved.Position = 0;
        using var doc = SpreadsheetDocument.Open(saved, false);
        return doc.WorkbookPart!.WorksheetParts
            .SelectMany(wsp => wsp.GetPartsOfType<PivotTablePart>())
            .Select(part => part.PivotTableDefinition!.CacheId!.Value)
            .OrderBy(id => id)
            .ToList();
    }

    private static async Task AssertCacheIdsAgreeWithWorkbook(MemoryStream saved)
    {
        saved.Position = 0;
        using var doc = SpreadsheetDocument.Open(saved, false);

        var declared = doc.WorkbookPart!.Workbook!.PivotCaches!
            .Elements<PivotCache>()
            .Select(cache => cache.CacheId!.Value)
            .ToList();

        var referenced = doc.WorkbookPart.WorksheetParts
            .SelectMany(wsp => wsp.GetPartsOfType<PivotTablePart>())
            .Select(part => part.PivotTableDefinition!.CacheId!.Value)
            .ToList();

        await Assert.That(declared).IsEquivalentTo(new uint[] { 0, 1 }, CollectionOrdering.Matching);
        await Assert.That(referenced.Count).IsEqualTo(2);

        // Each pivot table must point at a declared cache, and at a different one from the
        // other: the two were built from separate source ranges.
        await Assert.That(referenced.Distinct().Count()).IsEqualTo(2);
        foreach (var cacheId in referenced)
            await Assert.That(declared).Contains(cacheId);
    }

    [Test]
    public async Task Refresh_RenamedFieldIsRemovedFromPivotTable()
    {
        // Pivot table has only field for Pastry, the dough is no longer in the pivot table after refresh
        await TestHelper.CreateAndCompare(wb =>
        {
            var ws = wb.AddWorksheet();
            var range = ws.FirstCell().InsertData(new object[]
            {
                ("Pastry", "Dough"),
                ("Waffles", "Puff")
            });

            var table = range!.CreateTable();

            var pivotTable = ws.PivotTables.Add("pvt", ws.Cell("D1"), table);
            pivotTable.RowLabels.Add("Pastry");
            pivotTable.RowLabels.Add("Dough");
            pivotTable.Values.Add("Pastry").SetSummaryFormula(XLPivotSummary.Count);

            ws.Cell("B1").Value = "Mixture";
            pivotTable.PivotCache.Refresh();
        }, @"Other\PivotTableReferenceFiles\RenamedFieldIsRemovedFromPivotTable-output.xlsx");
    }

    [Test]
    public async Task Preserve_field_statistics_even_without_source_data()
    {
        // Even though the pivot table cache has no records in the workbook, it does contain
        // statistics about each field (e.g. types and min/max values). These are preserved
        // through load/save.
        // The cache fields in the file don't have any shared values or records, only stats,
        // and load/save preserves all Contains* flags and Min/Max values.
        await TestHelper.LoadAndAssert(async wb =>
        {
            await Assert.That(wb.Worksheets.Count).IsGreaterThan(0);
        }, @"Other\PivotTableReferenceFiles\PivotCacheWithoutSourceData-input.xlsx");

        await TestHelper.LoadSaveAndCompare(
            @"Other\PivotTableReferenceFiles\PivotCacheWithoutSourceData-input.xlsx",
            @"Other\PivotTableReferenceFiles\PivotCacheWithoutSourceData-output.xlsx");
    }
}
