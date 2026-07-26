using XLibur.Excel;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;

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

        var pivotCache = wb.PivotCaches.Add(range);
        ws.Cell("A1").Value = "Pastry";

        await Assert.That(pivotCache.FieldNames).IsEquivalentTo(PivotCacheFieldNameOnly, CollectionOrdering.Matching);
    }

    [Test]
    public async Task Refresh_UpdatesFieldNames()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.FirstCell().InsertData(PivotCacheFieldNamePie);

        var pivotCache = wb.PivotCaches.Add(range);
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

        var pivotCache = wb.PivotCaches.Add(range);

        pivotCache.ItemsToRetainPerField = XLItemsToRetain.None;
        pivotCache.SaveSourceData = false;
        pivotCache.RefreshDataOnOpen = true;

        pivotCache.Refresh();

        await Assert.That(pivotCache.ItemsToRetainPerField).IsEqualTo(XLItemsToRetain.None);
        await Assert.That(pivotCache.SaveSourceData).IsEqualTo(false);
        await Assert.That(pivotCache.RefreshDataOnOpen).IsEqualTo(true);
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

            var table = range.CreateTable();

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
