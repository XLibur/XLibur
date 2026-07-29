using System;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.Coordinates;

namespace XLibur.Report.Tests.Rewriting;

/// <summary>
/// Pins what a pivot cache's source, and a pivot table's own position, do when the rows around them
/// change.
/// </summary>
/// <remarks>
/// The spec's happy path for pivots is a static pivot over a bound named range: the name grows with
/// the report, so the cache should find the extra rows on its own. These tests establish which parts
/// of that hold and which the rewriter has to supply. ClosedXML.Report's #200 (corrupt output since
/// 2021) and #399 (the static-pivot regression) are what makes getting this right worth the care.
/// </remarks>
public class PivotMechanicsCharacterizationTests
{
    /// <summary>
    /// A <c>Data</c> sheet with a heading row and two data rows in A1:B3, and a pivot table on a
    /// <c>Pivot</c> sheet over whatever <paramref name="source"/> returns.
    /// </summary>
    private static (XLWorkbook Workbook, IXLPivotTable Pivot) Workbook(Func<IXLWorksheet, IXLRange> source)
    {
        var workbook = new XLWorkbook();
        var data = workbook.AddWorksheet("Data");

        data.Cell("A1").Value = "Product";
        data.Cell("B1").Value = "Quantity";
        data.Cell("A2").Value = "Product 1";
        data.Cell("B2").Value = 1;
        data.Cell("A3").Value = "Product 2";
        data.Cell("B3").Value = 2;

        var target = workbook.AddWorksheet("Pivot");
        var pivot = target.PivotTables.Add("pvt", target.Cell("A1"), source(data));
        pivot.RowLabels.Add("Product");
        pivot.Values.Add("Quantity");

        return (workbook, pivot);
    }

    /// <summary>
    /// Adds two data rows <em>inside</em> the A1:B3 source, which is where a report's expansion puts
    /// them: the expander inserts below the last data row, and the options row below that is still
    /// inside the bound range.
    /// </summary>
    private static void AddTwoRowsInsideTheSource(XLWorkbook workbook)
    {
        var data = workbook.Worksheet("Data");

        data.Row(2).InsertRowsBelow(2);
        data.Cell("A3").Value = "Product 3";
        data.Cell("B3").Value = 3;
        data.Cell("A4").Value = "Product 4";
        data.Cell("B4").Value = 4;
    }

    private static int RecordCount(IXLPivotTable pivot) => ((XLPivotCache)pivot.PivotCache).RecordCount;

    /// <summary>
    /// An area-sourced cache does not follow rows inserted into its source: the source is a plain
    /// sheet-plus-rectangle value, not a live range, so nothing shifts or stretches it. This is the
    /// first thing the rewriter has to fix.
    /// </summary>
    [Test]
    public async Task AnAreaSourcedCacheDoesNotGrowWithInsertedRows()
    {
        var (workbook, pivot) = Workbook(sheet => sheet.Range("A1:B3"));
        using var _ = workbook;

        await Assert.That(RecordCount(pivot)).IsEqualTo(2);

        AddTwoRowsInsideTheSource(workbook);
        pivot.PivotCache.Refresh();

        await Assert.That(RecordCount(pivot)).IsEqualTo(2);
    }

    /// <summary>
    /// Re-pointing the source and refreshing picks the new rows up, which is the whole of the
    /// mechanism the rewriter needs.
    /// </summary>
    [Test]
    public async Task RePointingAnAreaSourcedCacheAndRefreshingPicksUpTheNewRows()
    {
        var (workbook, pivot) = Workbook(sheet => sheet.Range("A1:B3"));
        using var _ = workbook;

        AddTwoRowsInsideTheSource(workbook);

        var cache = (XLPivotCache)pivot.PivotCache;
        cache.Source = new XLPivotSourceReference(SheetArea.From(workbook.Worksheet("Data").Range("A1:B5")));
        cache.Refresh();

        await Assert.That(RecordCount(pivot)).IsEqualTo(4);
    }

    /// <summary>
    /// A name-sourced cache follows on its own once refreshed, because the defined name is a live
    /// range: inserting rows inside it stretches it, and the cache resolves the name afresh on every
    /// refresh. This is why the spec makes "static pivot over a bound named range" the happy path —
    /// but note the refresh is still needed, which is what upstream #399 lost.
    /// </summary>
    [Test]
    public async Task ANameSourcedCacheGrowsWithItsNameWhenRefreshed()
    {
        var (workbook, pivot) = Workbook(sheet =>
        {
            var range = sheet.Range("A1:B3");
            sheet.Workbook.DefinedNames.Add("SourceData", range);
            return range;
        });
        using var _ = workbook;

        // The cache was created from the range, so it holds an area source; naming the range in the
        // pivot's source box is what makes it a name source, and this is that.
        var cache = (XLPivotCache)pivot.PivotCache;
        cache.Source = new XLPivotSourceReference("SourceData");
        cache.Refresh();
        await Assert.That(RecordCount(pivot)).IsEqualTo(2);

        AddTwoRowsInsideTheSource(workbook);
        cache.Refresh();

        await Assert.That(RecordCount(pivot)).IsEqualTo(4);
        await Assert.That(workbook.DefinedName("SourceData")!.RefersTo).Contains("$B$5");
    }

    /// <summary>A table-sourced cache follows too: inserting rows inside a table grows the table.</summary>
    [Test]
    public async Task ATableSourcedCacheGrowsWithItsTableWhenRefreshed()
    {
        var (workbook, pivot) = Workbook(sheet => sheet.Range("A1:B3").CreateTable("SourceTable"));
        using var _ = workbook;

        await Assert.That(RecordCount(pivot)).IsEqualTo(2);

        AddTwoRowsInsideTheSource(workbook);
        pivot.PivotCache.Refresh();

        await Assert.That(RecordCount(pivot)).IsEqualTo(4);
    }

    /// <summary>
    /// A pivot table does <em>not</em> move when rows are inserted above it, because its position is
    /// a plain rectangle rather than a live range. So a pivot below a bound range would be written
    /// over by the rows the range generated — the second thing the rewriter has to fix, and one the
    /// spec did not anticipate.
    /// </summary>
    [Test]
    public async Task APivotTableDoesNotMoveWhenRowsAreInsertedAboveIt()
    {
        var (workbook, pivot) = Workbook(sheet => sheet.Range("A1:B3"));
        using var _ = workbook;

        workbook.Worksheet("Pivot").Row(1).InsertRowsAbove(4);

        await Assert.That(pivot.TargetCell.Address.RowNumber).IsEqualTo(1);
    }

    /// <summary><see cref="IXLPivotTable.TargetCell"/> is settable, so moving one is a plain assignment.</summary>
    [Test]
    public async Task APivotTableCanBeMovedByAssigningItsTargetCell()
    {
        var (workbook, pivot) = Workbook(sheet => sheet.Range("A1:B3"));
        using var _ = workbook;

        var target = workbook.Worksheet("Pivot");
        pivot.TargetCell = target.Cell("A5");

        await Assert.That(pivot.TargetCell.Address.RowNumber).IsEqualTo(5);
    }
}
