using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Tests.Utils;
using SaveOptions = XLibur.Excel.SaveOptions;

namespace XLibur.Tests.Excel.Saving;

/// <summary>
/// A sheet holding a pivot table keeps its part, or its cells, on the first save of a loaded
/// workbook when the sheet has not changed (#702, #721).
/// </summary>
/// <remarks>
/// XLibur does not lay a pivot table out: the load reads no pivot values into the cells, and the
/// save writes a change to a pivot table into the pivot table's own part. So no change to a pivot
/// table changes the cells a save would write, and each test of such a change checks that the cells
/// are still kept and that the change itself is saved.
/// </remarks>
public class KeepPivotTableSheetTests
{
    private const string Sheet1 = "xl/worksheets/sheet1.xml";

    [Test]
    public async Task An_unchanged_sheet_with_a_pivot_table_keeps_its_part()
    {
        using var source = Source();

        using var saved = source.LoadAndSave();

        await Assert.That(saved.PartBytes(Sheet1).SequenceEqual(source.PartBytes(Sheet1))).IsTrue();
        await AssertValues(saved);
        await Assert.That(PivotTableNames(saved)).IsEquivalentTo(["pt"]);
    }

    [Test]
    public async Task A_sheet_with_a_pivot_table_keeps_its_cells_when_only_its_markup_changes()
    {
        using var source = Source();

        using var saved = source.LoadAndSave(wb => wb.Worksheet("Data").Column("H").Width = 30);

        await Assert.That(saved.IsKeptWhole(Sheet1)).IsFalse();
        await Assert.That(saved.HasKeptCells(Sheet1)).IsTrue();
        await AssertValues(saved);
    }

    [Test]
    public async Task An_edited_sheet_with_a_pivot_table_is_written_from_the_model()
    {
        using var source = Source();

        using var saved = source.LoadAndSave(wb => wb.Worksheet("Data").Cell("A5").Value = "c");

        await Assert.That(saved.IsWrittenFromModel(Sheet1)).IsTrue();
    }

    [Test]
    public async Task Adding_a_field_to_the_pivot_table_keeps_the_cells_and_saves_the_field()
    {
        using var source = Source();

        using var saved = source.LoadAndSave(wb =>
            wb.Worksheet("Data").PivotTable("pt").ColumnLabels.Add("Region"));

        await Assert.That(saved.HasKeptCells(Sheet1)).IsTrue();
        await AssertValues(saved);
        using var reloaded = new XLWorkbook(saved);
        var columnLabels = reloaded.Worksheet("Data").PivotTable("pt").ColumnLabels.Select(f => f.SourceName);
        await Assert.That(columnLabels).Contains("Region");
    }

    [Test]
    public async Task Deleting_the_pivot_table_keeps_the_cells_and_saves_the_deletion()
    {
        using var source = Source();

        using var saved = source.LoadAndSave(wb => wb.Worksheet("Data").PivotTables.Delete("pt"));

        await Assert.That(saved.HasKeptCells(Sheet1)).IsTrue();
        await AssertValues(saved);
        await Assert.That(PivotTableNames(saved)).IsEmpty();
    }

    [Test]
    public async Task Adding_a_pivot_table_keeps_the_cells_and_saves_the_pivot_table()
    {
        using var source = Source();

        using var saved = source.LoadAndSave(wb =>
        {
            var ws = wb.Worksheet("Data");
            var added = ws.PivotTables.Add("added", ws.Cell("J1"), ws.Range("A1:C3"));
            added.RowLabels.Add("Region");
            added.Values.Add("Amount");
        });

        await Assert.That(saved.HasKeptCells(Sheet1)).IsTrue();
        await AssertValues(saved);
        await Assert.That(PivotTableNames(saved)).IsEquivalentTo(["pt", "added"]);
    }

    [Test]
    public async Task Refreshing_the_pivot_cache_keeps_the_cells()
    {
        using var source = Source();

        using var saved = source.LoadAndSave(wb => wb.PivotCaches.Single().Refresh());

        await Assert.That(saved.HasKeptCells(Sheet1)).IsTrue();
        await AssertValues(saved);
        await Assert.That(PivotTableNames(saved)).IsEquivalentTo(["pt"]);
    }

    [Test]
    public async Task The_kept_cells_are_those_a_rewrite_writes()
    {
        using var source = Source();

        using var kept = source.LoadAndSave(wb => wb.Worksheet("Data").PivotTable("pt").ColumnLabels.Add("Region"));
        using var rewritten = source.LoadAndSave(wb => wb.Worksheet("Data").PivotTable("pt").ColumnLabels.Add("Region"),
            new SaveOptions { RewriteUnchangedSheets = true });

        await Assert.That(rewritten.HasKeptCells(Sheet1)).IsFalse();
        await Assert.That(CellValues(kept)).IsEquivalentTo(CellValues(rewritten));
    }

    /// <summary>
    /// Sheet Data holds Name, Amount and Region in A1:C3, with 1.5 in B2, and the pivot table
    /// <c>pt</c> over them at E1. Its part is marked by <see cref="KeptSheet.Mark"/>.
    /// </summary>
    private static MemoryStream Source()
    {
        var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Name";
            ws.Cell("B1").Value = "Amount";
            ws.Cell("C1").Value = "Region";
            ws.Cell("A2").Value = "a";
            ws.Cell("B2").Value = 1.5;
            ws.Cell("C2").Value = "North";
            ws.Cell("A3").Value = "b";
            ws.Cell("B3").Value = 2;
            ws.Cell("C3").Value = "South";

            var pt = ws.PivotTables.Add("pt", ws.Cell("E1"), ws.Range("A1:C3"));
            pt.RowLabels.Add("Name");
            pt.Values.Add("Amount");

            wb.AddWorksheet("Other").Cell("A1").Value = "Other";
            wb.SaveAs(stream);
        }

        return stream.Mark(Sheet1);
    }

    private static async Task AssertValues(MemoryStream package)
    {
        package.Position = 0;
        using var wb = new XLWorkbook(package);
        var ws = wb.Worksheet("Data");

        await Assert.That(ws.Cell("A1").GetText()).IsEqualTo("Name");
        await Assert.That(ws.Cell("B2").GetDouble()).IsEqualTo(1.5);
        await Assert.That(ws.Cell("C3").GetText()).IsEqualTo("South");
    }

    private static string[] PivotTableNames(MemoryStream package)
    {
        package.Position = 0;
        using var wb = new XLWorkbook(package);
        return wb.Worksheet("Data").PivotTables.Select(pt => pt.Name).ToArray();
    }

    private static string[] CellValues(MemoryStream package)
    {
        package.Position = 0;
        using var wb = new XLWorkbook(package);
        return wb.Worksheet("Data").CellsUsed(XLCellsUsedOptions.All)
            .Select(c => $"{c.Address}={c.Value}|{c.Style.NumberFormat.Format}")
            .ToArray();
    }
}
