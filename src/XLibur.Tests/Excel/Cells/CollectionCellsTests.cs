using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Cells;

/// <summary>
/// Pins <c>Cells()</c>/<c>CellsUsed()</c> on every row, column and range collection. They all
/// build their result through <c>XLCells.FromRanges</c> (#621), so they must agree on what a
/// used cell is: <c>CellsUsed()</c> counts contents only, and formatting counts only when asked.
/// </summary>
public class CollectionCellsTests
{
    private static XLWorkbook CreateWorkbook(out IXLWorksheet ws, out IXLTable table)
    {
        var wb = new XLWorkbook();
        ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "H1";
        ws.Cell("B1").Value = "H2";
        ws.Cell("C1").Value = "H3";
        table = ws.Range("A1:C4").CreateTable();

        // Data area A2:C4 holds a value, a format-only cell and another value.
        ws.Cell("A2").Value = "a";
        ws.Cell("B3").Style.Fill.BackgroundColor = XLColor.Yellow;
        ws.Cell("C4").Value = "c";
        return wb;
    }

    private static string[] Addresses(IXLCells cells) => cells.Select(c => c.Address.ToString()!).ToArray();

    private static async Task AssertDataAreaCells(IXLCells all, IXLCells used, IXLCells usedWithFormats)
    {
        await Assert.That(Addresses(all)).IsEquivalentTo(new[] { "A2", "B2", "C2", "A3", "B3", "C3", "A4", "B4", "C4" });
        await Assert.That(Addresses(used)).IsEquivalentTo(new[] { "A2", "C4" });
        await Assert.That(Addresses(usedWithFormats)).IsEquivalentTo(new[] { "A2", "B3", "C4" });
    }

    [Test]
    public async Task RangeRows_Cells()
    {
        using var wb = CreateWorkbook(out var ws, out _);
        var rows = ws.Range("A2:C4").Rows();

        await AssertDataAreaCells(rows.Cells(), rows.CellsUsed(), rows.CellsUsed(XLCellsUsedOptions.All));
    }

    [Test]
    public async Task RangeColumns_Cells()
    {
        using var wb = CreateWorkbook(out var ws, out _);
        var columns = ws.Range("A2:C4").Columns();

        await AssertDataAreaCells(columns.Cells(), columns.CellsUsed(), columns.CellsUsed(XLCellsUsedOptions.All));
    }

    [Test]
    public async Task Ranges_Cells_CoversEveryRangeOnce()
    {
        using var wb = CreateWorkbook(out var ws, out _);
        var ranges = ws.Ranges("A2:A4,B2:C4");

        await AssertDataAreaCells(ranges.Cells(), ranges.CellsUsed(), ranges.CellsUsed(XLCellsUsedOptions.All));
    }

    [Test]
    public async Task TableRows_Cells()
    {
        using var wb = CreateWorkbook(out _, out var table);
        var rows = table.DataRange!.Rows();

        await AssertDataAreaCells(rows.Cells(), rows.CellsUsed(), rows.CellsUsed(XLCellsUsedOptions.All));
    }

    [Test]
    public async Task Rows_Cells()
    {
        using var wb = CreateWorkbook(out var ws, out _);
        var rows = ws.Rows("2:4");

        // Whole rows are 3 x 16384 cells, so check the start of the enumeration only.
        await Assert.That(rows.Cells().Take(2).Select(c => c.Address.ToString()!).ToArray()).IsEquivalentTo(new[] { "A2", "B2" });
        await Assert.That(Addresses(rows.CellsUsed())).IsEquivalentTo(new[] { "A2", "C4" });
        await Assert.That(Addresses(rows.CellsUsed(XLCellsUsedOptions.All))).IsEquivalentTo(new[] { "A2", "B3", "C4" });
    }

    [Test]
    public async Task Columns_Cells()
    {
        using var wb = CreateWorkbook(out var ws, out _);
        var columns = ws.Columns("A:C");

        // Whole columns are too big to enumerate; take the first cell only. It is lazy.
        await Assert.That(columns.Cells().First().Address.ToString()).IsEqualTo("A1");
        await Assert.That(Addresses(columns.CellsUsed())).IsEquivalentTo(new[] { "A1", "B1", "C1", "A2", "C4" });
        await Assert.That(Addresses(columns.CellsUsed(includeFormats: true))).IsEquivalentTo(new[] { "A1", "B1", "C1", "A2", "B3", "C4" });
        await Assert.That(Addresses(columns.CellsUsed(includeFormats: false))).IsEquivalentTo(new[] { "A1", "B1", "C1", "A2", "C4" });
    }
}
