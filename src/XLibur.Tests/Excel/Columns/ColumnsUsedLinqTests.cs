using System.Linq;
using XLibur.Excel;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;

namespace XLibur.Tests.Excel.Columns;
/// <summary>
/// Guards that standard LINQ operators (<c>Skip</c>, <c>Where</c>) over the enumerable returned by
/// <c>IXLWorksheet.ColumnsUsed()</c> / <c>IXLWorksheet.RowsUsed()</c> behave as expected:
/// a filtered-out column/row must not be visited and therefore must not be adjusted. See
/// ClosedXML/ClosedXML#2867, which reports the opposite behavior on the parent library.
/// </summary>
public class ColumnsUsedLinqTests
{
    [Test]
    public async Task ColumnsUsed_EnumeratesInAscendingColumnOrder()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("C1").Value = "c";
        ws.Cell("A1").Value = "a";
        ws.Cell("B1").Value = "b";

        var order = ws.ColumnsUsed().Select(c => c.ColumnNumber()).ToList();

        await Assert.That(order).IsEquivalentTo(new[] { 1, 2, 3 }, CollectionOrdering.Matching);
    }

    [Test]
    public async Task ColumnsUsed_Skip_DoesNotAdjustSkippedColumn()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "Cheesecake sablfdsa daskjfhdsakjdsa and what more thing we have...!";
        ws.Cell("A2").Value = "Medovik";
        ws.Cell("B1").Value = 14;
        ws.Cell("B2").Value = 6;

        // Fix the first column width; skipping it must leave this untouched.
        ws.Column(1).Width = 20;
        var defaultBWidth = ws.Column(2).Width;

        var visited = ws.ColumnsUsed().Skip(1).Select(c => c.ColumnNumber()).ToList();
        foreach (var column in ws.ColumnsUsed().Skip(1))
            column.AdjustToContents();

        await Assert.That(visited).IsEquivalentTo(new[] { 2 }, CollectionOrdering.Matching).Because("Only the second column should be visited.");
        await Assert.That(ws.Column(1).Width).IsEqualTo(20).Within(1e-9).Because("Skipped first column must keep its width.");
        await Assert.That(ws.Column(2).Width).IsNotEqualTo(defaultBWidth).Because("Second column should have been adjusted.");
    }

    [Test]
    public async Task ColumnsUsed_Where_DoesNotAdjustFilteredColumn()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "Long long long long long long text";
        ws.Cell("B1").Value = "short";
        ws.Cell("C1").Value = "short";
        ws.Column(1).Width = 20;

        var visited = ws.ColumnsUsed()
            .Where(c => c.ColumnNumber() != 1)
            .Select(c => (Number: c.ColumnNumber(), WidthBefore: c.Width))
            .ToList();
        foreach (var column in ws.ColumnsUsed().Where(c => c.ColumnNumber() != 1))
            column.AdjustToContents();

        await Assert.That(visited.Select(v => v.Number)).IsEquivalentTo(new[] { 2, 3 }, CollectionOrdering.Matching);
        foreach (var v in visited)
            await Assert.That(ws.Column(v.Number).Width).IsNotEqualTo(v.WidthBefore).Because($"Visited column {v.Number} should have been adjusted.");
        await Assert.That(ws.Column(1).Width).IsEqualTo(20).Within(1e-9).Because("Filtered-out first column must keep its width.");
    }

    [Test]
    public async Task RowsUsed_Skip_DoesNotAdjustSkippedRow()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "row1";
        ws.Cell("A2").Value = "row2";
        ws.Cell("A3").Value = "row3";

        // Give every used row the same non-default height. Skipping row 1 must leave it untouched,
        // while the visited rows must be adjusted back to fit their content.
        ws.Row(1).Height = 40;
        ws.Row(2).Height = 40;
        ws.Row(3).Height = 40;

        var visited = ws.RowsUsed()
            .Skip(1)
            .Select(r => (Number: r.RowNumber(), HeightBefore: r.Height))
            .ToList();
        foreach (var row in ws.RowsUsed().Skip(1))
            row.AdjustToContents();

        await Assert.That(visited.Select(v => v.Number)).IsEquivalentTo(new[] { 2, 3 }, CollectionOrdering.Matching).Because("Only the trailing rows should be visited.");
        foreach (var v in visited)
            await Assert.That(ws.Row(v.Number).Height).IsNotEqualTo(v.HeightBefore).Because($"Visited row {v.Number} should have been adjusted.");
        await Assert.That(ws.Row(1).Height).IsEqualTo(40).Within(1e-9).Because("Skipped first row must keep its height.");
    }
}
