using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.Rows;

namespace XLibur.Tests.Excel.Ranges;

/// <summary>
/// Pins the string forms that the row and column list parsers accept: "A:C" and "1-3" pairs,
/// single entries, and comma lists. They all go through <c>XLHelper.SplitRangePair</c>, and the
/// <c>Range(string)</c> of a row or column goes through <c>XLRangeBase.RangeFromLineAddress</c>
/// (#621 finding 13).
/// </summary>
public class LineAddressParsingTests
{
    private static string Address(IXLRangeBase range)
        => $"{range.RangeAddress.FirstAddress}:{range.RangeAddress.LastAddress}";

    [Test]
    [Arguments("2:4", "B3:D3")]
    [Arguments("2-4", "B3:D3")]
    [Arguments("2", "B3:B3")]
    [Arguments("B3:D3", "B3:D3")]
    [Arguments("C3", "C3:C3")]
    public async Task WorksheetRow_Range_CountsNumbersAlongTheRow(string address, string expected)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        await Assert.That(Address(((XLRow)ws.Row(3)).Range(address))).IsEqualTo(expected);
    }

    [Test]
    [Arguments("2:4", "C2:C4")]
    [Arguments("2-4", "C2:C4")]
    [Arguments("5", "C5:C5")]
    [Arguments("C2:C4", "C2:C4")]
    public async Task WorksheetColumn_Range_CountsNumbersAlongTheColumn(string address, string expected)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        await Assert.That(Address(((XLColumn)ws.Column(3)).Range(address))).IsEqualTo(expected);
    }

    [Test]
    [Arguments("1:2", "B3:C3")]
    [Arguments("2-3", "C3:D3")]
    [Arguments("3", "D3:D3")]
    public async Task RangeRow_Range_IsRelativeToTheRowStart(string address, string expected)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var row = (XLRangeRow)ws.Range("B2:E6").Row(2);

        await Assert.That(Address(row.Range(address))).IsEqualTo(expected);
    }

    [Test]
    [Arguments("1:2", "C2:C3")]
    [Arguments("2-3", "C3:C4")]
    [Arguments("4", "C5:C5")]
    public async Task RangeColumn_Range_IsRelativeToTheColumnStart(string address, string expected)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var column = (XLRangeColumn)ws.Range("B2:E6").Column(2);

        await Assert.That(Address(column.Range(address))).IsEqualTo(expected);
    }

    [Test]
    [Arguments("2-3, 5", new[] { 2, 3, 5 })]
    [Arguments("2:3", new[] { 2, 3 })]
    [Arguments("4", new[] { 4 })]
    public async Task Worksheet_Rows_ParsesPairsAndSingles(string rows, int[] expected)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var actual = ws.Rows(rows).Select(r => r.RowNumber()).ToArray();

        await Assert.That(actual).IsEquivalentTo(expected);
    }

    [Test]
    [Arguments("B:C, 5", new[] { 2, 3, 5 })]
    [Arguments("2-3", new[] { 2, 3 })]
    [Arguments("D", new[] { 4 })]
    public async Task Worksheet_Columns_ParsesLettersNumbersAndSingles(string columns, int[] expected)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var actual = ws.Columns(columns).Select(c => c.ColumnNumber()).ToArray();

        await Assert.That(actual).IsEquivalentTo(expected);
    }

    [Test]
    [Arguments("1-2", new[] { 2, 3 })]
    [Arguments("C:D", new[] { 4, 5 })]
    [Arguments("1, 4", new[] { 2, 5 })]
    public async Task Range_Columns_IsRelativeToTheRange(string columns, int[] expected)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var actual = ws.Range("B2:E6").Columns(columns)
            .Select(c => c.RangeAddress.FirstAddress.ColumnNumber).ToArray();

        await Assert.That(actual).IsEquivalentTo(expected);
    }

    [Test]
    [Arguments("2-3", new[] { 3, 4 })]
    [Arguments("1:1, 5", new[] { 2, 6 })]
    public async Task Range_Rows_IsRelativeToTheRange(string rows, int[] expected)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var actual = ws.Range("B2:E6").Rows(rows)
            .Select(r => r.RangeAddress.FirstAddress.RowNumber).ToArray();

        await Assert.That(actual).IsEquivalentTo(expected);
    }

    [Test]
    public async Task RangeRow_Rows_AndRangeColumn_Columns_TakeCellPairs()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range("B2:E6");

        // The pair halves are worksheet cell addresses here, not offsets into the line.
        var rows = range.Row(2).Rows("B3:C3, D3-E3").Select(Address).ToArray();
        var columns = range.Column(2).Columns("C2:C3, C4-C5").Select(Address).ToArray();

        await Assert.That(rows).IsEquivalentTo(new[] { "B3:C3", "D3:E3" });
        await Assert.That(columns).IsEquivalentTo(new[] { "C2:C3", "C4:C5" });
    }

    [Test]
    public async Task TableRange_Rows_ParsesPairsAndSingles()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "H";
        for (var r = 2; r <= 6; r++)
            ws.Cell(r, 1).Value = r;
        var table = ws.Range("A1:A6").CreateTable();

        var actual = table.DataRange!.Rows("1-2, 4")
            .Select(r => r.RangeAddress.FirstAddress.RowNumber).ToArray();

        await Assert.That(actual).IsEquivalentTo(new[] { 2, 3, 5 });
    }
}
