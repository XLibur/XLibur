using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Ranges;

public class RangeColumnCopyToTests
{
    [Test]
    public async Task CopyTo_Cell_ReturnsColumnAtTargetCell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "First";
        ws.Cell("A3").Value = "Third";

        var result = ws.Range("A1:A3").Column(1).CopyTo(ws.Cell("E5"));

        await Assert.That(result.RangeAddress.ToString()).IsEqualTo("E5:E7");
        await Assert.That(result.Cell(1).Value).IsEqualTo((XLCellValue)"First");
        await Assert.That(result.Cell(3).Value).IsEqualTo((XLCellValue)"Third");
    }

    [Test]
    [Arguments("E5:E7")]
    [Arguments("E5")]
    [Arguments("E5:F9")]
    public async Task CopyTo_RangeBase_ReturnsColumnAtTargetFirstCell(string targetAddress)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "First";
        ws.Cell("A3").Value = "Third";

        var result = ws.Range("A1:A3").Column(1).CopyTo(ws.Range(targetAddress));

        await Assert.That(result.RangeAddress.ToString()).IsEqualTo("E5:E7");
        await Assert.That(result.Cell(1).Value).IsEqualTo((XLCellValue)"First");
        await Assert.That(result.Cell(3).Value).IsEqualTo((XLCellValue)"Third");
    }
}
