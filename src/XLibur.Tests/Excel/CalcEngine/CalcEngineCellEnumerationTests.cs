using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;

public class CalcEngineCellEnumerationTests
{
    [Test]
    public async Task CanEnumerateCellsOverEmptySheet()
    {
        using var wb = new XLWorkbook();
        var sheet1 = wb.AddWorksheet("Sheet1");
        var sheet2 = wb.AddWorksheet("Sheet2");

        var cell = sheet1.FirstCell();
        cell.FormulaA1 = "=SUMIFS(Sheet2!B:B, Sheet2!C:C, 1)";

        await Assert.That(cell.Value).IsEqualTo(0);
    }
}
