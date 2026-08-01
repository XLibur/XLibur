
using XLibur.Excel;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Worksheets;

public class XLSheetViewTests
{
    [Test]
    public async Task CopyWorksheetSheetViews()
    {
        using var wb1 = new XLWorkbook();
        using var wb2 = new XLWorkbook();

        var ws1 = wb1.AddWorksheet("WS1");
        ws1.SheetView.TopLeftCellAddress = ws1.Cell("AZ2000").Address;

        var ws2 = ws1.CopyTo(wb2, "WS2");

        await Assert.That(ws2.SheetView.Worksheet).IsEqualTo(ws2);
        await Assert.That(ws2.SheetView.TopLeftCellAddress.ToString()).IsEqualTo("AZ2000");
    }

    [Test]
    public async Task InvalidTopLeftCell()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet();
        var ws2 = wb.AddWorksheet();

        await Assert.That(() => ws1.SheetView.TopLeftCellAddress = ws2.Cell("A1").Address).Throws<ArgumentException>();
    }

    [Test]
    public async Task SheetViews()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            ws.SheetView.TopLeftCellAddress = ws.Cell("AZ2000").Address;
            wb.SaveAs(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();
            await Assert.That(ws.SheetView.TopLeftCellAddress.ToString()).IsEqualTo("AZ2000");

            ws.SheetView.TopLeftCellAddress = ws.Cell("AZ2000")
                .CellBelow()
                .CellRight()
                .Address;

            wb.Save();
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();
            await Assert.That(ws.SheetView.TopLeftCellAddress.ToString()).IsEqualTo("BA2001");
        }
    }
}
