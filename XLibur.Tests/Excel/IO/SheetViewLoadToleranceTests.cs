using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// Every other attribute <c>LoadSheetViewProperties</c> reads degrades gracefully when the file says
/// something the reader does not understand — a boolean either parses or the attribute is treated as
/// absent, and the sheet loads. <c>sheetView/@view</c> is the one attribute in that method whose value
/// is an enumeration, and it must not be the one that turns an odd file into an unopenable one.
/// </summary>
public class SheetViewLoadToleranceTests
{
    [Test]
    public async Task A_view_value_the_reader_does_not_know_loads_as_Normal()
    {
        var package = SaveWorksheet().RewriteSheet1(xml =>
            Regex.Replace(xml, "\\bview=\"[^\"]*\"", "view=\"someFutureView\""));

        // Guard the fixture: the strip has to have found something to replace.
        await Assert.That(package.Sheet1Xml()).Contains("view=\"someFutureView\"");

        package.Position = 0;
        using var wb = new XLWorkbook(package);
        var ws = wb.Worksheets.First();

        await Assert.That(ws.SheetView.View).IsEqualTo(XLSheetViewOptions.Normal)
            .Because("an unrecognised view falls back to Normal rather than failing the load");
    }

    private static MemoryStream SaveWorksheet()
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.SheetView.SetView(XLSheetViewOptions.PageLayout);
            wb.SaveAs(ms);
        }

        return ms;
    }
}
