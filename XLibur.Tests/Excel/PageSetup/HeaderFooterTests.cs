using System.IO;
using System.Linq;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.PageSetup;

public class HeaderFooterTests
{
    [Test]
    public async Task CanChangeWorksheetHeader()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.PageSetup.Header.Center.AddText("Initial page header", XLHFOccurrence.EvenPages);

        var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        wb = new XLWorkbook(ms);
        ws = wb.Worksheets.First();

        ws.PageSetup.Header.Center.Clear();
        ws.PageSetup.Header.Center.AddText("Changed header", XLHFOccurrence.EvenPages);

        wb.SaveAs(ms, true);

        wb = new XLWorkbook(ms);
        ws = wb.Worksheets.First();

        var newHeader = ws.PageSetup.Header.Center.GetText(XLHFOccurrence.EvenPages);
        await Assert.That(newHeader).IsEqualTo("Changed header");
    }

    [Test]
    [Arguments("")]
    [Arguments("&L&C&\"Arial\"&9 19-10-2017 \n&9&\"Arial\" &P    &N &R")] // https://github.com/XLibur/XLibur/issues/563
    public async Task CanSetHeaderFooter(string s)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var header = (XLHeaderFooter)ws.PageSetup.Header;
        header.SetInnerText(XLHFOccurrence.AllPages, s);

        // Verify header was changed (or remained empty for empty input)
        await Assert.That(header.Changed).IsEqualTo(!string.IsNullOrEmpty(s));
    }

    [Test]
    public async Task SaveDoesNotCrash_WhenSimpleTextAddedToHeaderCenter()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.PageSetup.Header.Center.AddText("Simple Text", XLHFOccurrence.AllPages);

        using var ms = new MemoryStream();
        await Assert.That(() => wb.SaveAs(ms)).ThrowsNothing();
    }

    [Test]
    public async Task SaveDoesNotCrash_WhenTextAddedToAllHeaderPositions()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.PageSetup.Header.Left.AddText("Left Header Text", XLHFOccurrence.AllPages);
        ws.PageSetup.Header.Center.AddText("Center Header Text", XLHFOccurrence.AllPages);
        ws.PageSetup.Header.Right.AddText("Right Header Text", XLHFOccurrence.AllPages);

        using var ms = new MemoryStream();
        await Assert.That(() => wb.SaveAs(ms)).ThrowsNothing();
    }

    [Test]
    public async Task SaveDoesNotCrash_WhenHeaderContainsFormattedText()
    {
        // Reproduces the crash from issue: formatted text with font names and sizes
        // in headers causes ArgumentOutOfRangeException due to 255-char limit.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.PageSetup.Header.Left.AddText("Monday\n1 January to 7 January 2024", XLHFOccurrence.AllPages)
            .SetFontName("Arial").SetBold().SetFontSize(10);
        ws.PageSetup.Header.Center.AddText("OFFICIAL\n1 January to 7 January 2024", XLHFOccurrence.AllPages)
            .SetFontName("Calibri").SetFontSize(10);
        ws.PageSetup.Header.Center.AddText("Route 1 - Main Timetable", XLHFOccurrence.AllPages)
            .SetFontName("Arial").SetBold().SetFontSize(12);
        ws.PageSetup.Header.Right.AddText("Route Description", XLHFOccurrence.AllPages)
            .SetFontName("Arial").SetBold().SetFontSize(12);

        using var ms = new MemoryStream();
        await Assert.That(() => wb.SaveAs(ms)).ThrowsNothing();
    }

    [Test]
    public async Task SaveDoesNotCrash_WhenFormattedHeaderExceeds255Chars()
    {
        // The 255-char limit in GetText() is too restrictive. Format codes like
        // &"FontName,Bold"&12 count toward the limit, making it easy to exceed
        // with realistic header content across left/center/right positions.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.PageSetup.Header.Left.AddText("Monday Timetable\n1 January to 7 January 2024", XLHFOccurrence.AllPages)
            .SetFontName("Arial").SetBold().SetFontSize(10);
        ws.PageSetup.Header.Center.AddText("OFFICIAL DOCUMENT HEADER\n1 January to 7 January 2024", XLHFOccurrence.AllPages)
            .SetFontName("Calibri").SetFontSize(10);
        ws.PageSetup.Header.Center.AddText("Route 100 - Main Regional Timetable Schedule Overview", XLHFOccurrence.AllPages)
            .SetFontName("Arial").SetBold().SetFontSize(12);
        ws.PageSetup.Header.Right.AddText("North West Regional Route Full Description", XLHFOccurrence.AllPages)
            .SetFontName("Arial").SetBold().SetFontSize(12);

        using var ms = new MemoryStream();
        await Assert.That(() => wb.SaveAs(ms)).ThrowsNothing();
    }
}
