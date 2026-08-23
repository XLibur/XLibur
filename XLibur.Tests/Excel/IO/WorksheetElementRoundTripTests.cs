using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// One workbook carrying every worksheet element the loader dispatches on, round-tripped and read
/// back. Spec 24 moves that dispatch from XLWorkbook_Load into WorksheetElementReader; this test is
/// what proves no element was dropped on the way.
/// </summary>
public class WorksheetElementRoundTripTests
{
    private static MemoryStream BuildWorkbookWithEveryElement()
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");

            // SheetFormatProperties
            ws.RowHeight = 22;
            ws.ColumnWidth = 14;

            // Columns
            ws.Column(2).Width = 33;

            // SheetData + MergeCells
            ws.Cell("A1").Value = "Header";
            ws.Cell("B1").Value = 42;
            ws.Range("D1:E1").Merge();

            // SheetViews
            ws.SheetView.FreezeRows(1);

            // AutoFilter
            ws.Range("A1:B1").SetAutoFilter();

            // SheetProtection
            ws.Protect("pw");

            // DataValidations
            ws.Range("A5:A6").CreateDataValidation().WholeNumber.Between(1, 10);

            // Hyperlinks
            ws.Cell("A8").SetValue("link").SetHyperlink(
                new XLHyperlink("https://example.invalid/"));

            // ConditionalFormatting
            ws.Range("B5:B6").AddConditionalFormat().WhenGreaterThan(5).Fill
                .SetBackgroundColor(XLColor.Red);

            // PrintOptions, PageMargins, PageSetup, HeaderFooter, breaks
            ws.PageSetup.CenterHorizontally = true;
            ws.PageSetup.Margins.Top = 1.25;
            ws.PageSetup.PaperSize = XLPaperSize.A4Paper;
            ws.PageSetup.Header.Left.AddText("hdr");
            ws.PageSetup.AddHorizontalPageBreak(3);
            ws.PageSetup.AddVerticalPageBreak(3);

            // SheetProperties -> tab colour
            ws.TabColor = XLColor.Blue;

            wb.SaveAs(ms);
        }

        ms.Position = 0;
        return ms;
    }

    [Test]
    public async Task Every_worksheet_element_survives_a_round_trip()
    {
        using var ms = BuildWorkbookWithEveryElement();
        using var wb = new XLWorkbook(ms);
        var ws = wb.Worksheet("Sheet1");

        await Assert.That(ws.RowHeight).IsEqualTo(22d);                       // SheetFormatProperties
        await Assert.That(ws.Column(2).Width).IsEqualTo(33d).Within(0.01);    // Columns
        await Assert.That(ws.Cell("A1").GetString()).IsEqualTo("Header");     // SheetData
        await Assert.That(ws.MergedRanges.Count).IsEqualTo(1);                // MergeCells
        await Assert.That(ws.SheetView.SplitRow).IsEqualTo(1);                // SheetViews
        await Assert.That(ws.AutoFilter.IsEnabled).IsTrue();                  // AutoFilter
        await Assert.That(ws.Protection.IsProtected).IsTrue();                // SheetProtection
        await Assert.That(ws.DataValidations.Count()).IsEqualTo(1);           // DataValidations
        await Assert.That(ws.Cell("A8").HasHyperlink).IsTrue();               // Hyperlinks
        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(1);        // ConditionalFormatting
        await Assert.That(ws.PageSetup.CenterHorizontally).IsTrue();          // PrintOptions
        await Assert.That(ws.PageSetup.Margins.Top).IsEqualTo(1.25).Within(0.01); // PageMargins
        await Assert.That(ws.PageSetup.PaperSize).IsEqualTo(XLPaperSize.A4Paper); // PageSetup
        // HeaderFooter. AddText(text) fans "AllPages" out into the three concrete occurrences, so
        // AllPages is never a key GetText can read back — ask for one of the concrete ones.
        await Assert.That(ws.PageSetup.Header.Left.GetText(XLHFOccurrence.OddPages))
            .IsEqualTo("hdr");
        await Assert.That(ws.PageSetup.RowBreaks.Count).IsEqualTo(1);         // RowBreaks
        await Assert.That(ws.PageSetup.ColumnBreaks.Count).IsEqualTo(1);      // ColumnBreaks
        await Assert.That(ws.TabColor).IsEqualTo(XLColor.Blue);               // SheetProperties
    }
}
