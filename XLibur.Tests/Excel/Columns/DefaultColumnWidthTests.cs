using System;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Graphics;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Columns;

/// <summary>
/// The default column width of a workbook created in memory follows Excel's rule, the same one the
/// load uses: 8 digits of the first font plus padding, rounded up to a multiple of 8 px. The digit
/// width depends on the DPI, so the width does too (#798).
/// </summary>
public class DefaultColumnWidthTests
{
    [Test]
    // Calibri 11 at 120 DPI has a 9 px digit: 8 digits and 7 px padding are 79 px, rounded up to
    // 80 px, which is 8.11 digits. Excel 16.0 (build 20430) at 120 DPI, over COM, opened the file
    // XLibur saves for a new workbook (sheetFormatPr states no width), drew column B 80 px (48 pt)
    // wide and reported StandardWidth 8.11.
    [Arguments(120, 8.11)]
    // At 96 DPI the digit is 7 px: 61 px rounded up to 64 px is 8.43 digits, Excel's well-known
    // default. This machine runs at 120 DPI only, so this case pins the rule, not a measurement.
    [Arguments(96, 8.43)]
    public async Task A_new_workbook_has_the_default_width_of_its_DPI(int dpi, double expected)
    {
        using var wb = new XLWorkbook(new LoadOptions { Dpi = new Point(dpi, dpi) });
        var ws = wb.AddWorksheet();

        await Assert.That(wb.ColumnWidth).IsEqualTo(expected).Within(0.005);
        await Assert.That(ws.ColumnWidth).IsEqualTo(expected).Within(0.005);
        await Assert.That(ws.Column(2).Width).IsEqualTo(expected).Within(0.005);
    }

    [Test]
    public async Task Creating_a_workbook_and_a_sheet_measures_no_font()
    {
        using var probe = new XLWorkbook();
        var engine = new CountingFontEngine(probe.FontEngine);
        using var wb = new XLWorkbook(new LoadOptions { FontEngine = engine });
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;

        await Assert.That(engine.MaxDigitWidthCalls).IsEqualTo(0);

        // The width is measured once, when it is first asked for.
        _ = ws.ColumnWidth;
        _ = wb.ColumnWidth;
        _ = wb.AddWorksheet().Column(3).Width;
        await Assert.That(engine.MaxDigitWidthCalls).IsEqualTo(1);
    }

    [Test]
    public async Task A_width_set_on_the_workbook_wins_over_the_default()
    {
        using var wb = new XLWorkbook(AtExcelDpi);
        var before = wb.AddWorksheet("Before");
        wb.ColumnWidth = 20;
        var after = wb.AddWorksheet("After");

        await Assert.That(wb.ColumnWidth).IsEqualTo(20);
        await Assert.That(after.ColumnWidth).IsEqualTo(20);
        await Assert.That(after.Column(2).Width).IsEqualTo(20);
        // A sheet keeps the width it was created with.
        await Assert.That(before.ColumnWidth).IsEqualTo(8.11).Within(0.005);
    }

    [Test]
    public async Task A_width_set_on_the_workbook_is_saved_on_the_sheets_created_after_it()
    {
        // Before #798 a new sheet took the workbook's width but did not count it as changed, so the
        // save left defaultColWidth out and a reload, or Excel, read the default width instead.
        // At 120 DPI the padding is 7 px of a 9 px digit: the 20 is saved as 20.777778, which Excel
        // 16.0 at 120 DPI shows as 20 (187 px). Before #803 the save added the 96 DPI padding,
        // 0.710625, and Excel showed 19.89 (defaultColWidth="20.710625", 186 px).
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(AtExcelDpi))
        {
            wb.AddWorksheet("Before");
            wb.ColumnWidth = 20;
            wb.AddWorksheet("After");
            wb.SaveAs(saved);
        }

        await Assert.That(saved.ReadPart("xl/worksheets/sheet2.xml")).Contains("defaultColWidth=\"20.777778\"");

        using var reloaded = new XLWorkbook(saved, AtExcelDpi);
        await Assert.That(reloaded.Worksheet("After").ColumnWidth).IsEqualTo(20).Within(0.005);
        await Assert.That(reloaded.Worksheet("Before").ColumnWidth).IsEqualTo(8.11).Within(0.005);
    }

    [Test]
    public async Task A_new_workbook_saves_no_default_width()
    {
        // With no defaultColWidth and no baseColWidth, Excel works the width out from the first
        // font at the DPI it runs at, so the file is right at any DPI.
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(AtExcelDpi))
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = 1;
            _ = ws.Column(2).Width;
            wb.SaveAs(saved);
        }

        await Assert.That(saved.Sheet1Xml()).Contains("<x:sheetFormatPr defaultRowHeight=\"15\" />");
    }

    private static readonly LoadOptions AtExcelDpi = new() { Dpi = new Point(120, 120) };

    /// <summary>
    /// The test font engine, counting how often the maximum digit width is measured.
    /// </summary>
    private sealed class CountingFontEngine(IXLFontEngine inner) : IXLFontEngine
    {
        public int MaxDigitWidthCalls { get; private set; }

        public double GetTextHeight(IXLFontBase font, double dpiY) => inner.GetTextHeight(font, dpiY);

        public double GetTextWidth(string text, IXLFontBase font, double dpiX) => inner.GetTextWidth(text, font, dpiX);

        public double GetMaxDigitWidth(IXLFontBase font, double dpiX)
        {
            MaxDigitWidthCalls++;
            return inner.GetMaxDigitWidth(font, dpiX);
        }

        public double GetDescent(IXLFontBase font, double dpiY) => inner.GetDescent(font, dpiY);

        public GlyphBox GetGlyphBox(ReadOnlySpan<int> graphemeCluster, IXLFontBase font, Dpi dpi) =>
            inner.GetGlyphBox(graphemeCluster, font, dpi);
    }
}
