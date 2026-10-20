using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Columns;

/// <summary>
/// A <c>&lt;col width&gt;</c> is the width a column shows in Excel plus a padding, both in digits of
/// the workbook's first font. The padding is <c>2 * ceil(mdw / 4) + 1</c> px, so in digits it
/// depends on the digit width, and with it on the DPI (#803).
/// </summary>
/// <remarks>
/// Measured in Excel 16.0 (build 20430) over COM at 120 DPI, with Calibri 11 as the Normal font
/// (a 9 px digit, a 7 px padding). Setting <c>ColumnWidth</c> and saving wrote:
/// <code>
/// set 20   -> 187 px, width="20.77734375"   (20 + 7/9 = 20.777..., truncated to 1/256)
/// set 10   -> 97 px,  width="10.77734375"
/// set 1    -> 16 px,  width="1.77734375"
/// set 0.5  -> 8 px,   width="0.88671875"    (0.5 * 16 px / 9; below 1 digit the padding scales)
/// set 0.3  -> 5 px,   width="0.5546875"
/// </code>
/// Excel truncates the width to 1/256 and draws whole pixels; XLibur keeps the width it was given.
/// And the other way, file widths Excel opened at 120 DPI:
/// <code>
/// 20.77734375, 20.777778          -> 20    (187 px)
/// 20.7109375, 20.714286, 20.710625 -> 19.89 (186 px)
/// 20.5 -> 19.67 (184 px)   20.9 -> 20.11 (188 px)   0.88671875 -> 0.5 (8 px)   1.2 -> 0.69 (11 px)
/// </code>
/// So the load takes off the same padding the save adds, worked out from the same digit width. It
/// does not snap to whole pixels as Excel draws them, so 20.5 loads as 19.72, not 19.67, and a width
/// XLibur saved loads back as it was set. Before #803 the save added a fixed 0.710625, the padding
/// of a 7 px digit (96 DPI), and Excel at 120 DPI showed a width of 20 saved that way as 19.89.
/// This machine runs at 120 DPI only; the 96 DPI cases are pinned by the rule, not by Excel.
/// </remarks>
public class ColumnWidthPaddingTests
{
    [Test]
    [Arguments(120, 20, "20.777778")]
    [Arguments(96, 20, "20.714286")]
    public async Task A_column_width_is_saved_with_the_padding_of_the_workbooks_digit(int dpi, double width, string expected)
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(new LoadOptions { Dpi = new Point(dpi, dpi) }))
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = 1;
            ws.Column(2).Width = width;
            wb.SaveAs(saved);
        }

        await Assert.That(saved.Sheet1Xml()).Contains($"<x:col min=\"2\" max=\"2\" width=\"{expected}\" customWidth=\"1\" />");
    }

    [Test]
    // Below one digit Excel scales the padding with the width: 0.5 is 8 px of 16, width 8/9.
    [Arguments(120, 0.5, "0.888889")]
    [Arguments(120, 0.3, "0.533333")]
    public async Task A_width_below_one_digit_is_saved_with_a_scaled_padding(int dpi, double width, string expected)
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(new LoadOptions { Dpi = new Point(dpi, dpi) }))
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = 1;
            ws.Column(2).Width = width;
            wb.SaveAs(saved);
        }

        await Assert.That(saved.Sheet1Xml()).Contains($"<x:col min=\"2\" max=\"2\" width=\"{expected}\" customWidth=\"1\" />");
    }

    [Test]
    // The widths Excel wrote at 120 DPI for the widths it showed (see the class remarks). Excel
    // truncates to 1/256, so each loads within 1/256 of a digit of what Excel showed. Before #803
    // the load took off 0.710625 and 20.77734375 loaded as 20.07.
    [Arguments("20.77734375", 20)]
    [Arguments("10.77734375", 10)]
    [Arguments("1.77734375", 1)]
    [Arguments("0.88671875", 0.5)]
    public async Task A_width_Excel_wrote_loads_as_the_width_Excel_showed(string fileWidth, double shown)
    {
        using var file = SheetWithColumnB(fileWidth);
        using var wb = new XLWorkbook(file, AtExcelDpi);

        await Assert.That(wb.Worksheet(1).Column(2).Width).IsEqualTo(shown).Within(1 / 256d);
    }

    [Test]
    [Arguments(120)]
    [Arguments(96)]
    [Arguments(144)]
    public async Task A_width_saved_at_any_DPI_loads_back_as_it_was(int dpi)
    {
        var options = new LoadOptions { Dpi = new Point(dpi, dpi) };
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(options))
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = 1;
            ws.Column(2).Width = 20;
            ws.Column(3).Width = 0.5;
            ws.Column(4).Width = 13.57;
            wb.SaveAs(saved);
        }

        using var reloaded = new XLWorkbook(saved, options);
        var sheet = reloaded.Worksheet(1);
        await Assert.That(sheet.Column(2).Width).IsEqualTo(20).Within(1e-6);
        await Assert.That(sheet.Column(3).Width).IsEqualTo(0.5).Within(1e-6);
        await Assert.That(sheet.Column(4).Width).IsEqualTo(13.57).Within(1e-6);
    }

    [Test]
    // 8.0078125 loads as 7.293527, to the six decimals the load keeps, which saves as 8.007813: not
    // the width the file states at six decimals, but the same width once loaded.
    [Arguments(96, "8.0078125")]
    [Arguments(120, "8.0078125")]
    [Arguments(120, "20.77734375")]
    public async Task A_width_the_file_spelled_with_more_digits_stays_while_it_is_unchanged(int dpi, string fileWidth)
    {
        var options = new LoadOptions { Dpi = new Point(dpi, dpi) };
        using var file = SheetWithColumnB(fileWidth);
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(file, options))
        {
            wb.Worksheet(1).Cell("A1").Value = 2;
            wb.SaveAs(saved);
        }

        await Assert.That(saved.Sheet1Xml()).Contains($"<x:col min=\"2\" max=\"2\" width=\"{fileWidth}\" customWidth=\"1\" />");
    }

    private static readonly LoadOptions AtExcelDpi = new() { Dpi = new Point(120, 120) };

    /// <summary>A workbook XLibur saved, with column B's width rewritten to <paramref name="fileWidth"/>.</summary>
    private static MemoryStream SheetWithColumnB(string fileWidth)
    {
        var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = 1;
            ws.Column(2).Width = 30;
            wb.SaveAs(saved);
        }

        return saved.RewriteSheet1(xml =>
            Regex.Replace(xml,
                "<x:col min=\"2\" max=\"2\" width=\"[^\"]+\"", $"<x:col min=\"2\" max=\"2\" width=\"{fileWidth}\""));
    }
}
