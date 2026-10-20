using System;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Columns;

/// <summary>
/// Excel lays columns out in digits of the stylesheet's first font (<c>&lt;fonts&gt;</c>, index 0),
/// whatever font the Normal style or the default cell format has (#786). It even reports the first
/// font as the Normal style's font.
/// </summary>
/// <remarks>
/// <para>
/// The expected values were measured in Excel (Microsoft 365, 16.0 build 20430, over COM) from the
/// same fixtures, at 120 DPI: <c>StandardWidth</c> for the sheet's width, and
/// <c>Columns(1).AutoFit()</c> and <c>Rows(2).AutoFit()</c> for the autofit cases, after typing
/// the same text into the same cells. So these tests load at 120 DPI.
/// </para>
/// <para>
/// The fixtures use only Calibri, at 11 and 20 pt. Its digit width at 120 DPI is 9 px and 17 px both
/// in Excel and in the test font engine, and the engine finds a metric-compatible font (Carlito)
/// where Calibri is not installed. Arial 20 is 18 px in Excel but 19 px in the engine, so it would
/// not test the font choice alone. Excel drew the sheets as follows:
/// </para>
/// <list type="bullet">
/// <item><c>plain</c>: Calibri 11 only, as XLibur writes it. Columns are measured in 9 px digits.</item>
/// <item><c>firstFont</c>: the first font is Calibri 20, which nothing uses; the Normal style and the
/// cells use Calibri 11. Excel reports Calibri 20 as the Normal style's font, and columns are
/// measured in 17 px digits.</item>
/// <item><c>cellFont</c>: the first font is Calibri 11, the default cell format uses Calibri 20.
/// Columns are measured in 9 px digits.</item>
/// <item><c>normalStyleFont</c>: the first font is Calibri 11, the Normal style uses Calibri 20, and
/// the cells Calibri 11. Excel reports Calibri 11 as the Normal style's font, and columns are
/// measured in 9 px digits.</item>
/// </list>
/// </remarks>
public class FirstFontDigitWidthTests
{
    private static readonly LoadOptions AtExcelDpi = new() { Dpi = new Point(120, 120) };

    [Test]
    // defaultColWidth="12.78" is 217 px of 17 px digits: 12.12 characters.
    [Arguments("firstFont", "defaultColWidth", 12.12)]
    // baseColWidth="12" is 12 digits and 11 px padding, 215 px, rounded up to 216 px: 12.06.
    [Arguments("firstFont", "baseColWidth", 12.06)]
    // No width is baseColWidth 8: 147 px, rounded up to 152 px: 8.29.
    [Arguments("firstFont", "none", 8.29)]
    // In 9 px digits: 115 px, 120 px and 80 px.
    [Arguments("cellFont", "defaultColWidth", 12)]
    [Arguments("cellFont", "baseColWidth", 12.56)]
    [Arguments("cellFont", "none", 8.11)]
    [Arguments("normalStyleFont", "defaultColWidth", 12)]
    [Arguments("normalStyleFont", "baseColWidth", 12.56)]
    [Arguments("normalStyleFont", "none", 8.11)]
    public async Task A_sheet_default_width_is_in_digits_of_the_first_font(string fonts, string sheetFormat,
        double excelStandardWidth)
    {
        using var wb = new XLWorkbook(Fixture(fonts, sheetFormat), AtExcelDpi);
        var ws = wb.Worksheet(1);

        // Excel shows a width to two decimals.
        await Assert.That(ws.ColumnWidth).IsEqualTo(excelStandardWidth).Within(0.005);
        await Assert.That(ws.Column(2).Width).IsEqualTo(excelStandardWidth).Within(0.005);
    }

    [Test]
    public async Task A_loaded_default_width_is_written_back_as_the_file_stated_it()
    {
        // The save recognises the width the load converted, so it keeps the file's value.
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(Fixture("firstFont", "defaultColWidth"), AtExcelDpi))
        {
            wb.Worksheet(1).Cell("A1").Value = "edited";
            wb.SaveAs(saved);
        }

        await Assert.That(saved.Sheet1Xml()).Contains("defaultColWidth=\"12.78\"");
    }

    [Test]
    [Arguments("firstFont", "20")]
    [Arguments("cellFont", "11")]
    [Arguments("normalStyleFont", "11")]
    public async Task A_saved_stylesheet_keeps_its_first_font_first(string fonts, string firstFontSize)
    {
        // The columns of the saved file are measured in the same digits as the loaded file's, so a
        // reload reads the same width, and so does Excel.
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(Fixture(fonts, "none"), AtExcelDpi))
        {
            wb.Worksheet(1).Cell("A1").Value = "edited";
            wb.SaveAs(saved);
        }

        await Assert.That(FirstFont(saved)).Contains($"<x:sz val=\"{firstFontSize}\" />").And.Contains("<x:name val=\"Calibri\" />");

        using var reloaded = new XLWorkbook(saved, AtExcelDpi);
        await Assert.That(reloaded.Worksheet(1).ColumnWidth).IsEqualTo(fonts == "firstFont" ? 8.29 : 8.11).Within(0.005);
    }

    private static string FirstFont(Stream saved) =>
        Regex.Match(saved.ReadPart("xl/styles.xml"), "<x:font>.*?</x:font>", RegexOptions.None, RegexTimeout).Value;

    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(1);

    private const string AutofitText = "The quick brown fox jumps";

    private const string WrappedText = "aaaa bbbb cccc dddd eeee ffff gggg hhhh";

    [Test]
    // Calibri 11 text everywhere: Excel autofits to 22.22 characters (207 px).
    [Arguments("plain", 22.22)]
    // The same text with a Normal style of Calibri 20 autofits the same: the column is still in
    // Calibri 11 digits.
    [Arguments("normalStyleFont", 22.22)]
    // The same text with an unused first font of Calibri 20 is 204 px, which is 11.35 of its digits.
    [Arguments("firstFont", 11.35)]
    // Calibri 20 text, in Calibri 11 digits: 41.11 characters (377 px).
    [Arguments("cellFont", 41.11)]
    public async Task Column_autofit_measures_the_width_in_digits_of_the_first_font(string fonts,
        double excelAutofitWidth)
    {
        using var wb = new XLWorkbook(Fixture(fonts, "none"), AtExcelDpi);
        var ws = wb.Worksheet(1);
        ws.Cell("A1").Value = AutofitText;

        ws.Column(1).AdjustToContents();

        await AssertWidthNearExcel(ws.Column(1).Width, excelAutofitWidth);
    }

    [Test]
    // Calibri 11 text in a column of 8.11 Calibri 11 digits (80 px) wraps to 100.8 pt.
    [Arguments("plain", 100.8)]
    [Arguments("normalStyleFont", 100.8)]
    // A column of 8.29 Calibri 20 digits is 152 px, and the same text wraps to 43.2 pt.
    [Arguments("firstFont", 43.2)]
    // Calibri 20 text in a column of 8.11 Calibri 11 digits (80 px) wraps to 283.8 pt.
    [Arguments("cellFont", 283.8)]
    public async Task Row_autofit_wraps_at_the_column_width_in_digits_of_the_first_font(string fonts,
        double excelAutofitHeight)
    {
        using var wb = new XLWorkbook(Fixture(fonts, "none"), AtExcelDpi);
        var ws = wb.Worksheet(1);
        ws.Cell("B2").Value = WrappedText;
        ws.Cell("B2").Style.Alignment.WrapText = true;

        ws.Row(2).AdjustToContents();

        await AssertHeightNearExcel(ws.Row(2).Height, excelAutofitHeight);
    }

    [Test]
    public async Task A_workbook_created_in_memory_measures_columns_in_the_default_font()
    {
        // The save writes the default font, Calibri 11, first, whatever font the workbook's style
        // has. Excel opened this workbook's file, saved with a Calibri 20 style, and autofitted
        // column A to 41.11 characters and row 2 to 283.8 pt, the same as the cellFont fixture.
        using var wb = new XLWorkbook(AtExcelDpi);
        wb.Style.Font.FontSize = 20;
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = AutofitText;
        ws.Cell("B2").Value = WrappedText;
        ws.Cell("B2").Style.Alignment.WrapText = true;

        ws.Column(1).AdjustToContents();
        ws.Row(2).AdjustToContents();

        await AssertWidthNearExcel(ws.Column(1).Width, 41.11);
        await AssertHeightNearExcel(ws.Row(2).Height, 283.8);

        using var saved = new MemoryStream();
        wb.SaveAs(saved);
        await Assert.That(FirstFont(saved))
            .Contains("<x:sz val=\"11\" />").And.Contains("<x:name val=\"Calibri\" />");
    }

    /// <summary>
    /// Autofit pads the text as Excel does (#789, <see cref="AutofitPaddingTests"/>), but measures it
    /// with the font engine, which is wider than Excel's measurement. Excel draws the text with GDI at
    /// an em of whole pixels and with hinted advance widths: <see cref="AutofitText"/> is 195 px in
    /// Calibri 11 and 357 px in Calibri 20 there, and 201 px and 366 px in the engine, which scales
    /// the outlines' advance widths to the exact em. So column A is 213 px where Excel is 207 px,
    /// 210 px where it is 204 px, and 387 px where it is 377 px. In characters, which leave out the
    /// column's own padding, that is 22.89 for 22.22, 11.71 for 11.35 and 42.22 for 41.11: within
    /// 3.5 %. A digit width of the wrong font misses by about half or double.
    /// </summary>
    private static async Task AssertWidthNearExcel(double actual, double excel) =>
        await Assert.That(actual).IsEqualTo(excel).Within(excel * 0.035);

    /// <summary>
    /// Wrapped text breaks into as many lines as in Excel, but a line is lower: the engine makes a
    /// line the em plus the descent, 23 px in Calibri 11 and 41 px in Calibri 20 at 120 DPI, where
    /// Excel's lines are 24 px and 43 px. Row 2 is 7 lines of 23 px (96.6 pt) where Excel is 100.8 pt,
    /// and 11 lines of 41 px (270.6 pt) where Excel is 283.8 pt: within 5 %.
    /// </summary>
    private static async Task AssertHeightNearExcel(double actual, double excel) =>
        await Assert.That(actual).IsEqualTo(excel).Within(excel * 0.05);

    private const string Calibri11 =
        "<x:font><x:sz val=\"11\" /><x:color rgb=\"FF000000\" /><x:name val=\"Calibri\" /><x:family val=\"2\" /></x:font>";

    private const string Calibri20 =
        "<x:font><x:sz val=\"20\" /><x:color rgb=\"FF000000\" /><x:name val=\"Calibri\" /><x:family val=\"2\" /></x:font>";

    /// <summary>
    /// An empty sheet (a 1 in E5) whose <c>sheetFormatPr</c> states <paramref name="sheetFormat"/>, in a
    /// stylesheet laid out as <paramref name="fonts"/> names.
    /// </summary>
    private static MemoryStream Fixture(string fonts, string sheetFormat)
    {
        var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.Worksheets.Add("Sheet1").Cell("E5").Value = 1;
            wb.SaveAs(package);
        }

        var sheetFormatPr = sheetFormat switch
        {
            "defaultColWidth" => "<x:sheetFormatPr defaultColWidth=\"12.78\" defaultRowHeight=\"15\" />",
            "baseColWidth" => "<x:sheetFormatPr baseColWidth=\"12\" defaultRowHeight=\"15\" />",
            "none" => "<x:sheetFormatPr defaultRowHeight=\"15\" />",
            _ => throw new ArgumentOutOfRangeException(nameof(sheetFormat)),
        };
        package.RewriteSheet1(xml => xml.Replace("<x:sheetFormatPr defaultRowHeight=\"15\" />", sheetFormatPr));

        return fonts switch
        {
            "plain" => package,
            "firstFont" => package.RewritePart("xl/styles.xml", xml => WithStyles(xml, Calibri20 + Calibri11,
                "<x:cellStyleXfs count=\"1\"><x:xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" /></x:cellStyleXfs>",
                "<x:cellXfs count=\"1\"><x:xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" /></x:cellXfs>",
                "<x:cellStyles count=\"1\"><x:cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\" /></x:cellStyles>")),
            "cellFont" => package.RewritePart("xl/styles.xml", xml => WithStyles(xml, Calibri11 + Calibri20,
                "<x:cellStyleXfs count=\"1\"><x:xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" /></x:cellStyleXfs>",
                "<x:cellXfs count=\"1\"><x:xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" /></x:cellXfs>",
                "<x:cellStyles count=\"1\"><x:cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\" /></x:cellStyles>")),
            "normalStyleFont" => package.RewritePart("xl/styles.xml", xml => WithStyles(xml, Calibri11 + Calibri20,
                "<x:cellStyleXfs count=\"2\"><x:xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" />" +
                "<x:xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" /></x:cellStyleXfs>",
                "<x:cellXfs count=\"2\"><x:xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" />" +
                "<x:xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"1\" /></x:cellXfs>",
                "<x:cellStyles count=\"2\"><x:cellStyle name=\"Plain\" xfId=\"0\" />" +
                "<x:cellStyle name=\"Normal\" xfId=\"1\" builtinId=\"0\" /></x:cellStyles>")),
            _ => throw new ArgumentOutOfRangeException(nameof(fonts)),
        };
    }

    private static string WithStyles(string xml, string fonts, string cellStyleXfs, string cellXfs, string cellStyles)
    {
        xml = Regex.Replace(xml, "<x:fonts.*?</x:fonts>", $"<x:fonts count=\"2\">{fonts}</x:fonts>", RegexOptions.None, RegexTimeout);
        xml = Regex.Replace(xml, "<x:cellStyleXfs.*?</x:cellStyleXfs>", cellStyleXfs, RegexOptions.None, RegexTimeout);
        xml = Regex.Replace(xml, "<x:cellXfs.*?</x:cellXfs>", cellXfs, RegexOptions.None, RegexTimeout);
        return Regex.Replace(xml, "<x:cellStyles.*?</x:cellStyles>", cellStyles, RegexOptions.None, RegexTimeout);
    }
}
