using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.RichText;

/// <summary>
/// The font of a text's phonetic properties is saved in the styles part, and
/// <c>phoneticPr/@fontId</c> is its index there (#801). It comes after the default font, which a new
/// workbook keeps at index 0 (#802).
/// </summary>
public class PhoneticFontIdSaveTests
{
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private static readonly XLibur.Excel.LoadOptions AtExcelDpi = new() { Dpi = new Point(120, 120) };

    [Test]
    public async Task A_shared_text_names_its_phonetic_font_when_the_stylesheet_already_has_fonts()
    {
        // The loaded stylesheet keeps its fonts in their order, and a cell registers another font
        // before the phonetic one. The phonetic font is appended, so its index is not 0.
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(PlainWorkbook()))
        {
            var ws = wb.Worksheet(1);
            ws.Cell("B1").Value = "other";
            ws.Cell("B1").Style.Font.SetFontName("Times New Roman").Font.SetFontSize(14);
            AddPhonetics(ws.Cell("A1"), "Arial", 9);
            wb.SaveAs(saved);
        }

        var phoneticPr = XDocument.Parse(saved.ReadPart("xl/sharedStrings.xml"))
            .Descendants(Main + "phoneticPr").Single();

        await Assert.That(FontName(Fonts(saved)[(int)phoneticPr.Attribute("fontId")!])).IsEqualTo("Arial 9");
    }

    [Test]
    public async Task A_shared_text_in_a_new_workbook_names_its_phonetic_font()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell("B1").Value = "other";
            ws.Cell("B1").Style.Font.SetFontName("Times New Roman").Font.SetFontSize(14);
            AddPhonetics(ws.Cell("A1"), "Arial", 9);
            wb.SaveAs(saved);
        }

        var phoneticPr = XDocument.Parse(saved.ReadPart("xl/sharedStrings.xml"))
            .Descendants(Main + "phoneticPr").Single();

        await Assert.That(FontName(Fonts(saved)[(int)phoneticPr.Attribute("fontId")!])).IsEqualTo("Arial 9");
    }

    [Test]
    public async Task An_inline_text_names_its_phonetic_font()
    {
        // An inline text is written with the worksheet, after the styles part. Its phonetic font,
        // which no cell style uses, still has to be in the styles part.
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell("B1").Value = "other";
            ws.Cell("B1").Style.Font.SetFontName("Times New Roman").Font.SetFontSize(14);
            AddPhonetics(ws.Cell("A1"), "Arial", 9);
            ws.Cell("A1").ShareString = false;
            wb.SaveAs(saved);
        }

        var phoneticPr = XDocument.Parse(saved.Sheet1Xml()).Descendants(Main + "phoneticPr").Single();

        await Assert.That(FontName(Fonts(saved)[(int)phoneticPr.Attribute("fontId")!])).IsEqualTo("Arial 9");
    }

    [Test]
    public async Task A_new_workbook_with_phonetics_keeps_the_default_font_first()
    {
        // Excel measures the columns in digits of the first font (#786), so a phonetic font saved
        // first would change every column width (#802).
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            AddPhonetics(wb.Worksheets.Add("Sheet1").Cell("A1"), "Calibri", 8);
            wb.SaveAs(saved);
        }

        var fonts = Fonts(saved);
        await Assert.That(FontName(fonts[0])).IsEqualTo("Calibri 11");
        await Assert.That(fonts.Select(FontName)).Contains("Calibri 8");
    }

    [Test]
    public async Task A_new_workbook_with_phonetics_has_the_default_column_width_of_one_without()
    {
        // Excel (Microsoft 365, over COM) reported StandardWidth 8.11 for both files, and Calibri 11
        // as the Normal style's font. For the file saved before the fix, with Calibri 8 first, it
        // reported 8.43 and Calibri 8.
        using var withPhonetics = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            AddPhonetics(wb.Worksheets.Add("Sheet1").Cell("A1"), "Calibri", 8);
            wb.SaveAs(withPhonetics);
        }

        using var withoutPhonetics = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.Worksheets.Add("Sheet1").Cell("A1").Value = "みんなさんはお元気ですか。";
            wb.SaveAs(withoutPhonetics);
        }

        using var reloadedWith = new XLWorkbook(withPhonetics, AtExcelDpi);
        using var reloadedWithout = new XLWorkbook(withoutPhonetics, AtExcelDpi);
        await Assert.That(reloadedWithout.Worksheet(1).ColumnWidth).IsEqualTo(8.11).Within(0.005);
        await Assert.That(reloadedWith.Worksheet(1).ColumnWidth).IsEqualTo(8.11).Within(0.005);
    }

    private static void AddPhonetics(IXLCell cell, string fontName, double fontSize)
    {
        cell.GetRichText().AddText("みんなさんはお元気ですか。");
        cell.GetRichText().Phonetics.SetFontName(fontName).SetFontSize(fontSize);
        cell.GetRichText().Phonetics.Add("げん", 7, 8);
        cell.ShowPhonetic = true;
    }

    private static MemoryStream PlainWorkbook()
    {
        var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.Worksheets.Add("Sheet1").Cell("E5").Value = 1;
            wb.SaveAs(package);
        }

        return package;
    }

    private static XElement[] Fonts(Stream saved) =>
        XDocument.Parse(saved.ReadPart("xl/styles.xml")).Root!
            .Element(Main + "fonts")!.Elements(Main + "font").ToArray();

    private static string FontName(XElement font) =>
        $"{(string)font.Element(Main + "name")!.Attribute("val")!} {(string)font.Element(Main + "sz")!.Attribute("val")!}";
}
