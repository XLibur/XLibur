using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Saving;

/// <summary>
/// Content a load and save used to drop or change (#910).
/// </summary>
public class SmallRoundTripLossTests
{
    private const string ConditionalFormatThemeColor = @"TryToLoad\LO\xlsx\condformat_theme_color.xlsx";

    [Test]
    public async Task Data_bar_theme_colour_keeps_its_tint()
    {
        using var saved = TestHelper.LoadAndSave(ConditionalFormatThemeColor);
        using var wb = new XLWorkbook(saved);

        var color = wb.Worksheet(1).ConditionalFormats.Single().Colors[1];
        await Assert.That(color.ThemeColor).IsEqualTo(XLThemeColor.Accent1);
        await Assert.That(color.ThemeTint).IsEqualTo(0.39997558519241921).Within(1e-12);
    }

    [Test]
    public async Task Colour_scale_theme_colours_keep_their_tint()
    {
        using var saved = TestHelper.LoadAndSave(ConditionalFormatThemeColor);
        using var wb = new XLWorkbook(saved);

        var colors = wb.Worksheet(2).ConditionalFormats.Single().Colors;
        await Assert.That(colors[1].ThemeColor).IsEqualTo(XLThemeColor.Accent4);
        await Assert.That(colors[1].ThemeTint).IsEqualTo(0.59999389629810485).Within(1e-12);
        await Assert.That(colors[2].ThemeColor).IsEqualTo(XLThemeColor.Accent1);
        await Assert.That(colors[2].ThemeTint).IsEqualTo(0.39997558519241921).Within(1e-12);
    }

    [Test]
    public async Task Header_text_with_no_section_code_stays_centred()
    {
        // The fixture's first sheet has <oddHeader>&A</oddHeader> and <oddFooter>Page &P</oddFooter>.
        using var saved = TestHelper.LoadAndSave(@"TryToLoad\LO\xlsm\tdf111974.xlsm");
        using var wb = new XLWorkbook(saved);

        var pageSetup = wb.Worksheet("XXXXX").PageSetup;
        await Assert.That(pageSetup.Header.Center.GetText(XLHFOccurrence.OddPages)).IsEqualTo("&A");
        await Assert.That(pageSetup.Header.Left.GetText(XLHFOccurrence.OddPages)).IsEmpty();
        await Assert.That(pageSetup.Footer.Center.GetText(XLHFOccurrence.OddPages)).IsEqualTo("Page &P");
        await Assert.That(pageSetup.Footer.Left.GetText(XLHFOccurrence.OddPages)).IsEmpty();
    }

    [Test]
    public async Task List_validation_in_alternate_content_keeps_its_items()
    {
        // The rule holds <x12ac:list>1,"2,3",4</x12ac:list> with <formula1>"1,2,3,4"</formula1> as the
        // fallback for readers that do not know x12ac. Excel writes this form when an item has a comma.
        using var saved = TestHelper.LoadAndSave(@"TryToLoad\LO\xlsx\tdf104310-2.xlsx");

        var dataValidation = XDocument.Parse(saved.Sheet1Xml()).Descendants(Main + "dataValidation").Single();
        await Assert.That(dataValidation.Descendants(X12ac + "list").Single().Value).IsEqualTo("1,\"2,3\",4");
        await Assert.That(dataValidation.Descendants(Main + "formula1").Single().Value).IsEqualTo("\"1,2,3,4\"");

        using var wb = new XLWorkbook(saved);
        var rule = wb.Worksheet(1).DataValidations.Single();
        await Assert.That(rule.AllowedValues).IsEqualTo(XLAllowedValues.List);
        await Assert.That(rule.MinValue).IsEqualTo("\"1,2,3,4\"");
    }

    [Test]
    public async Task List_validation_changed_after_load_drops_the_stale_x12ac_list()
    {
        using var source = TestHelper.LoadAndSave(@"TryToLoad\LO\xlsx\tdf104310-2.xlsx");
        using var saved = source.LoadAndSave(wb => wb.Worksheet(1).DataValidations.Single().List("\"a,b\""));

        var dataValidation = XDocument.Parse(saved.Sheet1Xml()).Descendants(Main + "dataValidation").Single();
        await Assert.That(dataValidation.Descendants(X12ac + "list").Any()).IsFalse();
        await Assert.That(dataValidation.Element(Main + "formula1")!.Value).IsEqualTo("\"a,b\"");
    }

    [Test]
    public async Task Xlm_macro_names_keep_their_function_and_xlm_flags()
    {
        using var saved = TestHelper.LoadAndSave(@"TryToLoad\LO\xlsm\tdf111974.xlsm");

        var plant2 = DefinedName(saved, "Plant2");
        await Assert.That((string?)plant2.Attribute("function")).IsEqualTo("1");
        await Assert.That((string?)plant2.Attribute("xlm")).IsEqualTo("1");
        await Assert.That(plant2.Value).IsEqualTo("\"Cuautitlan\"");
    }

    [Test]
    [Arguments(null)]
    [Arguments("0")]
    public async Task Macro_attributes_of_a_defined_name_survive_a_round_trip(string? localSheetId)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        if (localSheetId is null)
            wb.DefinedNames.Add("AddTax", "1");
        else
            ws.DefinedNames.Add("AddTax", "1");

        using var built = new MemoryStream();
        wb.SaveAs(built);
        const string macroAttributes =
            "name=\"AddTax\" function=\"1\" vbProcedure=\"1\" functionGroupId=\"14\" shortcutKey=\"q\" " +
            "description=\"Adds the tax\" help=\"tax.chm!12\"";
        built.RewriteWorkbook(xml => xml.Replace("name=\"AddTax\"", macroAttributes));

        using var saved = built.LoadAndSave();

        var name = DefinedName(saved, "AddTax");
        await Assert.That((string?)name.Attribute("function")).IsEqualTo("1");
        await Assert.That((string?)name.Attribute("vbProcedure")).IsEqualTo("1");
        await Assert.That((string?)name.Attribute("functionGroupId")).IsEqualTo("14");
        await Assert.That((string?)name.Attribute("shortcutKey")).IsEqualTo("q");
        await Assert.That((string?)name.Attribute("description")).IsEqualTo("Adds the tax");
        await Assert.That((string?)name.Attribute("help")).IsEqualTo("tax.chm!12");
        await Assert.That((string?)name.Attribute("localSheetId")).IsEqualTo(localSheetId);
    }

    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private static readonly XNamespace X12ac = "http://schemas.microsoft.com/office/spreadsheetml/2011/1/ac";

    private static XElement DefinedName(Stream package, string name) =>
        XDocument.Parse(package.WorkbookXml())
            .Descendants(Main + "definedName")
            .Single(e => (string?)e.Attribute("name") == name);
}
