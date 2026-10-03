using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;
using XLibur.Tests.Utils;
using SaveOptions = XLibur.Excel.SaveOptions;

namespace XLibur.Tests.Excel.Saving;

/// <summary>
/// A worksheet is written as Excel writes it: an attribute that holds its schema default is left
/// out, and so is an element left with nothing in it. What a loaded file spelled out is kept (#709).
/// </summary>
public class SchemaDefaultMarkupTests
{
    private const string Sheet1 = "xl/worksheets/sheet1.xml";

    private const string Declaration = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";

    private const string MainNs = "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"";

    /// <summary>
    /// A sheet as Excel 16 saves one with the number 1 in A1, with no element or attribute that only
    /// restates a default.
    /// </summary>
    private const string ExcelSheet =
        "<worksheet " + MainNs +
        " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"" +
        " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\"" +
        " xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">" +
        "<dimension ref=\"A1\"/><sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews>" +
        "<sheetFormatPr defaultRowHeight=\"14.4\" x14ac:dyDescent=\"0.3\"/>" +
        "<sheetData><row r=\"1\" spans=\"1:1\" x14ac:dyDescent=\"0.3\"><c r=\"A1\"><v>1</v></c></row></sheetData>" +
        "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>" +
        "</worksheet>";

    [Test]
    public async Task A_new_sheet_writes_only_what_Excel_writes()
    {
        using var saved = Save(wb => wb.AddWorksheet("Data").Cell("A1").Value = 1);
        var sheet = Root(saved);

        foreach (var name in new[] { "sheetPr", "printOptions", "pageSetup", "headerFooter", "tableParts" })
            await Assert.That(sheet.Elements().Any(e => e.Name.LocalName == name)).IsFalse().Because($"{name} only restates defaults");

        var margins = Element(sheet, "pageMargins");
        await Assert.That(Attributes(margins)).IsEqualTo("left=0.7 right=0.7 top=0.75 bottom=0.75 header=0.3 footer=0.3");
        await Assert.That(Descendant(sheet, "c").Attribute("s")).IsNull();
    }

    [Test]
    public async Task A_new_sheet_gets_Excels_normal_margins()
    {
        using var wb = new XLWorkbook();
        var margins = wb.AddWorksheet("Data").PageSetup.Margins;

        await Assert.That(margins.Left).IsEqualTo(0.7);
        await Assert.That(margins.Right).IsEqualTo(0.7);
        await Assert.That(margins.Top).IsEqualTo(0.75);
        await Assert.That(margins.Bottom).IsEqualTo(0.75);
        await Assert.That(margins.Header).IsEqualTo(0.3);
        await Assert.That(margins.Footer).IsEqualTo(0.3);
    }

    [Test]
    [Arguments(XLOutlineSummaryVLocation.Bottom, XLOutlineSummaryHLocation.Right, "")]
    [Arguments(XLOutlineSummaryVLocation.Top, XLOutlineSummaryHLocation.Right, "summaryBelow=0")]
    [Arguments(XLOutlineSummaryVLocation.Bottom, XLOutlineSummaryHLocation.Left, "summaryRight=0")]
    [Arguments(XLOutlineSummaryVLocation.Top, XLOutlineSummaryHLocation.Left, "summaryBelow=0 summaryRight=0")]
    public async Task Outline_properties_are_written_only_when_not_the_default(
        XLOutlineSummaryVLocation vertical, XLOutlineSummaryHLocation horizontal, string expected)
    {
        // As Excel 16 writes them, with or without grouped rows or columns.
        using var saved = Save(wb =>
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = 1;
            ws.Outline.SummaryVLocation = vertical;
            ws.Outline.SummaryHLocation = horizontal;
        });

        var outline = Root(saved).Descendants().FirstOrDefault(e => e.Name.LocalName == "outlinePr");
        await Assert.That(outline is null ? "" : Attributes(outline)).IsEqualTo(expected);

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("Data").Outline.SummaryVLocation).IsEqualTo(vertical);
        await Assert.That(reloaded.Worksheet("Data").Outline.SummaryHLocation).IsEqualTo(horizontal);
    }

    [Test]
    public async Task Page_setup_writes_only_what_is_not_the_default()
    {
        using var saved = Save(wb =>
        {
            var ws = wb.AddWorksheet("Data");
            ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
            ws.PageSetup.Scale = 80;
            ws.PageSetup.BlackAndWhite = true;
            ws.PageSetup.ShowGridlines = true;
        });
        var sheet = Root(saved);

        await Assert.That(Attributes(Element(sheet, "pageSetup"))).IsEqualTo("scale=80 orientation=landscape blackAndWhite=1");
        await Assert.That(Attributes(Element(sheet, "printOptions"))).IsEqualTo("gridLines=1");
    }

    [Test]
    public async Task A_header_writes_only_the_text_it_has()
    {
        using var saved = Save(wb =>
            wb.AddWorksheet("Data").PageSetup.Header.Center.AddText("Title", XLHFOccurrence.OddPages));
        var headerFooter = Element(Root(saved), "headerFooter");

        await Assert.That(Attributes(headerFooter)).IsEqualTo("");
        await Assert.That(string.Join(" ", headerFooter.Elements().Select(e => e.Name.LocalName))).IsEqualTo("oddHeader");
    }

    [Test]
    public async Task A_sheet_whose_tables_are_gone_has_no_table_parts()
    {
        using var source = Save(wb =>
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Header";
            ws.Cell("A2").Value = 1;
            ws.Range("A1:A2").CreateTable();
        });

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Data").Tables.Remove(0));

        await Assert.That(Root(saved).Elements().Any(e => e.Name.LocalName == "tableParts")).IsFalse();
    }

    [Test]
    public async Task A_validation_writes_no_default_and_reads_back_what_it_wrote()
    {
        using var saved = Save(wb =>
        {
            var dv = wb.AddWorksheet("Data").Range("A1:A3").CreateDataValidation();
            dv.WholeNumber.Between(1, 10);
            dv.IgnoreBlanks = false;
            dv.ShowErrorMessage = false;
            dv.ShowInputMessage = false;
        });

        await Assert.That(Attributes(Descendant(Root(saved), "dataValidation"))).IsEqualTo("type=whole sqref=A1:A3");

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var loaded = reloaded.Worksheet("Data").DataValidations.Single();
        await Assert.That(loaded.IgnoreBlanks).IsFalse();
        await Assert.That(loaded.ShowErrorMessage).IsFalse();
        await Assert.That(loaded.ShowInputMessage).IsFalse();
        await Assert.That(loaded.InCellDropdown).IsTrue();
    }

    [Test]
    public async Task An_unchanged_Excel_sheet_keeps_its_part()
    {
        using var source = Source(ExcelSheet);

        using var saved = LoadAndSave(source);

        await Assert.That(saved.PartBytes(Sheet1).SequenceEqual(source.PartBytes(Sheet1))).IsTrue();
    }

    [Test]
    public async Task Defaults_the_file_spelled_out_are_kept()
    {
        // As XLibur and ClosedXML wrote every sheet before #709.
        using var source = Source(ExcelSheet
            .Replace("<dimension", "<sheetPr><outlinePr summaryBelow=\"1\" summaryRight=\"1\"/></sheetPr><dimension", StringComparison.Ordinal)
            .Replace("<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>",
                "<printOptions horizontalCentered=\"0\" verticalCentered=\"0\" headings=\"0\" gridLines=\"0\"/>" +
                "<pageMargins left=\"0.75\" right=\"0.75\" top=\"0.75\" bottom=\"0.5\" header=\"0.5\" footer=\"0.75\"/>" +
                "<pageSetup paperSize=\"1\" scale=\"100\" orientation=\"default\" errors=\"displayed\"/>" +
                "<headerFooter/><tableParts count=\"0\"/>", StringComparison.Ordinal));

        using var kept = LoadAndSave(source);
        using var rewritten = LoadAndSave(source, options: new SaveOptions { RewriteUnchangedSheets = true });

        await Assert.That(kept.PartBytes(Sheet1).SequenceEqual(source.PartBytes(Sheet1))).IsTrue();
        await Assert.That(Shape(Root(rewritten))).IsEqualTo(Shape(Root(source)));
    }

    [Test]
    public async Task A_first_page_number_without_useFirstPageNumber_is_not_used()
    {
        using var source = Source(ExcelSheet.Replace("</worksheet>",
            "<pageSetup firstPageNumber=\"5\" orientation=\"portrait\"/></worksheet>", StringComparison.Ordinal));

        source.Position = 0;
        using (var wb = new XLWorkbook(source))
            await Assert.That(wb.Worksheet(1).PageSetup.FirstPageNumber).IsNull();

        using var saved = LoadAndSave(source, options: new SaveOptions { RewriteUnchangedSheets = true });
        await Assert.That(Attributes(Element(Root(saved), "pageSetup"))).IsEqualTo("firstPageNumber=5 orientation=portrait");
    }

    [Test]
    public async Task A_selection_without_an_active_cell_is_written_without_one()
    {
        using var source = Source(ExcelSheet.Replace("<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>",
            "<sheetView tabSelected=\"1\" workbookViewId=\"0\"><selection sqref=\"B2:C3\"/></sheetView>", StringComparison.Ordinal));

        using var saved = LoadAndSave(source, options: new SaveOptions { RewriteUnchangedSheets = true });

        await Assert.That(Attributes(Descendant(Root(saved), "selection"))).IsEqualTo("sqref=B2:C3");
    }

    [Test]
    public async Task A_new_selection_gets_an_active_cell()
    {
        using var source = Source(ExcelSheet.Replace("<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>",
            "<sheetView tabSelected=\"1\" workbookViewId=\"0\"><selection sqref=\"B2:C3\"/></sheetView>", StringComparison.Ordinal));

        using var saved = LoadAndSave(source, wb =>
        {
            var ws = wb.Worksheet("Data");
            ws.SelectedRanges.RemoveAll();
            ws.Range("D5:E8").Select();
        });

        await Assert.That(Attributes(Descendant(Root(saved), "selection"))).IsEqualTo("activeCell=D5 sqref=D5 D5:E8");
    }

    [Test]
    public async Task Deleting_every_column_drops_the_columns_the_file_had()
    {
        using var source = Source(ExcelSheet.Replace("<sheetData>",
            "<cols><col min=\"2\" max=\"2\" width=\"20.7109375\" hidden=\"1\" customWidth=\"1\"/></cols><sheetData>", StringComparison.Ordinal));

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Data").Columns().Delete());

        await Assert.That(Root(saved).Elements().Any(e => e.Name.LocalName == "cols")).IsFalse();
    }

    [Test]
    public async Task A_base_column_width_adds_no_default_column_width()
    {
        using var source = Source(ExcelSheet.Replace("<sheetFormatPr defaultRowHeight",
            "<sheetFormatPr baseColWidth=\"10\" defaultRowHeight", StringComparison.Ordinal));

        using var saved = LoadAndSave(source, options: new SaveOptions { RewriteUnchangedSheets = true });

        await Assert.That(Element(Root(saved), "sheetFormatPr").Attribute("defaultColWidth")).IsNull();
    }

    [Test]
    public async Task An_empty_sheet_keeps_a_width_set_for_all_columns()
    {
        // Excel writes "select all, set the width" as one range to the last column.
        using var source = Source(ExcelSheet
            .Replace("<sheetData>", "<cols><col min=\"1\" max=\"16384\" width=\"20.7109375\" customWidth=\"1\"/></cols><sheetData>", StringComparison.Ordinal)
            .Replace("<row r=\"1\" spans=\"1:1\" x14ac:dyDescent=\"0.3\"><c r=\"A1\"><v>1</v></c></row>", "", StringComparison.Ordinal));

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Data").SetTabColor(XLColor.Red));

        var col = Descendant(Root(saved), "col");
        await Assert.That(Attributes(col)).IsEqualTo("min=1 max=16384 width=20.710938 customWidth=1");
    }

    [Test]
    public async Task A_rule_with_only_a_hidden_input_message_is_kept()
    {
        // Excel leaves out showInputMessage when the message is not shown, and keeps its text.
        using var source = Source(ExcelSheet.Replace("<pageMargins",
            "<dataValidations count=\"1\"><dataValidation allowBlank=\"1\" promptTitle=\"Title\" prompt=\"Text\" sqref=\"B2\"/></dataValidations><pageMargins",
            StringComparison.Ordinal));

        using var saved = LoadAndSave(source, options: new SaveOptions { RewriteUnchangedSheets = true });

        var rule = Descendant(Root(saved), "dataValidation");
        await Assert.That(rule.Attribute("promptTitle")?.Value).IsEqualTo("Title");
        await Assert.That(rule.Attribute("prompt")?.Value).IsEqualTo("Text");
        await Assert.That(rule.Attribute("showInputMessage")).IsNull();
    }

    [Test]
    public async Task A_new_drawing_goes_after_smart_tags()
    {
        using var source = Source(ExcelSheet.Replace("</worksheet>",
            "<smartTags><cellSmartTags r=\"A1\"><cellSmartTag type=\"0\"/></cellSmartTags></smartTags></worksheet>",
            StringComparison.Ordinal));

        using var saved = LoadAndSave(source, wb =>
        {
            using var image = typeof(SchemaDefaultMarkupTests).Assembly
                .GetManifestResourceStream("XLibur.Tests.Resource.Images.ImageHandling.png")!;
            wb.Worksheet("Data").AddPicture(image);
        });

        var names = Root(saved).Elements().Select(e => e.Name.LocalName).ToList();
        await Assert.That(names.IndexOf("drawing")).IsEqualTo(names.IndexOf("smartTags") + 1);
    }

    [Test]
    public async Task A_column_with_the_default_style_on_a_styled_sheet_keeps_it()
    {
        using var saved = Save(wb =>
        {
            var ws = wb.AddWorksheet("Data");
            ws.Style.Fill.BackgroundColor = XLColor.Yellow;
            ws.Column(2).Style = XLWorkbook.DefaultStyle;
            ws.Column(2).Width = 20;
        });

        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("Data").Column(2).Style.Fill.PatternType).IsEqualTo(XLFillPatternValues.None);
    }

    [Test]
    public async Task A_copy_keeps_a_column_width_its_source_took_from_the_file()
    {
        using var source = Source(ExcelSheet.Replace("<sheetData>",
            "<cols><col min=\"1\" max=\"16384\" width=\"20.7109375\" customWidth=\"1\"/></cols><sheetData>", StringComparison.Ordinal));

        double width;
        using var saved = LoadAndSave(source, wb =>
        {
            var ws = wb.Worksheet("Data");
            ws.CopyTo("Copy");
        });
        source.Position = 0;
        using (var wb = new XLWorkbook(source))
            width = wb.Worksheet("Data").ColumnWidth;

        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("Copy").ColumnWidth).IsEqualTo(width).Within(0.01);
    }

    [Test]
    public async Task Fit_to_one_page_survives_without_a_page_setup_element()
    {
        using var saved = Save(wb => wb.AddWorksheet("Data").PageSetup.FitToPages(1, 1));

        await Assert.That(Root(saved).Elements().Any(e => e.Name.LocalName == "pageSetup")).IsFalse();

        using var reloaded = new XLWorkbook(saved);
        var pageSetup = reloaded.Worksheet("Data").PageSetup;
        await Assert.That(pageSetup.PagesWide).IsEqualTo(1);
        await Assert.That(pageSetup.PagesTall).IsEqualTo(1);
        await Assert.That(pageSetup.Scale).IsEqualTo(0);
    }

    /// <summary>A workbook with one sheet, Data, whose part is <paramref name="sheetXml"/>.</summary>
    private static MemoryStream Source(string sheetXml) =>
        Save(wb => wb.AddWorksheet("Data").Cell("A1").Value = 1).RewriteSheet1(_ => Declaration + sheetXml);

    private static MemoryStream Save(Action<XLWorkbook> build)
    {
        var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            build(wb);
            wb.SaveAs(stream);
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream LoadAndSave(MemoryStream source, Action<XLWorkbook>? edit = null,
        SaveOptions? options = null)
    {
        source.Position = 0;
        using var wb = new XLWorkbook(source);
        edit?.Invoke(wb);

        var saved = new MemoryStream();
        wb.SaveAs(saved, options ?? new SaveOptions());
        saved.Position = 0;
        return saved;
    }

    private static XElement Root(Stream package) => XDocument.Parse(package.Sheet1Xml()).Root!;

    private static XElement Element(XElement sheet, string name) =>
        sheet.Elements().Single(e => e.Name.LocalName == name);

    private static XElement Descendant(XElement sheet, string name) =>
        sheet.Descendants().First(e => e.Name.LocalName == name);

    /// <summary>The element's attributes, in document order, as <c>name=value</c>.</summary>
    private static string Attributes(XElement element) =>
        string.Join(" ", element.Attributes().Where(a => !a.IsNamespaceDeclaration)
            .Select(a => $"{a.Name.LocalName}={a.Value}"));

    /// <summary>Every element outside the cells, with its attributes.</summary>
    private static string Shape(XElement sheet) =>
        string.Join("\n", sheet.Descendants()
            .Where(e => e.Ancestors().All(a => a.Name.LocalName != "sheetData"))
            .Select(e => $"{e.Name.LocalName} {Attributes(e)}"));
}
