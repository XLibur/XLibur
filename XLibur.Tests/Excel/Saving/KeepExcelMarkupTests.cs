using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;
using XLibur.Excel.Rows;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Saving;

/// <summary>
/// A sheet from Excel keeps its part on save when nothing in it changed (#702, tier 1), though the
/// file spells things as XLibur would not: more digits in a width, a default spelled out, ranges in
/// another order. The writers keep what the file spelled out and add nothing the model does not
/// need (#722).
/// </summary>
public class KeepExcelMarkupTests
{
    private const string Sheet1 = "xl/worksheets/sheet1.xml";

    private const string Declaration = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";

    private const string MainNs = "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"";

    /// <summary>A sheet as Excel 16 saves one with the number 1 in A1.</summary>
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

    /// <summary>A sheet from a producer that declares no <c>x14ac</c> and no markup compatibility.</summary>
    private const string PlainSheet =
        "<worksheet " + MainNs + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
        "<dimension ref=\"A1\"/><sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews>" +
        "<sheetFormatPr defaultRowHeight=\"14.4\"/>" +
        "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>" +
        "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>" +
        "</worksheet>";

    private const string SheetView = "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>";

    /// <summary>Columns B and D with a width Excel spells with eight decimals, and C between them with none.</summary>
    private const string GappedColumns =
        "<cols><col min=\"2\" max=\"2\" width=\"14.85546875\" customWidth=\"1\"/>" +
        "<col min=\"4\" max=\"4\" width=\"14.85546875\" customWidth=\"1\"/></cols><sheetData>";

    private const string Merges =
        "<mergeCells count=\"2\"><mergeCell ref=\"D1:E1\"/><mergeCell ref=\"B1:C1\"/></mergeCells><pageMargins";

    [Test]
    public async Task A_sheet_without_x14ac_is_not_given_it()
    {
        using var source = Source(PlainSheet);

        using var saved = LoadAndSave(source);

        await AssertKept(source, saved);
    }

    [Test]
    [Arguments("widths with eight decimals, and no column between them", "<sheetData>", GappedColumns)]
    [Arguments("a defaultColWidth with eight decimals", "<sheetFormatPr defaultRowHeight", "<sheetFormatPr defaultColWidth=\"14.42578125\" defaultRowHeight")]
    [Arguments("a zoom and a top-left cell that restate defaults", SheetView, "<sheetView tabSelected=\"1\" topLeftCell=\"A1\" zoomScaleNormal=\"100\" workbookViewId=\"0\"/>")]
    [Arguments("rightToLeft=\"0\"", SheetView, "<sheetView tabSelected=\"1\" rightToLeft=\"0\" workbookViewId=\"0\"/>")]
    [Arguments("a selected column spelled by its cells", SheetView, "<sheetView tabSelected=\"1\" workbookViewId=\"0\"><selection activeCell=\"B1\" sqref=\"B1:B1048576\"/></sheetView>")]
    [Arguments("a frozen pane with one selection", SheetView, "<sheetView tabSelected=\"1\" workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/><selection pane=\"bottomLeft\" activeCell=\"A2\" sqref=\"A2\"/></sheetView>")]
    [Arguments("a margin with seventeen digits", "header=\"0.3\"", "header=\"0.51180555555555496\"")]
    [Arguments("a resolution past int.MaxValue", "</worksheet>", "<pageSetup orientation=\"portrait\" horizontalDpi=\"4294967294\"/></worksheet>")]
    [Arguments("a protection that protects nothing", "<pageMargins", "<sheetProtection selectLockedCells=\"1\"/><pageMargins")]
    [Arguments("a protection that spells out a default", "<pageMargins", "<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\"/><pageMargins")]
    [Arguments("merges in the order the file wrote them", "<pageMargins", Merges)]
    [Arguments("a validation of a single cell", "<pageMargins", "<dataValidations count=\"1\"><dataValidation type=\"whole\" allowBlank=\"1\" sqref=\"B2\"><formula1>1</formula1><formula2>9</formula2></dataValidation></dataValidations><pageMargins")]
    public async Task An_unchanged_sheet_keeps_its_part(string spelling, string find, string replace)
    {
        using var source = Source(ExcelSheet.Replace(find, replace, StringComparison.Ordinal));

        using var saved = LoadAndSave(source);

        await Assert.That(saved.PartBytes(Sheet1).SequenceEqual(source.PartBytes(Sheet1))).IsTrue()
            .Because($"a sheet with {spelling} did not change, so its part is kept as the file had it");
    }

    [Test]
    public async Task A_scale_beside_a_fit_to_pages_keeps_the_part()
    {
        // Excel keeps the scale for when the fit is turned off.
        using var source = Source(ExcelSheet
            .Replace("<dimension", "<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr><dimension", StringComparison.Ordinal)
            .Replace("</worksheet>", "<pageSetup scale=\"80\" orientation=\"portrait\"/></worksheet>", StringComparison.Ordinal));

        using var saved = LoadAndSave(source);

        await AssertKept(source, saved);
    }

    [Test]
    public async Task A_rewritten_part_does_not_join_columns_across_a_gap()
    {
        using var source = Source(ExcelSheet.Replace("<sheetData>", GappedColumns, StringComparison.Ordinal));

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Data").SetTabColor(XLColor.Red));

        var columns = Root(saved).Descendants().Where(e => e.Name.LocalName == "col").Select(Attributes).ToList();
        await Assert.That(columns).IsEquivalentTo(new[]
        {
            "min=2 max=2 width=14.85546875 customWidth=1",
            "min=4 max=4 width=14.85546875 customWidth=1",
        });

        using var reloaded = Reload(saved);
        var ws = reloaded.Worksheet("Data");
        await Assert.That(ws.Column(3).Width).IsEqualTo(ws.ColumnWidth);
        await Assert.That(ws.Column(2).Width).IsNotEqualTo(ws.ColumnWidth);
    }

    [Test]
    public async Task A_margin_keeps_its_digits_only_while_it_is_the_same_number()
    {
        using var source = Source(ExcelSheet.Replace("header=\"0.3\"", "header=\"0.51180555555555496\"", StringComparison.Ordinal));

        // A change elsewhere writes the part, and the margins with it.
        using var unchanged = LoadAndSave(source, wb => wb.Worksheet("Data").SetTabColor(XLColor.Red));
        using var changed = LoadAndSave(source, wb => wb.Worksheet("Data").PageSetup.Margins.Header = 0.5);

        await Assert.That(Margins(unchanged).Attribute("header")!.Value).IsEqualTo("0.51180555555555496");
        await Assert.That(Margins(changed).Attribute("header")!.Value).IsEqualTo("0.5");

        static XElement Margins(Stream package) =>
            Root(package).Elements().Single(e => e.Name.LocalName == "pageMargins");
    }

    [Test]
    public async Task A_changed_width_is_written()
    {
        using var source = Source(ExcelSheet.Replace("<sheetData>", GappedColumns, StringComparison.Ordinal));

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Data").Column(2).Width = 30);

        var column = Root(saved).Descendants().First(e => e.Name.LocalName == "col");
        await Assert.That(column.Attribute("width")!.Value).IsNotEqualTo("14.85546875");
        using var reloaded = Reload(saved);
        await Assert.That(reloaded.Worksheet("Data").Column(2).Width).IsEqualTo(30).Within(1e-6);
    }

    [Test]
    public async Task A_row_with_dyDescent_declares_x14ac_as_ignorable()
    {
        // The row declares x14ac itself, so the root has no use for it until the cells are written
        // from the model.
        using var source = Source(PlainSheet.Replace("<row r=\"1\">",
            "<row r=\"1\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" x14ac:dyDescent=\"0.3\">",
            StringComparison.Ordinal));

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Data").Cell("B1").Value = 2);

        var root = Root(saved);
        await Assert.That(root.Attribute(XNamespace.Get("http://schemas.openxmlformats.org/markup-compatibility/2006") + "Ignorable")?.Value)
            .IsEqualTo("x14ac");
        using var reloaded = Reload(saved);
        await Assert.That(((XLRow)reloaded.Worksheet("Data").Row(1)).DyDescent).IsEqualTo(0.3);
    }

    [Test]
    public async Task A_changed_selection_is_written()
    {
        using var source = Source(ExcelSheet.Replace(SheetView,
            "<sheetView tabSelected=\"1\" workbookViewId=\"0\"><selection activeCell=\"B1\" sqref=\"B1:B1048576\"/></sheetView>",
            StringComparison.Ordinal));

        using var saved = LoadAndSave(source, wb =>
        {
            var ws = wb.Worksheet("Data");
            ws.SetActiveCell("C3");
            ws.Range("C3:D4").Select();
        });

        var selection = Root(saved).Descendants().Single(e => e.Name.LocalName == "selection");
        await Assert.That(selection.Attribute("activeCell")!.Value).IsEqualTo("C3");
        await Assert.That(selection.Attribute("sqref")!.Value).Contains("C3:D4");
    }

    [Test]
    public async Task Unmerging_writes_the_merges_that_are_left()
    {
        using var source = Source(ExcelSheet.Replace("<pageMargins", Merges, StringComparison.Ordinal));

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Data").Range("B1:C1").Unmerge());

        var merges = Root(saved).Descendants().Where(e => e.Name.LocalName == "mergeCell")
            .Select(e => e.Attribute("ref")!.Value).ToList();
        await Assert.That(merges).IsEquivalentTo(new[] { "D1:E1" });
    }

    [Test]
    public async Task A_changed_setting_of_a_protection_that_protects_nothing_is_written()
    {
        using var source = Source(ExcelSheet.Replace("<pageMargins", "<sheetProtection selectLockedCells=\"1\"/><pageMargins",
            StringComparison.Ordinal));

        using var saved = LoadAndSave(source, wb =>
            wb.Worksheet("Data").Protection.AllowElement(XLSheetProtectionElements.SelectLockedCells));

        var protection = Root(saved).Descendants().Single(e => e.Name.LocalName == "sheetProtection");
        await Assert.That(protection.Attribute("selectLockedCells")).IsNull();
        using var reloaded = Reload(saved);
        var ws = reloaded.Worksheet("Data");
        await Assert.That(ws.Protection.IsProtected).IsFalse();
        await Assert.That(ws.Protection.AllowedElements.HasFlag(XLSheetProtectionElements.SelectLockedCells)).IsTrue();
    }

    [Test]
    public async Task A_malformed_merge_count_is_written_again()
    {
        using var source = Source(ExcelSheet.Replace("<pageMargins",
            "<mergeCells count=\"x\"><mergeCell ref=\"B1:C1\"/></mergeCells><pageMargins", StringComparison.Ordinal));

        using var saved = LoadAndSave(source);

        var merges = Root(saved).Descendants().Single(e => e.Name.LocalName == "mergeCells");
        await Assert.That(merges.Attribute("count")!.Value).IsEqualTo("1");
    }

    [Test]
    public async Task An_active_cell_set_on_a_view_whose_first_selection_is_not_the_active_panes_is_written()
    {
        // Excel puts the cursor in the active pane's selection, here the last one. The load reads the
        // first, so the model cannot tell setting the active cell to B1 from leaving it.
        using var source = Source(ExcelSheet.Replace(SheetView,
            "<sheetView tabSelected=\"1\" workbookViewId=\"0\">" +
            "<pane xSplit=\"1\" ySplit=\"1\" topLeftCell=\"B2\" activePane=\"bottomRight\" state=\"frozen\"/>" +
            "<selection pane=\"topRight\" activeCell=\"B1\" sqref=\"B1\"/>" +
            "<selection pane=\"bottomLeft\" activeCell=\"A2\" sqref=\"A2\"/>" +
            "<selection pane=\"bottomRight\" activeCell=\"C5\" sqref=\"C5\"/></sheetView>",
            StringComparison.Ordinal));

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Data").SetActiveCell("B1"));

        // The view is written again, so the active pane's selection has the cell, whichever pane that is.
        var root = Root(saved);
        var activePane = root.Descendants().Single(e => e.Name.LocalName == "pane").Attribute("activePane")?.Value ?? "topLeft";
        var active = root.Descendants()
            .Single(e => e.Name.LocalName == "selection" && (e.Attribute("pane")?.Value ?? "topLeft") == activePane);
        await Assert.That(active.Attribute("activeCell")!.Value).IsEqualTo("B1");
    }

    [Test]
    public async Task Unprotecting_a_sheet_removes_its_protection()
    {
        using var source = Source(ExcelSheet.Replace("<pageMargins", "<sheetProtection sheet=\"1\" scenarios=\"1\"/><pageMargins",
            StringComparison.Ordinal));

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Data").Unprotect());

        await Assert.That(Root(saved).Descendants().Any(e => e.Name.LocalName == "sheetProtection")).IsFalse();
    }

    [Test]
    public async Task A_column_keeps_a_style_index_that_repeats_another()
    {
        using var source = SourceWithRepeatedColumnStyle(out var repeated);

        using var saved = LoadAndSave(source);

        await AssertKept(source, saved);
        using var rewritten = LoadAndSave(source, wb => wb.Worksheet("Data").SetTabColor(XLColor.Red));
        var column = ColumnB(rewritten);
        await Assert.That(column.Attribute("style")!.Value).IsEqualTo(repeated);
        using var reloaded = Reload(rewritten);
        await Assert.That(reloaded.Worksheet("Data").Column(2).Style.Font.Bold).IsTrue();
    }

    [Test]
    public async Task A_column_whose_style_changed_is_written_with_the_new_one()
    {
        using var source = SourceWithRepeatedColumnStyle(out var repeated);

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Data").Column(2).Style.Font.Italic = true);

        var column = ColumnB(saved);
        await Assert.That(column.Attribute("style")!.Value).IsNotEqualTo(repeated);
        using var reloaded = Reload(saved);
        await Assert.That(reloaded.Worksheet("Data").Column(2).Style.Font.Italic).IsTrue();
        await Assert.That(reloaded.Worksheet("Data").Column(2).Style.Font.Bold).IsTrue();
    }

    /// <summary>
    /// A workbook whose column B is bold through a <c>&lt;cellXfs&gt;</c> index that repeats an
    /// earlier one, as files Excel has edited often have. <paramref name="repeated"/> is that index.
    /// </summary>
    private static MemoryStream SourceWithRepeatedColumnStyle(out string repeated)
    {
        var source = Save(wb =>
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = 1;
            ws.Column(2).Width = 20;
            ws.Column(2).Style.Font.Bold = true;
        });

        var original = Regex.Match(source.Sheet1Xml(), "<x:col min=\"2\" max=\"2\"[^>]*style=\"(\\d+)\"").Groups[1].Value;
        var index = -1;
        source.RewritePart("xl/styles.xml", xml =>
        {
            var styles = XDocument.Parse(xml);
            var cellXfs = styles.Root!.Elements().Single(e => e.Name.LocalName == "cellXfs");
            var xfs = cellXfs.Elements().ToList();
            cellXfs.Add(new XElement(xfs[int.Parse(original, CultureInfo.InvariantCulture)]));
            index = xfs.Count;
            cellXfs.SetAttributeValue("count", xfs.Count + 1);
            return styles.Declaration + styles.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        });

        repeated = index.ToString(CultureInfo.InvariantCulture);
        var styleIndex = repeated;
        source.RewriteSheet1(xml => Regex.Replace(xml, "(<x:col min=\"2\" max=\"2\"[^>]*style=\")\\d+\"",
            m => m.Groups[1].Value + styleIndex + "\""));
        return source;
    }

    private static async Task AssertKept(MemoryStream source, MemoryStream saved)
    {
        await Assert.That(saved.PartBytes(Sheet1).SequenceEqual(source.PartBytes(Sheet1))).IsTrue()
            .Because("the sheet did not change, so its part is kept as the file had it");
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

    private static MemoryStream LoadAndSave(MemoryStream source, Action<XLWorkbook>? edit = null)
    {
        source.Position = 0;
        using var wb = new XLWorkbook(source);
        edit?.Invoke(wb);

        var saved = new MemoryStream();
        wb.SaveAs(saved);
        saved.Position = 0;
        return saved;
    }

    private static XLWorkbook Reload(MemoryStream saved)
    {
        saved.Position = 0;
        return new XLWorkbook(saved);
    }

    private static XElement Root(Stream package) => XDocument.Parse(package.Sheet1Xml()).Root!;

    private static XElement ColumnB(Stream package) =>
        Root(package).Descendants().Single(e => e.Name.LocalName == "col" && e.Attribute("min")?.Value == "2");

    /// <summary>The element's attributes, in document order, as <c>name=value</c>.</summary>
    private static string Attributes(XElement element) =>
        string.Join(" ", element.Attributes().Where(a => !a.IsNamespaceDeclaration)
            .Select(a => $"{a.Name.LocalName}={a.Value}"));
}
