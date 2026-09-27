using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;
using XLibur.Tests.Utils;
using SaveOptions = XLibur.Excel.SaveOptions;

namespace XLibur.Tests.Excel.Saving;

/// <summary>
/// The first save of a loaded workbook copies the cells of a sheet that has not changed since the
/// load as the file had them, instead of writing them from the model (#702, tier 2).
/// </summary>
/// <remarks>
/// A source file writes the number 1.5 as <c>1.50</c>. XLibur would write it as <c>1.5</c>, so the
/// saved part holds <c>1.50</c> only when the cells were kept.
/// </remarks>
public class KeepSheetDataTests
{
    private const string Kept = "<x:v>1.50</x:v>";

    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    [Test]
    public async Task An_unchanged_sheet_is_saved_with_its_cells_as_the_file_had_them()
    {
        using var source = Source(_ => { });
        var sourceCells = SheetDataContent(source.Sheet1Xml());

        using var saved = LoadAndSave(source);

        await Assert.That(SheetDataContent(saved.Sheet1Xml())).IsEqualTo(sourceCells);
        await AssertValues(saved);
    }

    [Test]
    public async Task RewriteUnchangedSheets_writes_the_cells_from_the_model()
    {
        using var source = Source(_ => { });

        using var saved = LoadAndSave(source, options: new SaveOptions { RewriteUnchangedSheets = true });

        await Assert.That(saved.Sheet1Xml()).DoesNotContain(Kept);
        await AssertValues(saved);
    }

    [Test]
    public async Task An_edited_sheet_is_written_from_the_model()
    {
        using var source = Source(_ => { });

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Data").Cell("D4").Value = 7);

        await Assert.That(saved.Sheet1Xml()).DoesNotContain(Kept);
        await AssertValues(saved);
    }

    [Test]
    public async Task Editing_another_sheet_leaves_an_unchanged_sheet_kept()
    {
        using var source = Source(_ => { });

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Other").Cell("C3").Value = 7);

        await Assert.That(saved.Sheet1Xml()).Contains(Kept);
        await AssertValues(saved);
    }

    [Test]
    public async Task Only_the_first_save_keeps_the_cells()
    {
        // The markup the load kept describes the part only until the first save replaces it.
        using var source = Source(_ => { });
        using var wb = new XLWorkbook(source);
        using var first = new MemoryStream();
        using var second = new MemoryStream();

        wb.SaveAs(first);
        wb.SaveAs(second);

        await Assert.That(first.Sheet1Xml()).Contains(Kept);
        await Assert.That(second.Sheet1Xml()).DoesNotContain(Kept);
        await AssertValues(second);
    }

    [Test]
    public async Task A_sheet_with_a_formula_is_written_from_the_model()
    {
        using var source = Source(ws => ws.Cell("C2").FormulaA1 = "A2*2");

        using var saved = LoadAndSave(source);

        await Assert.That(saved.Sheet1Xml()).DoesNotContain(Kept);
        await AssertValues(saved);
    }

    [Test]
    public async Task A_sheet_with_a_totals_row_is_written_from_the_model()
    {
        using var source = Source(ws => ws.Range("A1:B2").CreateTable().ShowTotalsRow = true);

        using var saved = LoadAndSave(source);

        await Assert.That(saved.Sheet1Xml()).DoesNotContain(Kept);
    }

    [Test]
    public async Task Header_names_the_load_gives_a_table_are_saved()
    {
        // The file leaves the table's header cells empty. Loading the table names them after its
        // fields, so the cells no longer match the file, and the save has to write the names.
        var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A2").Value = 1.5;
            ws.Cell("B2").Value = 2;
            ws.Range("A1:B2").CreateTable();
            wb.SaveAs(stream);
        }

        stream.RewriteSheet1(xml => Regex.Replace(xml, "<x:c r=\"[AB]1\"[^>]*>.*?</x:c>", string.Empty));
        await Assert.That(stream.Sheet1Xml()).DoesNotContain("r=\"A1\"");

        using var saved = LoadAndSave(stream);

        await Assert.That(saved.Sheet1Xml()).Contains("r=\"A1\"");
    }

    [Test]
    public async Task A_sheet_with_a_table_comment_and_merge_keeps_its_cells()
    {
        using var source = Source(ws =>
        {
            ws.Range("A1:B2").CreateTable();
            ws.Cell("D1").CreateComment().AddText("Note");
            ws.Range("D5:E5").Merge();
        });

        using var saved = LoadAndSave(source);

        await Assert.That(saved.Sheet1Xml()).Contains(Kept);
        await AssertValues(saved);
    }

    [Test]
    public async Task Saving_a_comment_does_not_materialise_the_rows_and_columns_it_spans()
    {
        // The comment box reaches past column D and row 1. Sizing it used to create every column and
        // row it crossed, which gave the cells under them styles of their own.
        using var source = Source(ws => ws.Cell("D1").CreateComment().AddText("Note"));
        source.Position = 0;
        using var wb = new XLWorkbook(source);
        var ws = (XLWorksheet)wb.Worksheet("Data");
        var columns = ws.Internals.ColumnsCollection.Count;
        var rows = ws.Internals.RowsCollection.Count;

        wb.SaveAs(new MemoryStream());

        await Assert.That(ws.Internals.ColumnsCollection.Count).IsEqualTo(columns);
        await Assert.That(ws.Internals.RowsCollection.Count).IsEqualTo(rows);
    }

    [Test]
    public async Task A_shared_string_that_moves_makes_the_save_write_the_cells_from_the_model()
    {
        // An item no cell uses, ahead of the others, is not written again, so every later index
        // moves down by one.
        using var source = Source(_ => { });
        source.RewritePart("xl/sharedStrings.xml", xml =>
        {
            var first = xml.IndexOf("<x:si>", StringComparison.Ordinal);
            return xml[..first] + "<x:si><x:t>Unused</x:t></x:si>" + xml[first..];
        });
        source.RewriteSheet1(xml => xml
            .Replace("t=\"s\"><x:v>1</x:v>", "t=\"s\"><x:v>2</x:v>", StringComparison.Ordinal)
            .Replace("t=\"s\"><x:v>0</x:v>", "t=\"s\"><x:v>1</x:v>", StringComparison.Ordinal));
        source.RewritePart("xl/worksheets/sheet2.xml", xml =>
            xml.Replace("t=\"s\"><x:v>1</x:v>", "t=\"s\"><x:v>2</x:v>", StringComparison.Ordinal));

        using var saved = LoadAndSave(source);

        await Assert.That(saved.Sheet1Xml()).DoesNotContain(Kept);
        await AssertValues(saved);
    }

    [Test]
    public async Task A_text_the_file_lists_twice_makes_the_save_write_the_cells_from_the_model()
    {
        // Items 0 and 2 are both "Header", and a cell uses each. The save writes the text once, so
        // a kept cell naming item 2 would point past the end of the table.
        using var source = Source(_ => { });
        source.RewritePart("xl/sharedStrings.xml", xml =>
        {
            var end = xml.LastIndexOf("</x:si>", StringComparison.Ordinal) + "</x:si>".Length;
            return xml[..end] + "<x:si><x:t>Header</x:t></x:si>" + xml[end..];
        });
        source.RewriteSheet1(xml =>
        {
            var endOfFirstRow = xml.IndexOf("</x:row>", StringComparison.Ordinal);
            return xml[..endOfFirstRow] + "<x:c r=\"C1\" t=\"s\"><x:v>2</x:v></x:c>" + xml[endOfFirstRow..];
        });

        using var saved = LoadAndSave(source);
        using var reloaded = new XLWorkbook(saved);

        await Assert.That(saved.Sheet1Xml()).DoesNotContain(Kept);
        await Assert.That(reloaded.Worksheet("Data").Cell("C1").GetText()).IsEqualTo("Header");
    }

    [Test]
    public async Task Cells_in_the_default_namespace_are_kept_in_it()
    {
        // As Excel writes a sheet: the default namespace, row spans and dyDescent.
        using var source = Source(_ => { });
        source.RewriteSheet1(_ => ExcelStyleSheet(
            """<sheetData><row r="1" spans="1:2" x14ac:dyDescent="0.25"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>""" + "\n" +
            """<row r="2" spans="1:2" x14ac:dyDescent="0.25"><c r="A2"><v>1.50</v></c><c r="B2" t="n"><v>2</v></c></row></sheetData>"""));

        using var saved = LoadAndSave(source, options: new SaveOptions { ValidatePackage = true });
        var xml = saved.Sheet1Xml();

        await Assert.That(xml).Contains("<v>1.50</v></c><c r=\"B2\" t=\"n\"><v>2</v></c></row></sheetData>");
        await Assert.That(xml).Contains("</row>\n<row r=\"2\"");
        await Assert.That(RowsInMainNamespace(xml)).IsEqualTo(2);
        await AssertValues(saved);
    }

    [Test]
    public async Task A_prefix_declared_on_sheetData_itself_is_kept()
    {
        using var source = Source(_ => { });
        source.RewriteSheet1(_ => ExcelStyleSheet(
            """<s:sheetData xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><s:row r="1"><s:c r="A1" t="s"><s:v>0</s:v></s:c><s:c r="B1" t="s"><s:v>1</s:v></s:c></s:row>""" +
            """<s:row r="2"><s:c r="A2"><s:v>1.50</s:v></s:c><s:c r="B2"><s:v>2</s:v></s:c></s:row></s:sheetData>"""));

        using var saved = LoadAndSave(source);
        var xml = saved.Sheet1Xml();

        await Assert.That(xml).Contains("<s:v>1.50</s:v>");
        await Assert.That(RowsInMainNamespace(xml)).IsEqualTo(2);
        await AssertValues(saved);
    }

    /// <summary>
    /// A package with sheets Data and Other, whose Data sheet writes 1.5 as <c>1.50</c>. Data holds
    /// "Header" and "Second" in A1:B1, shared strings 0 and 1, and 1.5 and 2 in A2:B2.
    /// </summary>
    private static MemoryStream Source(Action<IXLWorksheet> addToData)
    {
        var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Header";
            ws.Cell("B1").Value = "Second";
            ws.Cell("A2").Value = 1.5;
            ws.Cell("B2").Value = 2;
            addToData(ws);

            var other = wb.AddWorksheet("Other");
            other.Cell("A1").Value = "Second";

            wb.SaveAs(stream);
        }

        return stream.RewriteSheet1(xml => xml.Replace("<x:v>1.5</x:v>", Kept, StringComparison.Ordinal));
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

    private static async Task AssertValues(MemoryStream package)
    {
        package.Position = 0;
        using var wb = new XLWorkbook(package);
        var ws = wb.Worksheet("Data");

        await Assert.That(ws.Cell("A1").GetText()).IsEqualTo("Header");
        await Assert.That(ws.Cell("B1").GetText()).IsEqualTo("Second");
        await Assert.That(ws.Cell("A2").GetDouble()).IsEqualTo(1.5);
        await Assert.That(ws.Cell("B2").GetDouble()).IsEqualTo(2);
        await Assert.That(wb.Worksheet("Other").Cell("A1").GetText()).IsEqualTo("Second");
    }

    /// <summary>Everything between the <c>&lt;sheetData&gt;</c> start and end tags.</summary>
    private static string SheetDataContent(string xml)
    {
        var start = xml.IndexOf('>', xml.IndexOf("sheetData", StringComparison.Ordinal)) + 1;
        var end = xml.LastIndexOf("</", xml.LastIndexOf("sheetData>", StringComparison.Ordinal), StringComparison.Ordinal);
        return xml[start..end];
    }

    private static int RowsInMainNamespace(string xml) =>
        XDocument.Parse(xml).Root!.Element(Main + "sheetData")!.Elements(Main + "row").Count();

    private static string ExcelStyleSheet(string sheetData) =>
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""" + "\r\n" +
        """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">""" +
        """<dimension ref="A1:B2"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>""" +
        sheetData +
        """<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>""";
}
