using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Tests.Utils;
using SaveOptions = XLibur.Excel.SaveOptions;

namespace XLibur.Tests.Excel.Saving;

/// <summary>
/// The first save of a loaded workbook leaves the part of an unchanged sheet as the file had it, when
/// the save would write the same markup around the cells (#702, tier 1).
/// </summary>
/// <remarks>
/// The source spells the sheet's XML declaration as XLibur never does, so the saved part has that
/// declaration only when the save left the part as it was.
/// </remarks>
public class KeepUnchangedPartTests
{
    private const string Sheet1 = "xl/worksheets/sheet1.xml";

    private const string Declaration = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>";

    /// <summary>The number 1.5 as XLibur never writes it, in the cells of the source's Data sheet.</summary>
    private const string KeptCells = "<x:v>1.50</x:v>";

    [Test]
    public async Task An_unchanged_sheet_keeps_its_part_as_the_file_had_it()
    {
        using var source = Source(_ => { });

        using var saved = LoadAndSave(source);

        await Assert.That(saved.PartBytes(Sheet1).SequenceEqual(source.PartBytes(Sheet1))).IsTrue();
        await AssertValues(saved);
    }

    [Test]
    public async Task A_sheet_with_a_table_comment_and_merge_keeps_its_part()
    {
        using var source = Source(ws =>
        {
            ws.Range("A1:B2").CreateTable();
            ws.Cell("D1").CreateComment().AddText("Note");
            ws.Range("D5:E5").Merge();
        });

        using var saved = LoadAndSave(source);

        await Assert.That(IsAsTheFileHadIt(saved, Sheet1)).IsTrue();
        await AssertValues(saved);
        using var reloaded = new XLWorkbook(saved);
        var ws = reloaded.Worksheet("Data");
        await Assert.That(ws.Tables.Count()).IsEqualTo(1);
        await Assert.That(ws.Cell("D1").GetComment().Text).IsEqualTo("Note");
        await Assert.That(ws.MergedRanges.Single().RangeAddress.ToString()).IsEqualTo("D5:E5");
    }

    [Test]
    public async Task A_sheet_with_an_external_hyperlink_keeps_its_part()
    {
        // The hyperlink is written with the relationship id it was loaded with (#705).
        using var source = Source(ws => ws.Cell("D1").SetHyperlink(new XLHyperlink("http://example.com/")));

        using var saved = LoadAndSave(source);

        await Assert.That(IsAsTheFileHadIt(saved, Sheet1)).IsTrue();
        await AssertValues(saved);
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("Data").Cell("D1").GetHyperlink().ExternalAddress!.ToString())
            .IsEqualTo("http://example.com/");
    }

    [Test]
    public async Task Editing_another_sheet_leaves_an_unchanged_sheets_part()
    {
        using var source = Source(_ => { });

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Other").Cell("C3").Value = 7);

        await Assert.That(IsAsTheFileHadIt(saved, Sheet1)).IsTrue();
        await Assert.That(IsAsTheFileHadIt(saved, "xl/worksheets/sheet2.xml")).IsFalse();
        await AssertValues(saved);
    }

    [Test]
    public async Task A_change_outside_the_cells_writes_the_part_and_keeps_the_cells()
    {
        using var source = Source(_ => { });

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Data").Column("C").Width = 30);

        await Assert.That(IsAsTheFileHadIt(saved, Sheet1)).IsFalse();
        await Assert.That(saved.Sheet1Xml()).Contains(KeptCells);
        await Assert.That(saved.Sheet1Xml()).Contains("<x:col min=\"3\" max=\"3\"");
    }

    [Test]
    public async Task A_change_to_a_view_writes_the_part()
    {
        using var source = Source(_ => { });

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Data").SheetView.FreezeRows(1));

        await Assert.That(IsAsTheFileHadIt(saved, Sheet1)).IsFalse();
        await Assert.That(saved.Sheet1Xml()).Contains("state=\"frozen\"");
    }

    [Test]
    public async Task An_edited_sheet_writes_the_part()
    {
        using var source = Source(_ => { });

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Data").Cell("D4").Value = 7);

        await Assert.That(IsAsTheFileHadIt(saved, Sheet1)).IsFalse();
        await Assert.That(saved.Sheet1Xml()).DoesNotContain(KeptCells);
    }

    [Test]
    public async Task A_sheet_with_a_formula_writes_the_part()
    {
        using var source = Source(ws => ws.Cell("C2").FormulaA1 = "A2*2");

        using var saved = LoadAndSave(source);

        await Assert.That(IsAsTheFileHadIt(saved, Sheet1)).IsFalse();
    }

    [Test]
    public async Task RewriteUnchangedSheets_writes_the_part()
    {
        using var source = Source(_ => { });

        using var saved = LoadAndSave(source, options: new SaveOptions { RewriteUnchangedSheets = true });

        await Assert.That(IsAsTheFileHadIt(saved, Sheet1)).IsFalse();
        await Assert.That(saved.Sheet1Xml()).DoesNotContain(KeptCells);
    }

    [Test]
    public async Task Only_the_first_save_keeps_the_part()
    {
        using var source = Source(_ => { });
        using var wb = new XLWorkbook(source);
        using var first = new MemoryStream();
        using var second = new MemoryStream();

        wb.SaveAs(first);
        wb.SaveAs(second);

        await Assert.That(IsAsTheFileHadIt(first, Sheet1)).IsTrue();
        await Assert.That(IsAsTheFileHadIt(second, Sheet1)).IsFalse();
        await AssertValues(second);
    }

    /// <summary>
    /// A package with sheets Data and Other. Data holds "Header" and "Second" in A1:B1, and 1.5 and 2
    /// in A2:B2. Its part has <see cref="Declaration"/>, and writes 1.5 as <see cref="KeptCells"/>.
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

        return stream.RewritePart(Sheet1, xml =>
        {
            var afterDeclaration = xml.IndexOf("?>", StringComparison.Ordinal) + 2;
            return Declaration + xml[afterDeclaration..].Replace("<x:v>1.5</x:v>", KeptCells, StringComparison.Ordinal);
        });
    }

    /// <summary>Does the part still have the declaration no save would write?</summary>
    private static bool IsAsTheFileHadIt(Stream package, string partName) =>
        package.ReadPart(partName).StartsWith(Declaration, StringComparison.Ordinal);

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
}
