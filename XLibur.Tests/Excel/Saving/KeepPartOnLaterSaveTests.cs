using System;
using System.IO;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using XLibur.Excel.IO;
using XLibur.Tests.Utils;
using SaveOptions = XLibur.Excel.SaveOptions;

namespace XLibur.Tests.Excel.Saving;

/// <summary>
/// A later save of a loaded workbook leaves a sheet's part as the last save left it, or keeps its
/// cells, as the first save does with the part the load read (#719).
/// </summary>
/// <remarks>
/// A part is marked by spelling its XML declaration as XLibur never does. A later package has the
/// mark only when every save since left the part as it was. Marking the output of the first save
/// shows whether the second one left a part the first one wrote.
/// </remarks>
public class KeepPartOnLaterSaveTests
{
    private const string Data = "xl/worksheets/sheet1.xml";

    private const string Other = "xl/worksheets/sheet2.xml";

    private const string Declaration = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>";

    /// <summary>The number 1.5 as XLibur never writes it, in the cells of the source's Data sheet.</summary>
    private const string KeptCells = "<x:v>1.50</x:v>";

    [Test]
    public async Task A_second_save_with_no_change_leaves_every_part_the_first_left()
    {
        using var source = Source();
        using var wb = new XLWorkbook(source);
        using var first = new MemoryStream();
        using var second = new MemoryStream();

        wb.SaveAs(first);
        wb.SaveAs(second);

        await Assert.That(IsMarked(first, Data)).IsTrue();
        await Assert.That(IsMarked(first, Other)).IsTrue();
        await Assert.That(IsMarked(second, Data)).IsTrue();
        await Assert.That(IsMarked(second, Other)).IsTrue();
        await AssertValues(second);
    }

    [Test]
    public async Task Editing_one_sheet_between_saves_writes_only_that_sheet()
    {
        using var source = Source();
        using var wb = new XLWorkbook(source);
        using var first = new MemoryStream();
        using var second = new MemoryStream();

        wb.SaveAs(first);
        wb.Worksheet("Other").Cell("C3").Value = 7;
        wb.SaveAs(second);

        await Assert.That(IsMarked(second, Data)).IsTrue();
        await Assert.That(IsMarked(second, Other)).IsFalse();
        await AssertValues(second);
        using var reloaded = Reload(second);
        await Assert.That(reloaded.Worksheet("Other").Cell("C3").GetDouble()).IsEqualTo(7);
    }

    [Test]
    public async Task A_part_the_first_save_wrote_is_left_by_the_second()
    {
        using var source = Source();
        using var wb = new XLWorkbook(source);
        using var first = new MemoryStream();
        using var second = new MemoryStream();

        wb.Worksheet("Data").Cell("D4").Value = 7;
        wb.SaveAs(first);
        await Assert.That(IsMarked(first, Data)).IsFalse();

        Mark(first, Data);
        wb.SaveAs(second);

        await Assert.That(IsMarked(second, Data)).IsTrue();
        await AssertValues(second);
        using var reloaded = Reload(second);
        await Assert.That(reloaded.Worksheet("Data").Cell("D4").GetDouble()).IsEqualTo(7);
    }

    [Test]
    public async Task A_part_rewritten_around_kept_cells_is_left_by_the_second_save()
    {
        using var source = Source();
        using var wb = new XLWorkbook(source);
        using var first = new MemoryStream();
        using var second = new MemoryStream();

        wb.Worksheet("Data").Column("C").Width = 30;
        wb.SaveAs(first);
        await Assert.That(IsMarked(first, Data)).IsFalse();
        await Assert.That(first.ReadPart(Data)).Contains(KeptCells);

        Mark(first, Data);
        wb.SaveAs(second);

        await Assert.That(IsMarked(second, Data)).IsTrue();
        await AssertValues(second);
    }

    [Test]
    public async Task Changes_outside_the_cells_on_two_saves_keep_the_cells_both_times()
    {
        using var source = Source();
        using var wb = new XLWorkbook(source);
        using var first = new MemoryStream();
        using var second = new MemoryStream();

        wb.Worksheet("Data").Column("C").Width = 30;
        wb.SaveAs(first);
        wb.Worksheet("Data").Column("D").Width = 20;
        wb.SaveAs(second);

        var xml = second.ReadPart(Data);
        await Assert.That(xml).Contains(KeptCells);
        await Assert.That(xml).Contains("<x:col min=\"3\" max=\"3\"");
        await Assert.That(xml).Contains("<x:col min=\"4\" max=\"4\"");
        await AssertValues(second);
    }

    [Test]
    public async Task A_change_outside_the_cells_keeps_the_cells_the_first_save_wrote()
    {
        using var source = Source();
        using var wb = new XLWorkbook(source);
        using var first = new MemoryStream();
        using var second = new MemoryStream();

        wb.Worksheet("Data").Cell("D4").Value = 7;
        wb.SaveAs(first);

        // 2 respelled in the cells the first save wrote: the second save has it only if it copied them.
        first.RewritePart(Data, xml => xml.Replace("<x:v>2</x:v>", "<x:v>2.0</x:v>", StringComparison.Ordinal));
        wb.Worksheet("Data").Column("C").Width = 30;
        wb.SaveAs(second);

        var saved = second.ReadPart(Data);
        await Assert.That(saved).Contains("<x:v>2.0</x:v>");
        await Assert.That(saved).Contains("<x:col min=\"3\" max=\"3\"");
        await AssertValues(second);
        using var reloaded = Reload(second);
        await Assert.That(reloaded.Worksheet("Data").Cell("D4").GetDouble()).IsEqualTo(7);
    }

    [Test]
    [Arguments("none")]
    [Arguments("column")]
    [Arguments("cell")]
    public async Task The_part_kept_for_the_next_save_is_the_saved_part_without_its_cells(string edit)
    {
        // Left by tier 1, rewritten around kept cells by tier 2, and written in full.
        using var source = Source();
        using var wb = new XLWorkbook(source);
        var data = wb.Worksheet("Data");
        var ws = (XLWorksheet)data;
        if (edit == "column")
            data.Column("C").Width = 30;
        else if (edit == "cell")
            data.Cell("D4").Value = 7;

        using var first = new MemoryStream();
        wb.SaveAs(first);

        var kept = ws.TakePartWithoutSheetData();
        await Assert.That(kept).IsNotNull();
        await Assert.That(kept!.AsSpan().SequenceEqual(SavedPartWithoutSheetData(first, ws.RelId!))).IsTrue();
    }

    [Test]
    public async Task Texts_the_first_save_moved_are_found_where_it_wrote_them()
    {
        // The source lists "Gone" first. Replacing it moves every other text up one, so the first
        // save writes all the cells; the second keeps them at the indices the first wrote.
        using var source = Source();
        using var wb = new XLWorkbook(source);
        using var first = new MemoryStream();
        using var second = new MemoryStream();

        wb.Worksheet("Other").Cell("A2").Value = "New";
        wb.SaveAs(first);
        await Assert.That(IsMarked(first, Data)).IsFalse();

        Mark(first, Data);
        wb.SaveAs(second);

        await Assert.That(IsMarked(second, Data)).IsTrue();
        await AssertValues(second);
        using var reloaded = Reload(second);
        await Assert.That(reloaded.Worksheet("Other").Cell("A2").GetText()).IsEqualTo("New");
    }

    [Test]
    public async Task An_edit_undone_after_the_first_save_is_written()
    {
        using var source = Source();
        using var wb = new XLWorkbook(source);
        using var first = new MemoryStream();
        using var second = new MemoryStream();

        var ws = wb.Worksheet("Data");
        ws.Cell("A2").Value = 3;
        wb.SaveAs(first);
        ws.Cell("A2").Value = 1.5;
        wb.SaveAs(second);

        await AssertValues(second);
    }

    [Test]
    public async Task A_save_that_fails_leaves_nothing_for_the_next_save()
    {
        // The encrypted save builds the whole package, then fails to create the file over a
        // directory. The next save still copies the source, so it must write the edited sheet.
        using var source = Source();
        using var wb = new XLWorkbook(source);
        using var saved = new MemoryStream();
        var directory = Directory.CreateTempSubdirectory().FullName;
        try
        {
            wb.Worksheet("Data").Cell("D4").Value = 7;
            Action failedSave = () => wb.SaveAs(directory, new SaveOptions { Password = "secret" });
            await Assert.That(failedSave).Throws<Exception>();

            wb.SaveAs(saved);
        }
        finally
        {
            Directory.Delete(directory);
        }

        await AssertValues(saved);
        using var reloaded = Reload(saved);
        await Assert.That(reloaded.Worksheet("Data").Cell("D4").GetDouble()).IsEqualTo(7);
    }

    [Test]
    public async Task Saving_in_place_twice_leaves_the_part()
    {
        using var source = Source();
        using var file = new TemporaryFile();
        using (var wb = new XLWorkbook(source))
        {
            wb.SaveAs(file.Path);
            wb.Save();
            wb.Save();
        }

        using var saved = new MemoryStream(File.ReadAllBytes(file.Path));
        await Assert.That(IsMarked(saved, Data)).IsTrue();
        await AssertValues(saved);
    }

    /// <summary>
    /// A package with sheets Data and Other, both parts marked. Data holds "Header" and "Second" in
    /// A1:B1, and 1.5 and 2 in A2:B2, and writes 1.5 as <see cref="KeptCells"/>. Other holds "Second"
    /// in A1 and "Gone" in A2; the shared string table lists "Gone" first.
    /// </summary>
    private static MemoryStream Source()
    {
        var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            var other = wb.AddWorksheet("Other");
            other.Cell("A2").Value = "Gone";
            ws.Cell("A1").Value = "Header";
            ws.Cell("B1").Value = "Second";
            ws.Cell("A2").Value = 1.5;
            ws.Cell("B2").Value = 2;
            other.Cell("A1").Value = "Second";

            wb.SaveAs(stream);
        }

        stream.RewritePart(Data, xml => xml.Replace("<x:v>1.5</x:v>", KeptCells, StringComparison.Ordinal));
        Mark(stream, Data);
        Mark(stream, Other);
        return stream;
    }

    /// <summary>Spells the part's XML declaration as no save does.</summary>
    private static void Mark(MemoryStream package, string partName) =>
        package.RewritePart(partName, xml =>
        {
            var afterDeclaration = xml.IndexOf("?>", StringComparison.Ordinal) + 2;
            return Declaration + xml[afterDeclaration..];
        });

    /// <summary>The part the sheet with <paramref name="relId"/> has in the package, cells cut out as the load cuts them.</summary>
    private static byte[] SavedPartWithoutSheetData(MemoryStream package, string relId)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        var part = (WorksheetPart)document.WorkbookPart!.GetPartById(relId);
        using var buffer = WorksheetPartBuffer.TryRead(part) ?? throw new InvalidOperationException("The saved part cannot be split.");
        return buffer.CopyWithoutSheetData();
    }

    private static bool IsMarked(Stream package, string partName) =>
        package.ReadPart(partName).StartsWith(Declaration, StringComparison.Ordinal);

    private static XLWorkbook Reload(MemoryStream package)
    {
        package.Position = 0;
        return new XLWorkbook(package);
    }

    private static async Task AssertValues(MemoryStream package)
    {
        using var wb = Reload(package);
        var ws = wb.Worksheet("Data");

        await Assert.That(ws.Cell("A1").GetText()).IsEqualTo("Header");
        await Assert.That(ws.Cell("B1").GetText()).IsEqualTo("Second");
        await Assert.That(ws.Cell("A2").GetDouble()).IsEqualTo(1.5);
        await Assert.That(ws.Cell("B2").GetDouble()).IsEqualTo(2);
        await Assert.That(wb.Worksheet("Other").Cell("A1").GetText()).IsEqualTo("Second");
    }
}
