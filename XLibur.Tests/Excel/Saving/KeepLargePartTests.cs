using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.IO;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Saving;

/// <summary>
/// A worksheet part larger than a load buffers (<see cref="WorksheetPartBuffer.MaxBufferedLength"/>)
/// takes the unchanged-sheet fast path of #702 too (#723): the load keeps the part without its cells
/// from a stream, so the first save leaves an unchanged sheet's part as the file had it.
/// </summary>
/// <remarks>
/// The large sheet's part spells its XML declaration as XLibur never does, so the saved part has that
/// declaration only when the save left the part as it was. How the entry is stored proves nothing:
/// the source and the save both use the default level.
/// </remarks>
public class KeepLargePartTests
{
    private const string Sheet1 = "xl/worksheets/sheet1.xml";

    private const string Declaration = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>";

    /// <summary>Rows of eight numbers under a header row of texts: a part of about 18 MB.</summary>
    private const int RowCount = 52_000;

    private const int ColumnCount = 8;

    private static readonly Lazy<byte[]> LargeSource = new(BuildLargeSource);

    [Test]
    public async Task An_unchanged_large_sheet_keeps_its_part()
    {
        using var source = new MemoryStream(LargeSource.Value);

        using var saved = LoadAndSave(source);

        await Assert.That(saved.PartBytes(Sheet1).AsSpan().SequenceEqual(source.PartBytes(Sheet1))).IsTrue();
        using var reloaded = Reload(saved);
        await AssertValues(reloaded);
    }

    [Test]
    public async Task Editing_another_sheet_leaves_the_large_sheets_part()
    {
        using var source = new MemoryStream(LargeSource.Value);

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Other").Cell("A2").Value = "Added");

        await Assert.That(IsAsTheFileHadIt(saved)).IsTrue();
        using var reloaded = Reload(saved);
        await AssertValues(reloaded);
        await Assert.That(reloaded.Worksheet("Other").Cell("A2").GetText()).IsEqualTo("Added");
    }

    [Test]
    public async Task Editing_a_cell_of_the_large_sheet_writes_its_part()
    {
        using var source = new MemoryStream(LargeSource.Value);

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Data").Cell(RowCount, 3).Value = "Edited");

        await Assert.That(IsAsTheFileHadIt(saved)).IsFalse();
        using var reloaded = Reload(saved);
        var ws = reloaded.Worksheet("Data");
        await Assert.That(ws.Cell(RowCount, 3).GetText()).IsEqualTo("Edited");
        await Assert.That(ws.Cell(RowCount, 2).GetDouble()).IsEqualTo(Number(RowCount, 2));
        await Assert.That(ws.Cell(1, 1).GetText()).IsEqualTo("Header1");
    }

    [Test]
    public async Task A_change_outside_the_cells_of_the_large_sheet_writes_its_part()
    {
        using var source = new MemoryStream(LargeSource.Value);

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Data").Column("C").Width = 30);

        await Assert.That(IsAsTheFileHadIt(saved)).IsFalse();
        using var reloaded = Reload(saved);
        await AssertValues(reloaded);
        await Assert.That(reloaded.Worksheet("Data").Column("C").Width).IsEqualTo(30);
    }

    /// <summary>
    /// The load holds the markup around the cells, never the part: a part made large by rows that add
    /// nothing to the model loads in a fraction of its length.
    /// </summary>
    [Test]
    public async Task Loading_a_large_part_allocates_far_less_than_the_part()
    {
        var padding = WorksheetPartBuffer.MaxBufferedLength + 1024 * 1024;
        using var source = new MemoryStream(BuildPaddedSource(padding));
        var partLength = PartLength(source);

        // Once to warm up: the first load in a process allocates for its one-off set-up.
        source.Position = 0;
        using (new XLWorkbook(source))
        {
        }

        source.Position = 0;
        var before = GC.GetAllocatedBytesForCurrentThread();
        using var wb = new XLWorkbook(source);
        var allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        await Assert.That(partLength).IsGreaterThan(WorksheetPartBuffer.MaxBufferedLength);
        await Assert.That(allocated).IsLessThan(partLength / 8);

        using var saved = new MemoryStream();
        wb.SaveAs(saved);
        await Assert.That(IsAsTheFileHadIt(saved)).IsTrue();
        using var reloaded = Reload(saved);
        await Assert.That(reloaded.Worksheet("Data").Cell("B2").GetDouble()).IsEqualTo(2);
    }

    /// <summary>
    /// A comment among the cells could hide the end tag from a byte search, so the load streams the
    /// part as before and keeps nothing; the save writes the part from the model.
    /// </summary>
    [Test]
    public async Task A_large_part_whose_cells_cannot_be_cut_out_is_loaded_and_written_whole()
    {
        var padding = WorksheetPartBuffer.MaxBufferedLength + 1024 * 1024;
        using var source = new MemoryStream(BuildPaddedSource(padding, "<!-- </x:sheetData> -->"));

        using var saved = LoadAndSave(source);

        await Assert.That(IsAsTheFileHadIt(saved)).IsFalse();
        using var reloaded = Reload(saved);
        await Assert.That(reloaded.Worksheet("Data").Cell("A1").GetDouble()).IsEqualTo(1);
        await Assert.That(reloaded.Worksheet("Data").Cell("B2").GetDouble()).IsEqualTo(2);
    }

    private static double Number(int row, int column) => row * 10 + column + 0.25;

    /// <summary>
    /// A package with sheets Data and Other. Data holds Header1 to Header8 over
    /// <see cref="RowCount"/> rows of <see cref="Number"/>, in a part longer than a load buffers,
    /// with <see cref="Declaration"/>.
    /// </summary>
    private static byte[] BuildLargeSource()
    {
        using var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            for (var c = 1; c <= ColumnCount; c++)
                ws.Cell(1, c).Value = $"Header{c}";

            for (var r = 2; r <= RowCount; r++)
            {
                for (var c = 1; c <= ColumnCount; c++)
                    ws.Cell(r, c).Value = Number(r, c);
            }

            wb.AddWorksheet("Other").Cell("A1").Value = "Header1";
            wb.SaveAs(stream);
        }

        RewriteSheet1(stream, (xml, part) =>
        {
            var afterDeclaration = xml.AsSpan().IndexOf("?>"u8) + 2;
            part.Write(Encoding.UTF8.GetBytes(Declaration));
            part.Write(xml.AsSpan(afterDeclaration));
        });

        if (PartLength(stream) <= WorksheetPartBuffer.MaxBufferedLength)
            throw new InvalidOperationException("The large sheet's part is not larger than a load buffers.");

        return stream.ToArray();
    }

    /// <summary>
    /// A package whose Data sheet holds 1 and 2 in A1 and B2, then empty rows, which add nothing to
    /// the model, for at least <paramref name="padding"/> bytes; its part has
    /// <see cref="Declaration"/>.
    /// </summary>
    /// <remarks>
    /// A part padded with whitespace instead is covered by <c>SheetDataWhitespaceLoadingTests</c>
    /// (#738).
    /// </remarks>
    private static byte[] BuildPaddedSource(int padding, string beforePadding = "")
    {
        using var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = 1;
            ws.Cell("B2").Value = 2;
            wb.AddWorksheet("Other").Cell("A1").Value = "Text";
            wb.SaveAs(stream);
        }

        RewriteSheet1(stream, (xml, part) =>
        {
            var afterDeclaration = xml.AsSpan().IndexOf("?>"u8) + 2;
            var sheetDataEnd = xml.AsSpan().IndexOf("</x:sheetData>"u8);
            part.Write(Encoding.UTF8.GetBytes(Declaration));
            part.Write(xml.AsSpan(afterDeclaration..sheetDataEnd));

            using (var rows = new StreamWriter(part, new UTF8Encoding(false), 64 * 1024, leaveOpen: true))
            {
                rows.Write(beforePadding);
                for (var (row, written) = (3, 0L); written < padding; row++)
                {
                    var empty = $"<x:row r=\"{row}\" spans=\"1:2\"/>";
                    rows.Write(empty);
                    written += empty.Length;
                }
            }

            part.Write(xml.AsSpan(sheetDataEnd));
        });

        return stream.ToArray();
    }

    /// <summary>Replaces the first sheet's part with what <paramref name="write"/> writes from its bytes.</summary>
    private static void RewriteSheet1(MemoryStream package, Action<byte[], Stream> write)
    {
        package.Position = 0;
        using (var archive = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry(Sheet1)!;
            byte[] xml;
            using (var read = entry.Open())
            using (var buffer = new MemoryStream())
            {
                read.CopyTo(buffer);
                xml = buffer.ToArray();
            }

            entry.Delete();
            using var part = archive.CreateEntry(Sheet1, CompressionLevel.Fastest).Open();
            write(xml, part);
        }

        package.Position = 0;
    }

    private static long PartLength(Stream package)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        return archive.GetEntry(Sheet1)!.Length;
    }

    /// <summary>Does the part still have the declaration no save would write?</summary>
    private static bool IsAsTheFileHadIt(Stream package)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        using var part = archive.GetEntry(Sheet1)!.Open();
        var start = new byte[Declaration.Length];
        part.ReadExactly(start);
        return Encoding.UTF8.GetString(start) == Declaration;
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

    private static XLWorkbook Reload(MemoryStream package)
    {
        package.Position = 0;
        return new XLWorkbook(package);
    }

    private static async Task AssertValues(XLWorkbook wb)
    {
        var ws = wb.Worksheet("Data");

        await Assert.That(ws.Cell(1, 1).GetText()).IsEqualTo("Header1");
        await Assert.That(ws.Cell(1, ColumnCount).GetText()).IsEqualTo($"Header{ColumnCount}");
        await Assert.That(ws.Cell(2, 1).GetDouble()).IsEqualTo(Number(2, 1));
        await Assert.That(ws.Cell(RowCount, ColumnCount).GetDouble()).IsEqualTo(Number(RowCount, ColumnCount));
        await Assert.That(ws.LastRowUsed()!.RowNumber()).IsEqualTo(RowCount);
        await Assert.That(wb.Worksheet("Other").Cell("A1").GetText()).IsEqualTo("Header1");
    }
}
