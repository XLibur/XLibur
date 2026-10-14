using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.IO;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Loading;

/// <summary>
/// Whitespace and other insignificant nodes between the elements of <c>&lt;sheetData&gt;</c> never
/// cut the read short (#738).
/// </summary>
/// <remarks>
/// An <see cref="System.Xml.XmlReader"/> told to ignore whitespace still reports a run of 4,096
/// characters or more as a text node. The cell reader used to stop at the first node that was not
/// the element it expected, so every row after such a run was dropped without an error.
/// <para>
/// The lengths straddle that limit: a run of 4,095 characters is dropped by the reader, one of
/// 4,096 is reported as text.
/// </para>
/// </remarks>
public class SheetDataWhitespaceLoadingTests
{
    private const string Sheet1 = "xl/worksheets/sheet1.xml";

    /// <summary>Where a test puts the whitespace.</summary>
    public enum Placement
    {
        /// <summary>Before every row, and before <c>&lt;/sheetData&gt;</c>.</summary>
        BetweenRows,

        /// <summary>Before every cell, and before every <c>&lt;/row&gt;</c>.</summary>
        BetweenCells,

        /// <summary>
        /// After every non-empty cell's start tag, between <c>&lt;f&gt;</c> and <c>&lt;v&gt;</c>, and
        /// before every <c>&lt;/c&gt;</c>.
        /// </summary>
        InsideCells,

        /// <summary>Before <c>&lt;sheetData&gt;</c> and after <c>&lt;/sheetData&gt;</c>.</summary>
        AroundSheetData,
    }

    /// <summary>Which of the loader's ways of reading the cells a test takes.</summary>
    public enum ReadPath
    {
        /// <summary>The part is buffered, and the cells are cut out of it by a byte search.</summary>
        Buffered,

        /// <summary>
        /// A comment ahead of the cells stops the byte search, so the part is streamed, as before #702.
        /// </summary>
        Streamed,
    }

    [Test]
    [MatrixDataSource]
    public async Task Whitespace_in_sheet_data_loses_no_cell(
        [Matrix(Placement.BetweenRows, Placement.BetweenCells, Placement.InsideCells, Placement.AroundSheetData)]
        Placement placement,
        [Matrix(1, 4095, 4096, 5000, 100_000)] int length,
        [Matrix(ReadPath.Buffered, ReadPath.Streamed)] ReadPath path)
    {
        using var source = Source();
        source.RewriteSheet1(xml => Pad(xml, placement, new string(' ', length), path));

        using var wb = new XLWorkbook(source);

        await AssertAllCells(wb);
    }

    [Test]
    public async Task Line_breaks_and_tabs_between_rows_lose_no_row()
    {
        using var source = Source();
        var run = string.Concat(Enumerable.Repeat("\r\n\t ", 25_000));
        source.RewriteSheet1(xml => Pad(xml, Placement.BetweenRows, run, ReadPath.Buffered));

        using var wb = new XLWorkbook(source);

        await AssertAllCells(wb);
    }

    [Test]
    public async Task Comments_and_processing_instructions_between_rows_and_cells_lose_no_cell()
    {
        using var source = Source();
        source.RewriteSheet1(xml =>
        {
            xml = Pad(xml, Placement.BetweenRows, "<!-- a comment --><?pi data?>", ReadPath.Buffered);
            xml = Pad(xml, Placement.BetweenCells, "<?pi data?><!-- a comment -->", ReadPath.Buffered);
            return Pad(xml, Placement.InsideCells, "<!-- a comment -->", ReadPath.Buffered);
        });

        using var wb = new XLWorkbook(source);

        await AssertAllCells(wb);
    }

    /// <summary>
    /// Text that is not whitespace has no place between rows or cells; it is passed over like
    /// whitespace rather than ending the read.
    /// </summary>
    [Test]
    public async Task Stray_text_between_rows_and_cells_loses_no_cell()
    {
        using var source = Source();
        source.RewriteSheet1(xml =>
        {
            xml = Pad(xml, Placement.BetweenRows, "stray", ReadPath.Buffered);
            return Pad(xml, Placement.BetweenCells, "stray", ReadPath.Buffered);
        });

        using var wb = new XLWorkbook(source);

        await AssertAllCells(wb);
    }

    /// <summary>
    /// An element the reader does not know is skipped, subtree and all, and the rows and cells after
    /// it are read.
    /// </summary>
    [Test]
    public async Task An_unknown_element_between_rows_and_cells_loses_no_cell()
    {
        const string unknown = "<u:unknown xmlns:u=\"urn:example\"><x:row r=\"9\"><x:c r=\"A9\"><x:v>9</x:v></x:c></x:row></u:unknown>";
        using var source = Source();
        source.RewriteSheet1(xml =>
        {
            xml = Pad(xml, Placement.BetweenRows, unknown, ReadPath.Buffered);
            return Pad(xml, Placement.BetweenCells, "<u:unknown xmlns:u=\"urn:example\"/>", ReadPath.Buffered);
        });

        using var wb = new XLWorkbook(source);

        await AssertAllCells(wb);
    }

    /// <summary>
    /// The save of an unchanged sheet keeps its part as the file had it (#702), and a save of a
    /// changed one writes every cell from the model: the reload has every cell either way.
    /// </summary>
    [Test]
    [MatrixDataSource]
    public async Task A_sheet_with_whitespace_in_its_part_saves_every_cell(
        [Matrix(Placement.BetweenRows, Placement.BetweenCells, Placement.InsideCells, Placement.AroundSheetData)]
        Placement placement,
        [Matrix(false, true)] bool edit)
    {
        using var source = Source();
        source.RewriteSheet1(xml => Pad(xml, placement, new string(' ', 100_000), ReadPath.Buffered));

        using var saved = LoadAndSave(source, edit);

        // Unchanged, the part is kept as the file had it (#702, tier 1), whitespace and all.
        if (!edit)
            await Assert.That(saved.PartBytes(Sheet1).AsSpan().SequenceEqual(source.PartBytes(Sheet1))).IsTrue();

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await AssertAllCells(reloaded, edit ? 1 : 0);
        if (edit)
            await Assert.That(reloaded.Worksheet("Sheet1").Cell("D5").GetText()).IsEqualTo("Edited");
    }

    /// <summary>
    /// A part larger than a load buffers (#723) is read from a stream; whitespace makes it large here,
    /// so every run is longer than the reader's buffer.
    /// </summary>
    [Test]
    [MatrixDataSource]
    public async Task A_large_part_made_large_by_whitespace_loads_and_saves_every_cell(
        [Matrix(Placement.BetweenRows, Placement.BetweenCells, Placement.InsideCells)] Placement placement,
        [Matrix(false, true)] bool edit)
    {
        using var source = Source();

        // Runs that add up to more than a load buffers, shared out between the places padded.
        source.RewriteSheet1(xml =>
        {
            const string marker = "\u0001";
            var marked = Pad(xml, placement, marker, ReadPath.Buffered);
            var places = marked.Split(marker).Length - 1;
            return marked.Replace(marker, new string(' ', WorksheetPartBuffer.MaxBufferedLength / places + 1));
        });
        await Assert.That(PartLength(source)).IsGreaterThan(WorksheetPartBuffer.MaxBufferedLength);

        source.Position = 0;
        using (var wb = new XLWorkbook(source))
            await AssertAllCells(wb);

        using var saved = LoadAndSave(source, edit);

        // Unchanged, the part is kept only when the load cut the cells out of it, on the path for a
        // large part, so a kept part shows that is the path the load took.
        if (!edit)
            await Assert.That(saved.PartBytes(Sheet1).AsSpan().SequenceEqual(source.PartBytes(Sheet1))).IsTrue();

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await AssertAllCells(reloaded, edit ? 1 : 0);
    }

    /// <summary>
    /// Sheet1 with numbers and a text in A1:A3 and a formula with its cached value in each of B1:B3,
    /// D1:E1 merged and column A widened, all written by XLibur.
    /// </summary>
    private static MemoryStream Source()
    {
        var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = "two";
            ws.Cell("A3").Value = 3;
            ws.Cell("B1").FormulaA1 = "A1*2";
            ws.Cell("B2").FormulaA1 = "LEN(A2)";
            ws.Cell("B3").FormulaA1 = "A3*2";
            ws.Range("D1:E1").Merge();
            ws.Column("A").Width = 20;
            wb.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = true });
        }

        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// Asserts that the sheet has every cell <see cref="Source"/> wrote, and nothing else but
    /// <paramref name="otherCells"/> more.
    /// </summary>
    private static async Task AssertAllCells(XLWorkbook wb, int otherCells = 0)
    {
        var ws = wb.Worksheet("Sheet1");

        await Assert.That(ws.Cell("A1").GetDouble()).IsEqualTo(1);
        await Assert.That(ws.Cell("A2").GetText()).IsEqualTo("two");
        await Assert.That(ws.Cell("A3").GetDouble()).IsEqualTo(3);

        await Assert.That(ws.Cell("B1").FormulaA1).IsEqualTo("A1*2");
        await Assert.That(ws.Cell("B2").FormulaA1).IsEqualTo("LEN(A2)");
        await Assert.That(ws.Cell("B3").FormulaA1).IsEqualTo("A3*2");

        // The cached values, as the file had them, and not recalculated.
        foreach (var address in new[] { "B1", "B2", "B3" })
            await Assert.That(ws.Cell(address).NeedsRecalculation).IsFalse();

        await Assert.That(ws.Cell("B1").CachedValue.GetNumber()).IsEqualTo(2);
        await Assert.That(ws.Cell("B2").CachedValue.GetNumber()).IsEqualTo(3);
        await Assert.That(ws.Cell("B3").CachedValue.GetNumber()).IsEqualTo(6);

        await Assert.That(ws.MergedRanges.Single().RangeAddress.ToString()).IsEqualTo("D1:E1");
        await Assert.That(ws.Column("A").Width).IsEqualTo(20);
        await Assert.That(ws.CellsUsed().Count()).IsEqualTo(6 + otherCells);
    }

    /// <summary>
    /// Inserts <paramref name="run"/> at every place <paramref name="placement"/> names in the XML of
    /// a part XLibur wrote, after a comment ahead of the cells for <see cref="ReadPath.Streamed"/>.
    /// </summary>
    private static string Pad(string xml, Placement placement, string run, ReadPath path)
    {
        if (path == ReadPath.Streamed)
            xml = InsertOnce(xml, "<x:sheetData>", "<x:sheetData><!-- no byte search -->");

        return placement switch
        {
            Placement.BetweenRows => InsertBefore(InsertBefore(xml, "<x:row ", run), "</x:sheetData>", run),
            Placement.BetweenCells => InsertBefore(InsertBefore(xml, "<x:c ", run), "</x:row>", run),
            Placement.InsideCells => InsertBefore(
                InsertAfter(Regex.Replace(xml, "<x:c [^>]*[^/]>", m => m.Value + run), "</x:f>", run),
                "</x:c>", run),
            Placement.AroundSheetData => InsertBefore(InsertAfter(xml, "</x:sheetData>", run), "<x:sheetData", run),
            _ => throw new ArgumentOutOfRangeException(nameof(placement)),
        };
    }

    private static string InsertBefore(string xml, string tag, string run) => Replace(xml, tag, run + tag);

    private static string InsertAfter(string xml, string tag, string run) => Replace(xml, tag, tag + run);

    /// <summary>Replaces every <paramref name="tag"/>, of which there has to be at least one.</summary>
    private static string Replace(string xml, string tag, string replacement)
    {
        if (!xml.Contains(tag, StringComparison.Ordinal))
            throw new InvalidOperationException($"The part has no '{tag}'.");

        return xml.Replace(tag, replacement, StringComparison.Ordinal);
    }

    private static string InsertOnce(string xml, string tag, string replacement)
    {
        var index = xml.IndexOf(tag, StringComparison.Ordinal);
        if (index < 0)
            throw new InvalidOperationException($"The part has no '{tag}'.");

        return string.Concat(xml.AsSpan(0, index), replacement, xml.AsSpan(index + tag.Length));
    }

    private static MemoryStream LoadAndSave(MemoryStream source, bool edit)
    {
        source.Position = 0;
        using var wb = new XLWorkbook(source);
        if (edit)
            wb.Worksheet("Sheet1").Cell("D5").Value = "Edited";

        var saved = new MemoryStream();
        wb.SaveAs(saved);
        saved.Position = 0;
        return saved;
    }

    private static long PartLength(Stream package)
    {
        package.Position = 0;
        using var archive = new System.IO.Compression.ZipArchive(package, System.IO.Compression.ZipArchiveMode.Read,
            leaveOpen: true);
        return archive.GetEntry(Sheet1)!.Length;
    }
}
