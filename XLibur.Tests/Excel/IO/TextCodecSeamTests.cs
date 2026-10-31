using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.Streaming;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// The <c>_xHHHH_</c> text codec is applied where the text element is written and read, so it
/// holds for every combination of write path (DOM, streaming) and storage (shared, inline).
/// </summary>
public class TextCodecSeamTests
{
    private const string Control = "a\u0001b";
    private const string Lookalike = "_x0041_";
    private const string Padded = "  padded  ";
    // Loaded text has its line breaks normalised to the platform's, so use that form.
    private static readonly string Newline = "line1" + Environment.NewLine + "line2";
    private const string Ordinary = "plain";

    private static string[] Texts => [Control, Lookalike, Padded, Newline, Ordinary, "pre_x000D_post"];

    private static MemoryStream SaveDom(bool shareStrings)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        for (var i = 0; i < Texts.Length; i++)
        {
            var cell = ws.Cell(i + 1, 1);
            cell.Value = Texts[i];
            cell.ShareString = shareStrings;
        }

        var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms;
    }

    private static MemoryStream SaveStreaming(XLStreamingStringStorage storage)
    {
        var ms = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(ms, new XLStreamingOptions { StringStorage = storage }))
        {
            var sheet = wb.AddWorksheet("S");
            foreach (var text in Texts)
                sheet.AppendRow(text);
            wb.Finish();
        }

        return ms;
    }

    private static async Task AssertRoundTrip(MemoryStream ms)
    {
        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);
        var ws = loaded.Worksheet("S");
        for (var i = 0; i < Texts.Length; i++)
            await Assert.That(Lf(ws.Cell(i + 1, 1).GetString())).IsEqualTo(Lf(Texts[i]));
    }

    // Line-break normalisation on load differs between shared and inline strings and is not this
    // codec's concern, so compare line breaks in one form.
    private static string Lf(string s) => s.Replace("\r\n", "\n");

    private static string ReadEntry(MemoryStream ms, string name)
    {
        ms.Position = 0;
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: true);
        var entry = zip.GetEntry(name) ?? throw new InvalidOperationException(name);
        using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    [Test]
    public async Task DomSharedStringsRoundTrip()
    {
        using var ms = SaveDom(shareStrings: true);
        await AssertRoundTrip(ms);
    }

    [Test]
    public async Task DomInlineStringsRoundTrip()
    {
        using var ms = SaveDom(shareStrings: false);
        await AssertRoundTrip(ms);
    }

    [Test]
    public async Task StreamingSharedStringsRoundTrip()
    {
        using var ms = SaveStreaming(XLStreamingStringStorage.SharedStrings);
        await AssertRoundTrip(ms);
    }

    [Test]
    public async Task StreamingInlineStringsRoundTrip()
    {
        using var ms = SaveStreaming(XLStreamingStringStorage.Inline);
        await AssertRoundTrip(ms);
    }

    [Test]
    public async Task DomInlineStringEscapesLiteralEscapeSequenceInTheBytes()
    {
        using var ms = SaveDom(shareStrings: false);
        var sheet = ReadEntry(ms, "xl/worksheets/sheet1.xml");
        await Assert.That(sheet).Contains(">_x005F_x0041_</");
        await Assert.That(sheet).Contains("a_x0001_b");
    }

    [Test]
    public async Task StreamingInlineStringEscapesLiteralEscapeSequenceInTheBytes()
    {
        using var ms = SaveStreaming(XLStreamingStringStorage.Inline);
        var sheet = ReadEntry(ms, "xl/worksheets/sheet1.xml");
        await Assert.That(sheet).Contains(">_x005F_x0041_</");
        await Assert.That(sheet).Contains("a_x0001_b");
    }

    /// <summary>
    /// The shared string path used to encode at its caller; moving the encode to the element must
    /// not escape twice.
    /// </summary>
    [Test]
    public async Task SharedStringsAreEscapedExactlyOnce()
    {
        using var ms = SaveDom(shareStrings: true);
        var sst = ReadEntry(ms, "xl/sharedStrings.xml");
        // Text that looks like an escape is escaped once, never twice.
        await Assert.That(sst).Contains(">_x005F_x0041_</");
        await Assert.That(sst).Contains(">pre_x005F_x000D_post</");
        await Assert.That(sst).DoesNotContain("_x005F_x005F_");
        await Assert.That(sst).DoesNotContain("_x005F_x0001_");
    }

    [Test]
    public async Task RichTextRunsKeepTheirBehaviour()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").GetRichText().AddText("bold").SetBold().AddText(" rest");
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);
        await Assert.That(loaded.Worksheet("S").Cell("A1").GetString()).IsEqualTo("bold rest");
    }
}
