using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.Streaming;
using XLibur.Utils;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// The text codec at the places the first pass missed: a formula's cached string, rich text runs
/// (cell and comment), inline strings with runs, and escape lookalikes that overlap.
/// </summary>
public class TextCodecRemainingTests
{
    private const string Control = "a\u0001b";

    private static string ReadEntry(MemoryStream ms, string name)
    {
        ms.Position = 0;
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: true);
        var entry = zip.GetEntry(name) ?? throw new InvalidOperationException(name);
        using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    [Test]
    public async Task FormulaCachedStringWithControlCharacterRoundTripsThroughDom()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").FormulaA1 = "\"a\"&CHAR(1)&\"b\"";
            ws.Cell("A2").FormulaA1 = "\"_x\"&\"0041_\"";
            _ = ws.Cell("A1").Value; // Evaluating stores the cached results the save writes.
            _ = ws.Cell("A2").Value;
            wb.SaveAs(ms);
        }

        var sheet = ReadEntry(ms, "xl/worksheets/sheet1.xml");
        await Assert.That(sheet).Contains(">a_x0001_b</");
        await Assert.That(sheet).Contains(">_x005F_x0041_</");

        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);
        await Assert.That(loaded.Worksheet("S").Cell("A1").CachedValue.GetText()).IsEqualTo(Control);
        await Assert.That(loaded.Worksheet("S").Cell("A2").CachedValue.GetText()).IsEqualTo("_x0041_");
    }

    [Test]
    public async Task FormulaCachedStringWithControlCharacterRoundTripsThroughStreaming()
    {
        using var ms = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(ms))
        {
            var sheet = wb.AddWorksheet("S");
            sheet.AddRow().Formula("X", Control);
            sheet.AddRow().Formula("X", "_x0041_");
            wb.Finish();
        }

        var xml = ReadEntry(ms, "xl/worksheets/sheet1.xml");
        await Assert.That(xml).Contains(">a_x0001_b</");
        await Assert.That(xml).Contains(">_x005F_x0041_</");

        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);
        await Assert.That(loaded.Worksheet("S").Cell("A1").CachedValue.GetText()).IsEqualTo(Control);
        await Assert.That(loaded.Worksheet("S").Cell("A2").CachedValue.GetText()).IsEqualTo("_x0041_");
    }

    [Test]
    public async Task RichTextRunsRoundTripControlCharactersAndLookalikes()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").GetRichText().AddText(Control).SetBold().AddText("_x0041_");
            ws.Cell("A2").GetRichText().AddText(Control).SetBold().AddText("_x0041_");
            ws.Cell("A2").ShareString = false;
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);
        await Assert.That(loaded.Worksheet("S").Cell("A1").GetString()).IsEqualTo(Control + "_x0041_");
        await Assert.That(loaded.Worksheet("S").Cell("A2").GetString()).IsEqualTo(Control + "_x0041_");
        await Assert.That(ReadEntry(ms, "xl/sharedStrings.xml")).Contains("_x005F_x0041_");
    }

    [Test]
    public async Task CommentRunsRoundTripControlCharactersAndLookalikes()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").CreateComment().AddText(Control + "_x0041_");
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);
        await Assert.That(loaded.Worksheet("S").Cell("A1").GetComment().Text).IsEqualTo(Control + "_x0041_");
    }

    [Test]
    public async Task ThirdPartyInlineStringWithRunsIsDecoded()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").Value = "MARKER";
            ws.Cell("A1").ShareString = false;
            wb.SaveAs(ms);
        }

        var patched = Patch(ms, "xl/worksheets/sheet1.xml",
            sheet => Regex.Replace(sheet, "<x:t>MARKER</x:t>",
                "<x:r><x:t>a_x0001_b</x:t></x:r><x:r><x:t>_x005F_x0041_</x:t></x:r>"));

        using var loaded = new XLWorkbook(patched);
        await Assert.That(loaded.Worksheet("S").Cell("A1").GetString()).IsEqualTo(Control + "_x0041_");
    }

    private static MemoryStream Patch(MemoryStream source, string entryName, Func<string, string> edit)
    {
        var original = ReadEntry(source, entryName);
        var edited = edit(original);
        if (edited == original)
            throw new InvalidOperationException("Patch did not change the part: " + original);

        var result = new MemoryStream();
        source.Position = 0;
        using (var input = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true))
        using (var output = new ZipArchive(result, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in input.Entries)
            {
                var copy = output.CreateEntry(entry.FullName);
                using var to = copy.Open();
                if (entry.FullName == entryName)
                {
                    var bytes = new UTF8Encoding(false).GetBytes(edited);
                    to.Write(bytes, 0, bytes.Length);
                }
                else
                {
                    using var from = entry.Open();
                    from.CopyTo(to);
                }
            }
        }

        result.Position = 0;
        return result;
    }

    [Test]
    [Arguments("_x005F_x0041_")]
    [Arguments("_x005F_")]
    [Arguments("_x005F_x005F_")]
    [Arguments("_x0041__x0042_")]
    [Arguments("_x0041_x0042_")]
    [Arguments("__x0041_")]
    [Arguments("_x_x0041_")]
    [Arguments("a_x005F_x0001_b")]
    [Arguments("_x0001_\u0001_x0001_")]
    [Arguments("_X0041_")]
    [Arguments("_x005F_x005F_x0041_")]
    public async Task CodecRoundTripsLookalikes(string text)
    {
        await Assert.That(XmlEncoder.DecodeString(XmlEncoder.EncodeString(text))).IsEqualTo(text);
    }

    [Test]
    public async Task CodecRoundTripsRandomTokenSoup()
    {
        string[] tokens = ["_", "x", "X", "0041", "005F", "0001", "_x0041_", "_x005F_", "\u0001", "a", "F", "_x005f_"];
        var rnd = new Random(845);
        var failures = new System.Collections.Generic.List<string>();
        for (var n = 0; n < 5000; n++)
        {
            var sb = new StringBuilder();
            for (var k = rnd.Next(1, 8); k > 0; k--)
                sb.Append(tokens[rnd.Next(tokens.Length)]);
            var text = sb.ToString();
            var decoded = XmlEncoder.DecodeString(XmlEncoder.EncodeString(text));
            if (decoded != text)
                failures.Add($"'{text}' came back as '{decoded}'");
        }

        await Assert.That(failures).IsEmpty();
    }

    [Test]
    public async Task SharedAndInlineStringsNormaliseLineBreaksTheSame()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").Value = "x\ny";
            ws.Cell("A2").Value = "x\ny";
            ws.Cell("A2").ShareString = false;
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);
        var shared = loaded.Worksheet("S").Cell("A1").GetString();
        var inline = loaded.Worksheet("S").Cell("A2").GetString();
        await Assert.That(inline).IsEqualTo(shared);
    }
}
