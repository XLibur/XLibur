using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLibur.Excel.IO;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// <see cref="PartWithoutSheetDataReader"/> makes, from a stream, the copy of a worksheet part that
/// <see cref="WorksheetPartBuffer.CopyWithoutSheetData"/> makes of a buffered one, and refuses the
/// parts <see cref="SheetDataLocator"/> refuses (#723).
/// </summary>
/// <remarks>
/// Each part is read in small windows and shifted a byte at a time, so the windows split every
/// pattern the reader looks for, the end tag among them, at every position.
/// </remarks>
public class PartWithoutSheetDataReaderTests
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Declaration = "<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?>";

    private static readonly string Rows = string.Concat(Enumerable.Range(1, 40)
        .Select(r => $"<row r=\"{r}\"><c r=\"A{r}\"><v>{r}</v></c><c r=\"B{r}\" t=\"s\"><v>{r % 3}</v></c></row>"));

    public static string[] Contents() =>
    [
        "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>",
        $"<sheetData>{Rows}</sheetData>",
        $"<sheetData foo=\"a>b\" bar='/>'>{Rows}</sheetData>",
        $"<sheetData>{Rows}</sheetData \t\r\n>",
        $"<sheetData>{Rows}</sheetData{new string(' ', 600)}>",
        $"<sheetData>{Rows}</sheetDataX></sheetData>",
        "<sheetData/>",
        "<sheetData />",
        "<sheetData></sheetData>",
        $"<sheetData>{Rows}</sheetData><!-- after --><mergeCells count=\"1\"><mergeCell ref=\"A1:B1\"/></mergeCells>",
        $"<sheetPr codeName=\"sheetData\"/><sheetView name=\"x:sheetData \">sheetData x:sheetData</sheetView><sheetData>{Rows}</sheetData>",

        // Refused, as the locator refuses them.
        "<!-- <sheetData> --><sheetData/>",
        "<?pi <sheetData?><sheetData/>",
        $"<sheetData>{Rows}<row r=\"41\"><!-- </sheetData> --></row></sheetData>",
        $"<sheetData>{Rows}<row r=\"41\"><c r=\"A41\"><f><![CDATA[</sheetData>]]></f></c></row></sheetData>",
        $"<sheetData>{Rows}<row r=\"41\"><?pi?></row></sheetData>",
        $"<sheetData>{Rows}<sheetData/></sheetData>",
        $"<sheetData>{Rows}",
        "<sheetPr/>",
    ];

    [Test]
    [MethodDataSource(nameof(Contents))]
    public async Task Makes_the_copy_the_buffer_makes_or_refuses_as_the_locator_does(string content)
    {
        foreach (var prefix in new[] { string.Empty, "x" })
        {
            for (var shift = 0; shift < 64; shift++)
            {
                var xml = Part(prefix, content, shift);
                var expected = Expected(xml, out var expectedPrefix);

                foreach (var readLength in new[] { 1, 5, int.MaxValue })
                {
                    using var stream = new TrickleStream(xml, readLength);
                    var copy = PartWithoutSheetDataReader.TryCopy(stream, windowLength: 16, out var copyPrefix);

                    if (expected is null)
                    {
                        await Assert.That(copy).IsNull();
                        continue;
                    }

                    await Assert.That(copy).IsNotNull();
                    await Assert.That(Encoding.UTF8.GetString(copy!)).IsEqualTo(Encoding.UTF8.GetString(expected));
                    await Assert.That(copyPrefix).IsEqualTo(expectedPrefix);
                }
            }
        }
    }

    [Test]
    public async Task Reads_the_part_to_its_end()
    {
        var xml = Part(string.Empty, $"<sheetData>{Rows}</sheetData>", 0);
        using var stream = new MemoryStream(xml);

        var copy = PartWithoutSheetDataReader.TryCopy(stream, out _);

        await Assert.That(copy).IsNotNull();
        await Assert.That(stream.Position).IsEqualTo(xml.Length);
    }

    [Test]
    public async Task Refuses_a_copy_longer_than_the_most_it_keeps()
    {
        var padding = new string(' ', PartWithoutSheetDataReader.MaxCopyLength);
        var xml = Part(string.Empty, $"<sheetData>{Rows}</sheetData>{padding}", 0);

        var copy = PartWithoutSheetDataReader.TryCopy(new MemoryStream(xml), out _);

        await Assert.That(copy).IsNull();
    }

    [Test]
    public async Task Refuses_a_part_in_another_encoding()
    {
        var body = Encoding.Unicode.GetBytes($"<worksheet xmlns=\"{Main}\"><sheetData/></worksheet>");

        var copy = PartWithoutSheetDataReader.TryCopy(new MemoryStream(body), out _);
        var withMark = PartWithoutSheetDataReader.TryCopy(
            new MemoryStream([.. Encoding.Unicode.GetPreamble(), .. body]), out _);

        await Assert.That(copy).IsNull();
        await Assert.That(withMark).IsNull();
    }

    [Test]
    public async Task Keeps_a_utf8_byte_order_mark()
    {
        var sheetData = $"<sheetData>{Rows}</sheetData>";
        byte[] xml = [.. Encoding.UTF8.GetPreamble(), .. Part(string.Empty, sheetData, 0)];

        var copy = PartWithoutSheetDataReader.TryCopy(new MemoryStream(xml), windowLength: 16, out _);

        await Assert.That(copy).IsNotNull();
        await Assert.That(copy!.AsSpan(0, 3).SequenceEqual(Encoding.UTF8.GetPreamble())).IsTrue();
        await Assert.That(Encoding.UTF8.GetString(copy, 3, copy.Length - 3))
            .IsEqualTo(Encoding.UTF8.GetString(xml, 3, xml.Length - 3).Replace(sheetData, "<sheetData/>"));
    }

    /// <summary>
    /// What <see cref="WorksheetPartBuffer.CopyWithoutSheetData"/> would make of the part, or null
    /// when the locator refuses it.
    /// </summary>
    private static byte[]? Expected(byte[] xml, out string prefix)
    {
        prefix = string.Empty;
        if (!SheetDataLocator.TryLocate(xml, out var location))
            return null;

        prefix = location.Prefix;
        if (location.IsEmptyElement)
            return xml;

        return [.. xml[..(location.StartTagEnd - 1)], (byte)'/', (byte)'>', .. xml[location.End..]];
    }

    /// <summary>
    /// A part holding <paramref name="content"/> after a dimension, the names in it written with
    /// <paramref name="prefix"/>, and moved along by <paramref name="shift"/> spaces.
    /// </summary>
    private static byte[] Part(string prefix, string content, int shift)
    {
        var qualify = prefix.Length == 0 ? string.Empty : prefix + ":";
        var root = qualify + "worksheet";
        var declaration = prefix.Length == 0 ? $"xmlns=\"{Main}\"" : $"xmlns:{prefix}=\"{Main}\"";
        var qualified = prefix.Length == 0
            ? content
            : content.Replace("<", "<" + qualify).Replace("<" + qualify + "/", "</" + qualify)
                .Replace("<" + qualify + "!", "<!").Replace("<" + qualify + "?", "<?");

        return Encoding.UTF8.GetBytes(Declaration + $"<{root} {declaration}>{new string(' ', shift)}"
            + $"<{qualify}dimension ref=\"A1\"/>{qualified}<!-- end --></{root}>");
    }

    /// <summary>A stream that hands out at most a given number of bytes a read.</summary>
    private sealed class TrickleStream(byte[] bytes, int readLength) : MemoryStream(bytes, writable: false)
    {
        public override int Read(byte[] buffer, int offset, int count) =>
            base.Read(buffer, offset, Math.Min(count, readLength));

        public override int Read(Span<byte> buffer) =>
            base.Read(buffer[..Math.Min(buffer.Length, readLength)]);
    }
}
