using System.Text;
using System.Threading.Tasks;
using XLibur.Excel.IO;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// <see cref="SheetDataLocator"/> cuts <c>&lt;sheetData&gt;</c> out of a worksheet part by byte
/// search. It must find exactly the element an XML reader would, and refuse any part where a byte
/// search could be fooled rather than guess.
/// </summary>
public class SheetDataLocatorTests
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Declaration = "<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?>";

    [Test]
    [Arguments("", "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>")]
    [Arguments("x", "<x:sheetData><x:row r=\"1\"><x:c r=\"A1\"><x:v>1</x:v></x:c></x:row></x:sheetData>")]
    [Arguments("", "<sheetData/>")]
    [Arguments("x", "<x:sheetData />")]
    [Arguments("", "<sheetData ></sheetData >")]
    [Arguments("", "<sheetData\r\n><row r=\"1\"/></sheetData\t>")]
    public async Task Finds_the_element_with_or_without_a_prefix(string prefix, string sheetData)
    {
        var xml = Part(prefix, sheetData);

        var found = SheetDataLocator.TryLocate(xml, out var location);

        await Assert.That(found).IsTrue();
        await Assert.That(location.Prefix).IsEqualTo(prefix);
        await Assert.That(Slice(xml, location.Start, location.End)).IsEqualTo(sheetData);
        await Assert.That(Slice(xml, location.NameStart, location.NameStart + location.NameLength))
            .IsEqualTo(prefix.Length == 0 ? "sheetData" : prefix + ":sheetData");
    }

    [Test]
    [Arguments("<x:sheetData xmlns:x=\"" + Main + "\"><x:row r=\"1\"/></x:sheetData>",
        "<x:sheetData xmlns:x=\"" + Main + "\">")]
    [Arguments("<sheetData foo=\"a>b\" ><row r=\"1\"/></sheetData>", "<sheetData foo=\"a>b\" >")]
    [Arguments("<x:sheetData xmlns:x=\"" + Main + "\" />", "<x:sheetData xmlns:x=\"" + Main + "\" />")]
    public async Task Marks_where_the_start_tag_ends(string sheetData, string startTag)
    {
        var xml = Part(string.Empty, sheetData);

        var found = SheetDataLocator.TryLocate(xml, out var location);

        await Assert.That(found).IsTrue();
        await Assert.That(Slice(xml, location.Start, location.StartTagEnd)).IsEqualTo(startTag);
        await Assert.That(location.IsEmptyElement).IsEqualTo(sheetData.EndsWith("/>"));
    }

    [Test]
    public async Task Skips_the_name_where_it_is_an_attribute_value_or_text()
    {
        const string sheetData = "<x:sheetData><x:row r=\"1\"/></x:sheetData>";
        var xml = Encoding.UTF8.GetBytes(Declaration
            + $"<x:worksheet xmlns:x=\"{Main}\"><x:sheetPr codeName=\"x:sheetData\"/>"
            + "<x:sheetViews><x:sheetView name=\"sheetData\">sheetData x:sheetData</x:sheetView></x:sheetViews>"
            + sheetData + "</x:worksheet>");

        var found = SheetDataLocator.TryLocate(xml, out var location);

        await Assert.That(found).IsTrue();
        await Assert.That(Slice(xml, location.Start, location.End)).IsEqualTo(sheetData);
    }

    [Test]
    public async Task A_closing_angle_bracket_in_a_quoted_attribute_does_not_end_the_start_tag()
    {
        const string sheetData = "<sheetData foo=\"a>b\" bar='/>'><row r=\"1\"/></sheetData>";
        var xml = Part(string.Empty, sheetData);

        var found = SheetDataLocator.TryLocate(xml, out var location);

        await Assert.That(found).IsTrue();
        await Assert.That(Slice(xml, location.Start, location.End)).IsEqualTo(sheetData);
    }

    [Test]
    public async Task Accepts_a_byte_order_mark_and_a_single_quoted_encoding()
    {
        var body = Encoding.UTF8.GetBytes("<?xml version='1.0' encoding='UTF-8'?>"
            + $"<worksheet xmlns=\"{Main}\"><sheetData/></worksheet>");
        var xml = new byte[body.Length + 3];
        xml[0] = 0xEF;
        xml[1] = 0xBB;
        xml[2] = 0xBF;
        body.CopyTo(xml, 3);

        await Assert.That(SheetDataLocator.TryLocate(xml, out _)).IsTrue();
    }

    [Test]
    [Arguments("<!-- <sheetData> --><sheetData/>")]
    [Arguments("<?pi <sheetData?><sheetData/>")]
    [Arguments("<sheetData><row r=\"1\"><!-- </sheetData> --></row></sheetData>")]
    [Arguments("<sheetData><row r=\"1\"><c r=\"A1\"><f><![CDATA[</sheetData>]]></f></c></row></sheetData>")]
    [Arguments("<sheetData><row r=\"1\"><?pi?></row></sheetData>")]
    [Arguments("<sheetData><sheetData/></sheetData>")]
    [Arguments("<sheetData><row r=\"1\"/>")]
    [Arguments("<sheetPr/>")]
    public async Task Refuses_a_part_a_byte_search_could_misread(string content)
    {
        var xml = Part(string.Empty, content);

        await Assert.That(SheetDataLocator.TryLocate(xml, out _)).IsFalse();
    }

    [Test]
    public async Task Refuses_a_part_in_another_encoding()
    {
        var utf16 = Encoding.Unicode.GetPreamble();
        var body = Encoding.Unicode.GetBytes($"<worksheet xmlns=\"{Main}\"><sheetData/></worksheet>");
        var withMark = new byte[utf16.Length + body.Length];
        utf16.CopyTo(withMark, 0);
        body.CopyTo(withMark, utf16.Length);

        var declared = Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"windows-1252\"?>"
            + $"<worksheet xmlns=\"{Main}\"><sheetData/></worksheet>");

        await Assert.That(SheetDataLocator.TryLocate(withMark, out _)).IsFalse();
        await Assert.That(SheetDataLocator.TryLocate(body, out _)).IsFalse();
        await Assert.That(SheetDataLocator.TryLocate(declared, out _)).IsFalse();
    }

    private static byte[] Part(string prefix, string sheetData)
    {
        var root = prefix.Length == 0 ? "worksheet" : prefix + ":worksheet";
        var declaration = prefix.Length == 0 ? $"xmlns=\"{Main}\"" : $"xmlns:{prefix}=\"{Main}\"";
        var element = prefix.Length == 0 ? "dimension" : prefix + ":dimension";
        return Encoding.UTF8.GetBytes(Declaration
            + $"<{root} {declaration}><{element} ref=\"A1\"/>{sheetData}<!-- after --></{root}>");
    }

    private static string Slice(byte[] xml, int start, int end) => Encoding.UTF8.GetString(xml, start, end - start);
}
