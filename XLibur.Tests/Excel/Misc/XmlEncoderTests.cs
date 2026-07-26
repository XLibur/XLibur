using System.IO;
using XLibur.Excel;
using XLibur.Utils;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Misc;

public class XmlEncoderTest
{
    [Test]
    public async Task TestControlChars()
    {
        await Assert.That(XmlEncoder.EncodeString("\u0001 \u0002 \u0003 \u0004")).IsEqualTo("_x0001_ _x0002_ _x0003_ _x0004_");
        await Assert.That(XmlEncoder.EncodeString("\u0005 \u0006 \u0007 \u0008")).IsEqualTo("_x0005_ _x0006_ _x0007_ _x0008_");

        await Assert.That(XmlEncoder.DecodeString("_x0001_ _x0002_ _x0003_ _x0004_")).IsEqualTo("\u0001 \u0002 \u0003 \u0004");
        await Assert.That(XmlEncoder.DecodeString("_x0005_ _x0006_ _x0007_ _x0008_")).IsEqualTo("\u0005 \u0006 \u0007 \u0008");
        await Assert.That(XmlEncoder.DecodeString("_xaaBB_ _xAAbb_")).IsEqualTo("\uAABB \uAABB");

        // https://github.com/XLibur/XLibur/issues/1154
        await Assert.That(XmlEncoder.DecodeString("_Xceed_Something")).IsEqualTo("_Xceed_Something");
    }

    [Test]
    public async Task AstralUnicodeCharsAreWrittenWithoutOpenXmlEncoding()
    {
        using var sr = new StreamReader(TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Unicode\let_it_go_in_emoji.txt")));
        var surrogateEmoji = sr.ReadToEnd();

        await TestHelper.CreateAndCompare(() =>
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            var cell = ws.FirstCell();
            cell.Value = "This emoji version of Let It Go from Frozen:";
            cell.CellBelow().Value = surrogateEmoji;

            return wb;
        }, @"Other\Unicode\let_it_go_in_emoji-outputfile.xlsx");
    }
}
