using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Comments;

/// <summary>
/// VML drawing parts are written as UTF-8 without a byte order mark, like every other part XLibur
/// writes and like Excel's own <c>vmlDrawing*.vml</c> (checked against Excel 16 over COM).
/// </summary>
public class VmlDrawingEncodingTests
{
    [Test]
    public async Task Save_CommentVml_HasNoByteOrderMark()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet().Cell("A1").CreateComment().AddText("note");
        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        await AssertWellFormedWithoutBom(ms, SingleVmlPart(ms));
    }

    [Test]
    public async Task Save_HeaderFooterImageVml_HasNoByteOrderMark()
    {
        var imagePath = Path.Combine(Path.GetTempPath(), "XLibur_VmlBom_" + Guid.NewGuid().ToString("N") + ".png");
        try
        {
            await using (var resource = TestHelper.GetStreamFromResource("Images.SampleImagePng.png"))
            await using (var file = File.Create(imagePath))
                await resource.CopyToAsync(file);

            using var wb = new XLWorkbook();
            wb.AddWorksheet().PageSetup.Header.Left.AddImage(imagePath);
            using var ms = new MemoryStream();
            wb.SaveAs(ms);

            await AssertWellFormedWithoutBom(ms, SingleVmlPart(ms));
        }
        finally
        {
            File.Delete(imagePath);
        }
    }

    /// <summary>
    /// A resave rewrites an existing VML part in place: the old comment shapes are stripped, then
    /// the new ones are written over the same stream. A shorter part must not keep the old tail.
    /// </summary>
    [Test]
    public async Task Resave_WithShorterComment_RewritesVmlWithoutBomOrStaleTail()
    {
        using var first = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            ws.Cell("A1").CreateComment().AddText(new string('x', 500));
            ws.Cell("B1").CreateComment().AddText("second");
            wb.SaveAs(first);
        }

        using var second = new MemoryStream();
        using (var wb = new XLWorkbook(first))
        {
            wb.Worksheet(1).Cell("B1").GetComment().Delete();
            wb.SaveAs(second);
        }

        var before = first.PartBytes(SingleVmlPart(first)).Length;
        var vml = SingleVmlPart(second);
        var after = second.PartBytes(vml).Length;
        await Assert.That(after).IsLessThan(before);
        await AssertWellFormedWithoutBom(second, vml);

        var shapes = XDocument.Parse(second.ReadPart(vml)).Root!.Elements()
            .Count(e => e.Name.LocalName == "shape");
        await Assert.That(shapes).IsEqualTo(1);
    }

    private static string SingleVmlPart(Stream package) =>
        package.PartsUnder("xl/drawings/").Single(p => p.EndsWith(".vml", StringComparison.OrdinalIgnoreCase));

    private static async Task AssertWellFormedWithoutBom(Stream package, string partName)
    {
        var bytes = package.PartBytes(partName);
        var hasBom = bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF;
        await Assert.That(hasBom).IsFalse().Because(partName);

        // Parsing the whole part catches a stale tail left behind by an in-place rewrite.
        using var reader = new MemoryStream(bytes);
        await Assert.That(() => XDocument.Load(reader)).ThrowsNothing().Because(partName);
    }
}
