using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;
using XLibur.Tests.Excel.Drawings.Shapes;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Drawings;

/// <summary>
/// A drawing whose pictures, or whose shapes, share a drawing id loads, and the later duplicate is
/// given a fresh id that the first save writes into the drawing, as Excel does when it repairs such a
/// file (#859).
/// </summary>
public class DuplicateDrawingIdTests
{
    private static readonly XNamespace Xdr = ShapeCreationTests.Xdr;

    /// <summary>The picture fixture with Rotated given Plain's id, 3. The highest id in it is 25.</summary>
    private static MemoryStream DuplicatePictures() =>
        TestHelper.OpenResource(PictureFidelityTests.Fixture).RewritePart(PictureFidelityTests.Drawing1,
            xml => xml.Replace("id=\"5\" name=\"Rotated\"", "id=\"3\" name=\"Rotated\""));

    /// <summary>The shapes fixture with Empty given Plain TextBox's id, 2. The highest id left in it is the connector's, 8.</summary>
    private static MemoryStream DuplicateShapes() =>
        TestHelper.OpenResource(ShapeCreationTests.Fixture).RewritePart(ShapeCreationTests.Drawing1,
            xml => xml.Replace("id=\"9\" name=\"Empty\"", "id=\"2\" name=\"Empty\""));

    [Test]
    public async Task Pictures_that_share_an_id_load_and_the_later_one_takes_a_fresh_id()
    {
        using var wb = new XLWorkbook(DuplicatePictures());
        var ws = wb.Worksheet(1);

        await Assert.That(ws.Pictures.Count).IsEqualTo(12);
        await Assert.That(ws.Picture("Plain").Id).IsEqualTo(3);
        await Assert.That(ws.Picture("Rotated").Id).IsEqualTo(26);
    }

    [Test]
    public async Task The_first_save_writes_the_fresh_id_and_nothing_else()
    {
        using var source = DuplicatePictures();
        var before = source.ReadPart(PictureFidelityTests.Drawing1);
        using var wb = new XLWorkbook(source);
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        await Assert.That(XmlChangeSet.Between(before, saved.ReadPart(PictureFidelityTests.Drawing1)).ToString())
            .IsEqualTo(XmlChangeSet.Expect(
                "modified /xdr:wsDr[1]/xdr:twoCellAnchor[2]/xdr:pic[1]/xdr:nvPicPr[1]/xdr:cNvPr[1] @id: '3' -> '26'"));
    }

    [Test]
    public async Task An_edit_to_the_duplicate_lands_on_it_and_not_on_the_first()
    {
        using var source = DuplicatePictures();
        var before = source.ReadPart(PictureFidelityTests.Drawing1);
        using var wb = new XLWorkbook(source);
        wb.Worksheet(1).Picture("Rotated").Name = "Turned";
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        await Assert.That(XmlChangeSet.Between(before, saved.ReadPart(PictureFidelityTests.Drawing1)).ToString())
            .IsEqualTo(XmlChangeSet.Expect(
                "modified /xdr:wsDr[1]/xdr:twoCellAnchor[2]/xdr:pic[1]/xdr:nvPicPr[1]/xdr:cNvPr[1] @id: '3' -> '26', @name: 'Rotated' -> 'Turned'"));
    }

    [Test]
    public async Task Deleting_the_duplicate_removes_it_and_not_the_first()
    {
        using var wb = new XLWorkbook(DuplicatePictures());
        wb.Worksheet(1).Pictures.Delete("Rotated");
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        var names = XDocument.Parse(saved.ReadPart(PictureFidelityTests.Drawing1)).Descendants(Xdr + "cNvPr")
            .Select(p => (string)p.Attribute("name")!).ToList();
        await Assert.That(names).Contains("Plain");
        await Assert.That(names).DoesNotContain("Rotated");
    }

    [Test]
    public async Task A_second_save_leaves_the_repaired_drawing_as_the_first_wrote_it()
    {
        using var wb = new XLWorkbook(DuplicatePictures());
        using var first = new MemoryStream();
        wb.SaveAs(first);
        using var second = new MemoryStream();
        wb.SaveAs(second);

        await Assert.That(Convert.ToHexString(second.PartBytes(PictureFidelityTests.Drawing1)))
            .IsEqualTo(Convert.ToHexString(first.PartBytes(PictureFidelityTests.Drawing1)));
    }

    [Test]
    public async Task Shapes_with_id_zero_each_take_a_fresh_id()
    {
        using var source = TestHelper.OpenResource(ShapeCreationTests.Fixture).RewritePart(ShapeCreationTests.Drawing1,
            xml => xml.Replace("id=\"9\" name=\"Empty\"", "id=\"0\" name=\"Empty\"")
                .Replace("id=\"5\" name=\"Arc\"", "id=\"0\" name=\"Arc\""));
        using var wb = new XLWorkbook(source);
        var shapes = wb.Worksheet(1).Shapes;

        await Assert.That(shapes.Shape("Arc").Id).IsEqualTo(9);
        await Assert.That(shapes.Shape("Empty").Id).IsEqualTo(10);

        shapes.Delete("Empty");
        using var saved = new MemoryStream();
        wb.SaveAs(saved);
        var names = XDocument.Parse(saved.ReadPart(ShapeCreationTests.Drawing1)).Descendants(Xdr + "cNvPr")
            .Select(p => (string)p.Attribute("name")!).ToList();
        await Assert.That(names).Contains("Arc");
        await Assert.That(names).DoesNotContain("Empty");
    }

    [Test]
    public async Task Shapes_that_share_an_id_load_and_an_edit_lands_on_the_one_edited()
    {
        using var source = DuplicateShapes();
        var before = source.ReadPart(ShapeCreationTests.Drawing1);
        using var wb = new XLWorkbook(source);
        var shapes = wb.Worksheet(1).Shapes;
        await Assert.That(shapes.Shape("Plain TextBox").Id).IsEqualTo(2);
        await Assert.That(shapes.Shape("Empty").Id).IsEqualTo(9);

        shapes.Shape("Empty").TextBody.AddParagraph("Filled");
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        await Assert.That(XmlChangeSet.Between(before, saved.ReadPart(ShapeCreationTests.Drawing1)).ToString())
            .IsEqualTo(XmlChangeSet.Expect(
                "modified /xdr:wsDr[1]/xdr:twoCellAnchor[8]/xdr:sp[1]/xdr:nvSpPr[1]/xdr:cNvPr[1] @id: '2' -> '9'",
                "added /xdr:wsDr[1]/xdr:twoCellAnchor[8]/xdr:sp[1]/xdr:txBody[1]/a:p[1]/a:r[1] (subtree)"));
    }
}
