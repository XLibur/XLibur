using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;
using XLibur.Excel.Drawings;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Drawings;

/// <summary>
/// A loaded picture is written back by editing the anchor it was loaded from, so whatever Excel put
/// in it — rotation, a border, a shadow, a crop, a recolour — survives a save (spec 17 task 1).
/// Before, every save rebuilt every picture's anchor with a bare <c>spPr</c> and <c>blipFill</c>.
/// </summary>
/// <remarks>
/// The fixture was made by Excel over COM: twelve pictures of one image, each styled one way in
/// Excel's own Format Picture model, and a value in A30.
/// </remarks>
public class PictureFidelityTests
{
    internal const string Fixture = @"Other\Pictures\ExcelStyledPictures.xlsx";
    internal const string Drawing1 = "xl/drawings/drawing1.xml";

    [Test]
    public async Task A_save_that_edits_no_picture_leaves_the_drawing_part_as_it_was()
    {
        using var source = TestHelper.OpenResource(Fixture);
        var before = source.PartBytes(Drawing1);
        using var wb = new XLWorkbook(source);
        wb.Worksheet(1).Cell("A30").Value = "edited";

        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        await Assert.That(Convert.ToHexString(saved.PartBytes(Drawing1))).IsEqualTo(Convert.ToHexString(before));
    }

    [Test]
    public async Task Resizing_a_styled_picture_changes_only_its_size()
    {
        // Combo is rotated, flipped, bordered, shadowed, cropped, grayscale and brightened.
        var (before, after) = SaveEdited(ws => ws.Picture("Combo").WithSize(200, 150));

        await Assert.That(PictureChanges(before, after, "Combo")).IsEqualTo(XmlChangeSet.Expect(
            "modified /xdr:pic[1]/xdr:spPr[1]/a:xfrm[1]/a:ext[1] @cx: '1216251' -> '1905000', @cy: '952500' -> '1428750'"));
    }

    [Test]
    public async Task Renaming_a_picture_changes_only_its_name()
    {
        var (before, after) = SaveEdited(ws => ws.Picture("Combo").Name = "Renamed");

        await Assert.That(XmlChangeSet.Between(before, after).ToString()).IsEqualTo(XmlChangeSet.Expect(
            "modified /xdr:wsDr[1]/xdr:twoCellAnchor[12]/xdr:pic[1]/xdr:nvPicPr[1]/xdr:cNvPr[1] @name: 'Combo' -> 'Renamed'"));
    }

    [Test]
    public async Task Changing_the_placement_of_a_styled_picture_keeps_its_styling()
    {
        // A new placement is a new form of anchor, and the picture moves into it with all it holds.
        var (before, after) = SaveEdited(ws => ws.Picture("Combo").MoveTo(30, 40));

        await Assert.That(PictureChanges(before, after, "Combo")).IsEqualTo(XmlChangeSet.Expect());
        await Assert.That(XDocument.Parse(after).Root!.Elements().Last().Name.LocalName).IsEqualTo("absoluteAnchor");
    }

    [Test]
    public async Task A_second_save_after_an_edit_leaves_the_drawing_part_as_the_first_wrote_it()
    {
        // The first save writes the rename, and the pictures are then as the file holds them.
        using var source = TestHelper.OpenResource(Fixture);
        using var wb = new XLWorkbook(source);
        wb.Worksheet(1).Picture("Combo").Name = "Renamed";
        using var first = new MemoryStream();
        wb.SaveAs(first);
        using var second = new MemoryStream();
        wb.SaveAs(second);

        await Assert.That(Convert.ToHexString(second.PartBytes(Drawing1)))
            .IsEqualTo(Convert.ToHexString(first.PartBytes(Drawing1)));
    }

    [Test]
    public async Task An_image_written_asynchronously_is_saved()
    {
        using var wb = TestHelper.LoadWorkbook(Fixture);
        var picture = wb.Worksheet(1).Picture("Plain");
        var changed = picture.ImageStream.ToArray().Concat(new byte[] { 1, 2, 3 }).ToArray();
        picture.ImageStream.SetLength(0);
        await picture.ImageStream.WriteAsync(changed);

        using var saved = new MemoryStream();
        wb.SaveAs(saved);
        using var reloaded = new XLWorkbook(saved);

        await Assert.That(reloaded.Worksheet(1).Picture("Plain").ImageStream.ToArray().SequenceEqual(changed)).IsTrue();
    }

    /// <summary>The fixture's drawing part, and the one a save after <paramref name="edit"/> writes.</summary>
    internal static (string Before, string After) SaveEdited(Action<IXLWorksheet> edit)
    {
        using var source = TestHelper.OpenResource(Fixture);
        var before = source.ReadPart(Drawing1);
        using var wb = new XLWorkbook(source);
        edit(wb.Worksheet(1));
        using var saved = new MemoryStream();
        wb.SaveAs(saved);
        return (before, saved.ReadPart(Drawing1));
    }

    /// <summary>The change set of one picture's <c>xdr:pic</c>, wherever its anchor put it.</summary>
    internal static string PictureChanges(string before, string after, string name) =>
        XmlChangeSet.Between(PictureXml(before, name), PictureXml(after, name)).ToString();

    private static string PictureXml(string drawing, string name)
    {
        XNamespace xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        return XDocument.Parse(drawing).Descendants(xdr + "pic")
            .Single(p => (string?)p.Descendants(xdr + "cNvPr").Single().Attribute("name") == name)
            .ToString();
    }
}
