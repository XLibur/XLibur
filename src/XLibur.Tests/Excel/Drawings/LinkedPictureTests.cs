using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Drawings;

/// <summary>
/// A new picture beside a linked picture takes a name the drawing does not already use (issue #796).
/// </summary>
/// <remarks>
/// <para>
/// The fixture <c>picture-linked.xlsx</c> was made in Excel 16.0 (build 20430) over COM: a new
/// workbook, values in A1:B2, <c>Range("A1:B2").Copy()</c>, then <c>Pictures().Paste(true)</c> at D4,
/// saved as .xlsx. Excel names the linked picture "Picture 1" (shape type 13) and writes it as an
/// <c>mc:AlternateContent</c> that is a child of <c>xdr:wsDr</c> itself: the
/// <c>mc:Choice Requires="a14"</c> holds a <c>xdr:twoCellAnchor editAs="oneCell"</c> with the
/// <c>xdr:pic</c> (id 1025, an <c>a14:cameraTool cellRange="$A$1:$B$2"</c>), and the
/// <c>mc:Fallback</c> is empty.
/// </para>
/// <para>
/// XLibur does not load the linked picture, and the save keeps it as the file has it. A picture
/// added beside it must still not take its name.
/// </para>
/// <para>
/// Checked in the same Excel over COM: the fixture loaded in XLibur, a picture added at H2 and the
/// workbook saved. Excel opens it with two pictures (shape type 13), "Picture 1" at D4, still linked
/// (<c>Formula</c> is <c>$A$1:$B$2</c>), and "Picture 2" at H2.
/// </para>
/// </remarks>
public class LinkedPictureTests
{
    [Test]
    public async Task A_new_picture_beside_a_linked_picture_takes_the_next_free_name()
    {
        using var wb = new XLWorkbook(Fixture());
        var ws = wb.Worksheet(1);
        await Assert.That(ws.Pictures.Count).IsEqualTo(0);

        var picture = ws.AddPicture(new MemoryStream(Png));

        await Assert.That(picture.Name).IsEqualTo("Picture 2");
    }

    [Test]
    public async Task A_new_picture_cannot_be_given_the_linked_picture_name()
    {
        // The name is taken, as a loaded picture's name is, in any case.
        using var wb = new XLWorkbook(Fixture());
        var ws = wb.Worksheet(1);

        await Assert.That(() => ws.AddPicture(new MemoryStream(Png), "Picture 1")).Throws<ArgumentException>();
        await Assert.That(() => ws.AddPicture(new MemoryStream(Png), "PICTURE 1")).Throws<ArgumentException>();
    }

    [Test]
    public async Task A_failed_add_under_the_linked_picture_name_adds_no_picture()
    {
        // The add throws before the picture is created (issue #806), so the save has only the
        // linked picture.
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(Fixture()))
        {
            var ws = wb.Worksheet(1);

            await Assert.That(() => ws.AddPicture(new MemoryStream(Png), "picture 1")).Throws<ArgumentException>();

            await Assert.That(ws.Pictures.Count).IsEqualTo(0);
            wb.SaveAs(saved);
        }

        await Assert.That(Shapes(saved)).IsEqualTo("AlternateContent pic 'Picture 1' camera $A$1:$B$2");
    }

    [Test]
    public async Task A_picture_cannot_be_renamed_to_the_linked_picture_name()
    {
        using var wb = new XLWorkbook(Fixture());
        var picture = wb.Worksheet(1).AddPicture(new MemoryStream(Png));

        await Assert.That(() => picture.Name = "picture 1").Throws<ArgumentException>();
        await Assert.That(picture.Name).IsEqualTo("Picture 2");
    }

    [Test]
    public async Task A_save_keeps_the_linked_picture_beside_the_new_one_under_its_own_name()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(Fixture()))
        {
            wb.Worksheet(1).AddPicture(new MemoryStream(Png));
            wb.SaveAs(saved);
        }

        await Assert.That(Shapes(saved)).IsEqualTo("AlternateContent pic 'Picture 1' camera $A$1:$B$2, pic 'Picture 2'");
    }

    [Test]
    public async Task A_picture_with_no_name_in_the_file_is_not_given_the_linked_picture_name()
    {
        // Another producer may leave a picture's name empty. The loader names it, and the name must
        // not be one a shape later in the drawing already has.
        using var source = Fixture().RewritePart(Drawing1, xml => xml.Replace(
            "<mc:AlternateContent", UnnamedPicture + "<mc:AlternateContent"));
        using var wb = new XLWorkbook(source);
        var ws = wb.Worksheet(1);

        await Assert.That(ws.Pictures.Single().Name).IsEqualTo("Picture 2");
        await Assert.That(ws.AddPicture(new MemoryStream(Png)).Name).IsEqualTo("Picture 3");
    }

    private const string Drawing1 = "xl/drawings/drawing1.xml";

    private static readonly XNamespace XdrNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private static readonly XNamespace McNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly XNamespace A14Ns = "http://schemas.microsoft.com/office/drawing/2010/main";

    /// <summary>A one-cell anchor at A1 with a picture whose name is empty, showing the fixture's image.</summary>
    private const string UnnamedPicture =
        "<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>" +
        "<xdr:ext cx=\"381000\" cy=\"190500\"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"2\" name=\"\"/><xdr:cNvPicPr/></xdr:nvPicPr>" +
        "<xdr:blipFill><a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rId1\"/>" +
        "<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"381000\" cy=\"190500\"/></a:xfrm>" +
        "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>";

    /// <summary>
    /// Each top-level element of the saved drawing, as the picture it holds and its name, with the
    /// wrapper and the camera tool's range where it has them, e.g.
    /// <c>AlternateContent pic 'Picture 1' camera $A$1:$B$2</c>.
    /// </summary>
    private static string Shapes(Stream saved)
    {
        var shapes = XDocument.Parse(saved.ReadPart(Drawing1)).Root!.Elements().Select(element =>
        {
            var picture = element.Descendants(XdrNs + "pic").Single();
            var name = (string?)picture.Descendants(XdrNs + "cNvPr").Single().Attribute("name");
            var camera = (string?)picture.Descendants(A14Ns + "cameraTool").SingleOrDefault()?.Attribute("cellRange");
            var wrapper = element.Name == McNs + "AlternateContent" ? "AlternateContent " : "";
            return $"{wrapper}pic '{name}'{(camera is null ? "" : $" camera {camera}")}";
        });
        return string.Join(", ", shapes);
    }

    private static readonly byte[] Png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEklEQVR4nGP4z8CAFTEMLQkAI7c/wTpVXRAAAAAASUVORK5CYII=");

    private static MemoryStream Fixture()
    {
        using var resource = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Drawings\picture-linked.xlsx"));
        var package = new MemoryStream();
        resource.CopyTo(package);
        package.Position = 0;
        return package;
    }
}
