using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;
using XLibur.Excel.Drawings;
using XLibur.Tests.Utils;
using static XLibur.Tests.Excel.Drawings.PictureFidelityTests;

namespace XLibur.Tests.Excel.Drawings;

/// <summary>
/// Picture styling through a save: what XLibur reads from Excel's file, what it writes, and that it
/// writes it the way Excel does (spec 17).
/// </summary>
/// <remarks>
/// The fixture (<see cref="PictureFidelityTests.Fixture"/>) was made by Excel over COM: twelve
/// copies of one image, each styled one way.
/// </remarks>
public class PictureStylingRoundTripTests
{
    private static readonly XNamespace Xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static readonly byte[] Png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEklEQVR4nGP4z8CAFTEMLQkAI7c/wTpVXRAAAAAASUVORK5CYII=");

    [Test]
    [Arguments("Plain", "border - - - no=False fill - no=False rot - flip -/- transparency - recolor - lum -/- crop - shadow -")]
    [Arguments("Rotated", "border - - - no=False fill - no=False rot 340 flip True/- transparency - recolor - lum -/- crop - shadow -")]
    [Arguments("FlippedV", "border - - - no=False fill - no=False rot - flip -/True transparency - recolor - lum -/- crop - shadow -")]
    [Arguments("Border", "border FFFF0000 1.5 Dash no=False fill - no=False rot - flip -/- transparency - recolor - lum -/- crop - shadow -")]
    [Arguments("Shadow", "border - - - no=False fill - no=False rot - flip -/- transparency - recolor - lum -/- crop - shadow FF000000 60 4 4.24259842519685 45.00026666666667")]
    [Arguments("Crop", "border - - - no=False fill - no=False rot - flip -/- transparency - recolor - lum -/- crop 5.29,0,0,4.385 shadow -")]
    [Arguments("Lum", "border - - - no=False fill - no=False rot - flip -/- transparency - recolor - lum 20/-10 crop - shadow -")]
    [Arguments("Gray", "border - - - no=False fill - no=False rot - flip -/- transparency - recolor Grayscale lum -/- crop - shadow -")]
    [Arguments("BW", "border - - - no=False fill - no=False rot - flip -/- transparency - recolor BlackAndWhite lum -/- crop - shadow -")]
    [Arguments("Washout", "border - - - no=False fill - no=False rot - flip -/- transparency - recolor Washout lum 70/-70 crop - shadow -")]
    [Arguments("Fill", "border - - - no=False fill FFFFFF00 no=False rot - flip -/- transparency - recolor - lum -/- crop - shadow -")]
    [Arguments("Combo", "border FF0000FF 2 - no=False fill - no=False rot 345 flip True/- transparency - recolor Grayscale lum 20/- crop 4.232,0,0,0 shadow FF000000 0 5 2.8284251968503935 45.000033333333334")]
    public async Task The_styling_Excel_wrote_is_read(string name, string expected)
    {
        using var wb = TestHelper.LoadWorkbook(Fixture);

        await Assert.That(PictureStylingModelTests.Describe(wb.Worksheet(1).Picture(name))).IsEqualTo(expected);
    }

    public static IEnumerable<Func<(string Property, Action<IXLPicture> Assign)>> Assignments()
    {
        yield return () => ("border colour", p => p.BorderColor = XLColor.Red);
        yield return () => ("theme border colour", p => p.BorderColor = XLColor.FromTheme(XLThemeColor.Accent2));
        yield return () => ("border width", p => p.BorderWidthPt = 2.25);
        yield return () => ("border dash", p => p.BorderDash = XLLineDashStyle.LongDashDot);
        yield return () => ("no border", p => p.NoBorder = true);
        yield return () => ("fill", p => p.FillColor = XLColor.Yellow);
        yield return () => ("no fill", p => p.NoFill = true);
        yield return () => ("rotation", p => p.Rotation = 33.5);
        yield return () => ("flips", p => { p.FlipHorizontal = true; p.FlipVertical = true; });
        yield return () => ("transparency", p => p.Transparency = 35);
        yield return () => ("grayscale", p => p.Recolor = XLPictureRecolor.Grayscale);
        yield return () => ("sepia", p => p.Recolor = XLPictureRecolor.Sepia);
        yield return () => ("washout", p => p.Recolor = XLPictureRecolor.Washout);
        yield return () => ("black and white", p => p.Recolor = XLPictureRecolor.BlackAndWhite);
        yield return () => ("brightness and contrast", p => { p.Brightness = -25; p.Contrast = 40; });
        yield return () => ("crop", p => p.Crop = new XLCrop(10, 5, 2.5, 0));
        yield return () => ("shadow", p => p.Shadow = new XLPictureShadow(XLColor.DarkBlue, 40, 6, 3, 135));
        yield return () => ("everything", p =>
        {
            p.BorderColor = XLColor.Green;
            p.BorderWidthPt = 1;
            p.BorderDash = XLLineDashStyle.SystemDot;
            p.FillColor = XLColor.White;
            p.Rotation = 270;
            p.FlipVertical = true;
            p.Transparency = 10;
            p.Recolor = XLPictureRecolor.Sepia;
            p.Brightness = 5;
            p.Contrast = 5;
            p.Crop = new XLCrop(1, 2, 3, 4);
            p.Shadow = new XLPictureShadow(XLColor.Gray, 0, 2, 2, 90);
        });
    }

    [Test]
    [MethodDataSource(nameof(Assignments))]
    public async Task A_styled_new_picture_saves_valid_and_reads_back(string property, Action<IXLPicture> assign)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var picture = ws.AddPicture(new MemoryStream(Png)).MoveTo(ws.Cell("B2"));
        assign(picture);
        var expected = PictureStylingModelTests.Describe(picture);

        using var saved = new MemoryStream();
        wb.SaveAs(saved, validate: true);
        using var reloaded = new XLWorkbook(saved);

        await Assert.That(PictureStylingModelTests.Describe(reloaded.Worksheet(1).Pictures.Single()))
            .IsEqualTo(expected).Because(property);
    }

    [Test]
    [MethodDataSource(nameof(Assignments))]
    public async Task A_styled_loaded_picture_saves_valid_and_reads_back(string property, Action<IXLPicture> assign)
    {
        // Combo already has a border, a shadow, a crop, a recolour and a brightness to be edited.
        using var wb = TestHelper.LoadWorkbook(Fixture);
        var picture = wb.Worksheet(1).Picture("Combo");
        assign(picture);
        var expected = PictureStylingModelTests.Describe(picture);

        using var saved = new MemoryStream();
        wb.SaveAs(saved, validate: true);
        using var reloaded = new XLWorkbook(saved);

        await Assert.That(PictureStylingModelTests.Describe(reloaded.Worksheet(1).Picture("Combo")))
            .IsEqualTo(expected).Because(property);
    }

    [Test]
    [Arguments("Gray", "blipFill")]
    [Arguments("BW", "blipFill")]
    [Arguments("Washout", "blipFill")]
    [Arguments("Lum", "blipFill")]
    [Arguments("Crop", "blipFill")]
    [Arguments("Border", "spPr")]
    [Arguments("Fill", "spPr")]
    [Arguments("FlippedV", "spPr")]
    public async Task Styling_a_plain_picture_as_Excel_did_writes_what_Excel_wrote(string styledAs, string element)
    {
        // The plain picture is given the values read from the one Excel styled, and its element is
        // compared with Excel's. Only the a:xfrm offset may differ: the two pictures are at
        // different places on the sheet.
        var (_, after) = SaveEdited(ws =>
        {
            var styled = ws.Picture(styledAs);
            var plain = ws.Picture("Plain");
            plain.BorderColor = styled.BorderColor;
            plain.BorderWidthPt = styled.BorderWidthPt;
            plain.BorderDash = styled.BorderDash;
            plain.FillColor = styled.FillColor;
            plain.FlipVertical = styled.FlipVertical;
            plain.Recolor = styled.Recolor;
            plain.Brightness = styled.Brightness;
            plain.Contrast = styled.Contrast;
            plain.Crop = styled.Crop;
        });

        var excel = Element(after, styledAs, element);
        var xlibur = Element(after, "Plain", element);
        var changes = XmlChangeSet.Between(excel, xlibur).Describe().Where(c => !c.Contains("/a:off[1]")).ToArray();

        await Assert.That(XmlChangeSet.Expect(changes)).IsEqualTo(XmlChangeSet.Expect());
    }

    [Test]
    public async Task A_recolour_that_is_no_preset_reads_custom_and_survives_an_edit()
    {
        const string duotone = "<a:duotone><a:schemeClr val=\"accent1\"><a:shade val=\"45000\"/></a:schemeClr>" +
                               "<a:schemeClr val=\"accent1\"><a:tint val=\"40000\"/></a:schemeClr></a:duotone>";
        using var source = TestHelper.OpenResource(Fixture).RewritePart(Drawing1, xml =>
        {
            // The first a:grayscl after the picture's name is its own.
            var at = xml.IndexOf("<a:grayscl/>", xml.IndexOf("name=\"Gray\"", StringComparison.Ordinal), StringComparison.Ordinal);
            return xml[..at] + duotone + xml[(at + "<a:grayscl/>".Length)..];
        });
        using var wb = new XLWorkbook(source);
        var picture = wb.Worksheet(1).Picture("Gray");
        await Assert.That(picture.Recolor).IsEqualTo(XLPictureRecolor.Custom);

        picture.Rotation = 10;
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        await Assert.That(Element(saved.ReadPart(Drawing1), "Gray", "blipFill")).Contains("duotone");
    }

    [Test]
    public async Task A_border_width_is_written_into_the_border_that_is_there()
    {
        var (before, after) = SaveEdited(ws => ws.Picture("Border").BorderWidthPt = 3);

        await Assert.That(PictureChanges(before, after, "Border")).IsEqualTo(XmlChangeSet.Expect(
            "modified /xdr:pic[1]/xdr:spPr[1]/a:ln[1] @w: '19050' -> '38100'"));
    }

    [Test]
    public async Task A_fill_colour_replaces_the_fill_that_is_there()
    {
        var (before, after) = SaveEdited(ws => ws.Picture("Fill").FillColor = XLColor.Red);

        await Assert.That(PictureChanges(before, after, "Fill")).IsEqualTo(XmlChangeSet.Expect(
            "modified /xdr:pic[1]/xdr:spPr[1]/a:solidFill[1]/a:srgbClr[1] @val: 'FFFF00' -> 'FF0000'"));
    }

    [Test]
    public async Task A_rotation_is_written_into_the_transform_and_nothing_else()
    {
        var (before, after) = SaveEdited(ws => ws.Picture("Combo").Rotation = 90);

        await Assert.That(PictureChanges(before, after, "Combo")).IsEqualTo(XmlChangeSet.Expect(
            "modified /xdr:pic[1]/xdr:spPr[1]/a:xfrm[1] @rot: '20700000' -> '5400000'"));
    }

    [Test]
    public async Task Removing_the_shadow_and_the_crop_removes_them_and_nothing_else()
    {
        var (before, after) = SaveEdited(ws =>
        {
            var combo = ws.Picture("Combo");
            combo.Shadow = null;
            combo.Crop = null;
        });

        await Assert.That(PictureChanges(before, after, "Combo")).IsEqualTo(XmlChangeSet.Expect(
            "modified /xdr:pic[1]/xdr:blipFill[1]/a:srcRect[1] -@l='4232'",
            "removed /xdr:pic[1]/xdr:spPr[1]/a:effectLst[1] (subtree)"));
    }

    [Test]
    public async Task A_new_recolour_goes_between_the_transparency_and_the_brightness()
    {
        // Excel writes a:alphaModFix, then the recolour, then a:lum.
        var (_, after) = SaveEdited(ws =>
        {
            var lum = ws.Picture("Lum");
            lum.Transparency = 50;
            lum.Recolor = XLPictureRecolor.BlackAndWhite;
        });

        var blip = XElement.Parse(Element(after, "Lum", "blipFill")).Element(A + "blip")!;
        await Assert.That(string.Join(", ", blip.Elements().Select(e => e.Name.LocalName)))
            .IsEqualTo("alphaModFix, biLevel, lum, extLst");
    }

    [Test]
    public async Task A_copy_of_a_loaded_picture_keeps_what_XLibur_does_not_model()
    {
        // A glow is an effect the model does not hold. A copy starts from the loaded xdr:pic.
        const string glow = "<a:effectLst><a:glow rad=\"63500\"><a:schemeClr val=\"accent1\"/></a:glow></a:effectLst>";
        using var source = TestHelper.OpenResource(Fixture)
            .RewritePart(Drawing1, xml => xml.Replace(
                "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor><xdr:twoCellAnchor editAs=\"oneCell\"><xdr:from><xdr:col>3</xdr:col>",
                $"<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>{glow}</xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor><xdr:twoCellAnchor editAs=\"oneCell\"><xdr:from><xdr:col>3</xdr:col>"));
        using var wb = new XLWorkbook(source);
        var ws = wb.Worksheet(1);
        var copy = ws.Picture("Plain").Duplicate();
        copy.Name = "Copy";
        copy.BorderColor = XLColor.Red;

        using var saved = new MemoryStream();
        wb.SaveAs(saved, validate: true);

        var copied = Element(saved.ReadPart(Drawing1), "Copy", "spPr");
        await Assert.That(copied).Contains("glow");
        await Assert.That(copied).Contains("FF0000");
    }

    [Test]
    public async Task A_picture_in_a_group_is_styled_in_place()
    {
        using var wb = TestHelper.LoadWorkbook(@"Other\Drawings\GroupedPictures.xlsx");
        var picture = wb.Worksheet(1).Picture("Picture 1");
        await Assert.That(picture.IsInGroup).IsTrue();
        picture.BorderColor = XLColor.Red;
        picture.Recolor = XLPictureRecolor.Grayscale;

        using var saved = new MemoryStream();
        wb.SaveAs(saved, validate: true);
        using var reloaded = new XLWorkbook(saved);
        var styled = reloaded.Worksheet(1).Picture("Picture 1");

        await Assert.That($"{styled.IsInGroup} {styled.BorderColor} {styled.Recolor}").IsEqualTo("True FFFF0000 Grayscale");
    }

    [Test]
    public async Task A_copy_of_a_picture_whose_shadow_was_removed_has_no_shadow()
    {
        using var wb = TestHelper.LoadWorkbook(Fixture);
        var ws = wb.Worksheet(1);
        var source = ws.Picture("Shadow");
        source.Shadow = null;
        source.Duplicate().Name = "Copy";

        using var saved = new MemoryStream();
        wb.SaveAs(saved, validate: true);

        await Assert.That(Element(saved.ReadPart(Drawing1), "Copy", "spPr")).DoesNotContain("outerShdw");
    }

    [Test]
    public async Task A_copy_made_after_a_save_has_what_the_save_wrote()
    {
        // The save writes the removal; the copy then starts from what the file holds now.
        using var wb = TestHelper.LoadWorkbook(Fixture);
        var ws = wb.Worksheet(1);
        ws.Picture("Shadow").Shadow = null;
        using var first = new MemoryStream();
        wb.SaveAs(first);

        ws.Picture("Shadow").Duplicate().Name = "Copy";
        using var second = new MemoryStream();
        wb.SaveAs(second, validate: true);

        await Assert.That(Element(second.ReadPart(Drawing1), "Copy", "spPr")).DoesNotContain("outerShdw");
    }

    [Test]
    public async Task A_copy_keeps_a_colour_the_model_reads_only_in_part()
    {
        // The model reads a theme colour and not its shade; the copy keeps the colour as written.
        const string shaded = "<a:schemeClr val=\"accent1\"><a:lumMod val=\"75000\"/></a:schemeClr>";
        using var source = TestHelper.OpenResource(Fixture)
            .RewritePart(Drawing1, xml => xml.Replace("<a:srgbClr val=\"FF0000\"/>", shaded));
        using var wb = new XLWorkbook(source);
        wb.Worksheet(1).Picture("Border").Duplicate().Name = "Copy";

        using var saved = new MemoryStream();
        wb.SaveAs(saved, validate: true);

        await Assert.That(Element(saved.ReadPart(Drawing1), "Copy", "spPr")).Contains("lumMod");
    }

    /// <summary>One child of a picture's <c>xdr:pic</c>, by the picture's name and the child's local name.</summary>
    private static string Element(string drawing, string picture, string localName) =>
        XDocument.Parse(drawing).Descendants(Xdr + "pic")
            .Single(p => (string?)p.Descendants(Xdr + "cNvPr").Single().Attribute("name") == picture)
            .Element(Xdr + localName)!
            .ToString();
}
