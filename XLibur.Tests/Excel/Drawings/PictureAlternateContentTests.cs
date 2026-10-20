using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;
using XLibur.Tests.Utils;
using LoadOptions = XLibur.Excel.LoadOptions;

namespace XLibur.Tests.Excel.Drawings;

/// <summary>
/// A picture whose <c>xdr:pic</c> sits inside <c>mc:AlternateContent</c> keeps the wrapper and both
/// of its branches through a load and save (issue #787).
/// </summary>
/// <remarks>
/// <para>
/// The fixture <c>picture-3d-model.xlsx</c> was made in Excel 16.0 (build 20430) over COM: a new
/// workbook, <c>Shapes.Add3DModel</c> of an ASCII STL tetrahedron at left 100 pt, top 50 pt, saved
/// as .xlsx. Excel writes the model as one <c>xdr:twoCellAnchor</c> (C4 to F14, no <c>editAs</c>)
/// holding <c>mc:AlternateContent</c>: the <c>mc:Choice Requires="am3d"</c> holds the model's
/// <c>xdr:graphicFrame</c>, and the <c>mc:Fallback</c> holds an <c>xdr:pic</c> of its PNG render,
/// with the same drawing id (2) and name ("3D Model 1") as the frame. The render's image part is
/// also the model's <c>am3d:raster</c>.
/// </para>
/// <para>
/// XLibur loads the fallback <c>xdr:pic</c> as an ordinary picture, so the model can be moved,
/// resized, renamed or deleted through it. The anchor is the model's, so a move or resize moves the
/// model; a rename names both branches, which Excel gives one name.
/// </para>
/// </remarks>
public class PictureAlternateContentTests
{
    private const string Drawing1 = "xl/drawings/drawing1.xml";

    private static readonly XNamespace XdrNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private static readonly XNamespace McNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace Am3dNs = "http://schemas.microsoft.com/office/drawing/2017/model3d";
    private static readonly XNamespace RNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    [Test]
    public async Task The_3d_model_loads_as_its_fallback_picture()
    {
        using var wb = LoadFixture();

        var picture = wb.Worksheet(1).Pictures.Single();
        await Assert.That(picture.Name).IsEqualTo("3D Model 1");
        await Assert.That(picture.Id).IsEqualTo(2);
    }

    [Test]
    public async Task A_save_keeps_the_3d_model()
    {
        // The save renumbers the drawing ids from 1. The two branches are one drawing, so they keep
        // one id between them.
        using var saved = LoadAndSave(_ => { });

        await Assert.That(Shape(saved)).IsEqualTo(
            "twoCellAnchor > AlternateContent > Choice(am3d) graphicFrame 1 '3D Model 1' | Fallback pic 1 '3D Model 1'");
    }

    [Test]
    public async Task Moving_the_picture_moves_the_3d_model()
    {
        // MoveTo(cell) makes a one-cell anchor, so the anchor changes its form and keeps what it holds.
        using var saved = LoadAndSave(ws => ws.Picture("3D Model 1").MoveTo(ws.Cell("H2")));

        await Assert.That(Shape(saved)).IsEqualTo(
            "oneCellAnchor > AlternateContent > Choice(am3d) graphicFrame 1 '3D Model 1' | Fallback pic 1 '3D Model 1'");
        await Assert.That(Markers(saved)).IsEqualTo("from 7,1");
    }

    [Test]
    public async Task Resizing_the_picture_resizes_the_3d_model()
    {
        using var saved = LoadAndSave(ws =>
        {
            var picture = ws.Picture("3D Model 1");
            picture.MoveTo(ws.Cell("H2"));
            picture.Width = 40;
            picture.Height = 20;
        });

        await Assert.That(Shape(saved)).IsEqualTo(
            "oneCellAnchor > AlternateContent > Choice(am3d) graphicFrame 1 '3D Model 1' | Fallback pic 1 '3D Model 1'");
        await Assert.That(Extent(saved)).IsEqualTo($"{40 * 9525}x{20 * 9525}");
    }

    [Test]
    public async Task Moving_the_3d_model_to_a_cell_keeps_the_size_excel_shows_it_at()
    {
        // Excel, at 120 DPI, shows the model 147.89 pt square over columns of 48 pt (80 px, a digit
        // width of 9 px for Aptos Narrow 11) and rows of 14.4 pt (24 px). Loaded at that DPI, its box
        // is 246 px each way, and the one-cell anchor written for the move keeps it (#793). 7620 EMU
        // is a pixel at 120 DPI.
        using var saved = LoadAndSave(Fixture(), AtExcelDpi, ws => ws.Picture("3D Model 1").MoveTo(ws.Cell("H2")));

        await Assert.That(Extent(saved)).IsEqualTo($"{246 * 7620}x{246 * 7620}");
    }

    [Test]
    public async Task Moving_the_3d_model_to_a_position_keeps_the_size_excel_shows_it_at()
    {
        using var saved = LoadAndSave(Fixture(), AtExcelDpi, ws => ws.Picture("3D Model 1").MoveTo(30, 40));

        await Assert.That(Extent(saved)).IsEqualTo($"{246 * 7620}x{246 * 7620}");
    }

    [Test]
    public async Task The_3d_model_measures_its_columns_at_96_dpi()
    {
        // At 96 DPI Aptos Narrow 11 has a digit width of 7 px, so a column is 64 px and the box is
        // the model's 1878203 EMU, 197 px wide. Only the width: the rows are at a default height the
        // file does not state, which Excel derives from the font at its own DPI.
        using var wb = LoadFixture();

        await Assert.That(wb.Worksheet(1).Picture("3D Model 1").Width).IsEqualTo(197);
    }

    [Test]
    public async Task Moving_the_3d_model_to_a_cell_keeps_its_width_at_96_dpi()
    {
        // The 197 px the box measures at 96 DPI (above), by the rule: Excel on this machine runs at
        // 120 DPI and cannot be measured at 96. The height is left unpinned, as above.
        using var saved = LoadAndSave(ws => ws.Picture("3D Model 1").MoveTo(ws.Cell("H2")));

        await Assert.That(Extent(saved)).StartsWith($"{197 * 9525}x");
    }

    [Test]
    public async Task Moving_the_picture_between_cells_keeps_its_two_cell_anchor_in_place()
    {
        // MoveTo(from, to) keeps the two-cell form, so the anchor is the one Excel wrote, edited.
        using var saved = LoadAndSave(ws => ws.Picture("3D Model 1").MoveTo(ws.Cell("H2"), ws.Cell("J9")));

        await Assert.That(Shape(saved)).IsEqualTo(
            "twoCellAnchor > AlternateContent > Choice(am3d) graphicFrame 1 '3D Model 1' | Fallback pic 1 '3D Model 1'");
        await Assert.That(Markers(saved)).IsEqualTo("from 7,1 to 9,8");
    }

    [Test]
    public async Task Renaming_the_picture_renames_both_forms_of_the_3d_model()
    {
        using var saved = LoadAndSave(ws => ws.Picture("3D Model 1").Name = "Tetrahedron");

        await Assert.That(Shape(saved)).IsEqualTo(
            "twoCellAnchor > AlternateContent > Choice(am3d) graphicFrame 1 'Tetrahedron' | Fallback pic 1 'Tetrahedron'");
    }

    [Test]
    public async Task The_3d_model_name_is_free_again_once_the_picture_is_renamed()
    {
        // The frame and the fallback picture are one drawing with one name, which the picture holds:
        // no other shape keeps it once the picture lets it go (#796).
        using var wb = LoadFixture();
        var ws = wb.Worksheet(1);
        ws.Picture("3D Model 1").Name = "Model";

        var other = ws.AddPicture(new MemoryStream(Png), "3D Model 1");

        await Assert.That(other.Name).IsEqualTo("3D Model 1");
    }

    [Test]
    public async Task A_second_save_keeps_the_3d_model_and_finds_it_again()
    {
        // The second save edits what the first wrote, so it must find the picture by the id the first
        // gave it; a rename between the saves shows it did.
        using var wb = LoadFixture();
        using var first = new MemoryStream();
        using var second = new MemoryStream();
        wb.SaveAs(first);
        wb.Worksheet(1).Picture("3D Model 1").Name = "Tetrahedron";
        wb.SaveAs(second);

        await Assert.That(Shape(second)).IsEqualTo(
            "twoCellAnchor > AlternateContent > Choice(am3d) graphicFrame 1 'Tetrahedron' | Fallback pic 1 'Tetrahedron'");
    }

    [Test]
    public async Task A_one_cell_3d_model_keeps_its_two_cell_anchor_when_rows_are_inserted_above_it()
    {
        // Excel writes editAs="oneCell" for a model set to "Move but don't size with cells". The
        // fixture is rewritten to say so, and the kept anchor is edited as a plain picture's is (#780).
        using var source = Fixture().RewritePart(Drawing1, xml => xml.Replace("<xdr:twoCellAnchor>", "<xdr:twoCellAnchor editAs=\"oneCell\">"));
        using var wb = new XLWorkbook(source);
        wb.Worksheet(1).Row(1).InsertRowsAbove(2);
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        await Assert.That(saved.ReadPart(Drawing1)).Contains("<xdr:twoCellAnchor editAs=\"oneCell\">");
        await Assert.That(Shape(saved)).IsEqualTo(
            "twoCellAnchor > AlternateContent > Choice(am3d) graphicFrame 1 '3D Model 1' | Fallback pic 1 '3D Model 1'");
        await Assert.That(Markers(saved)).IsEqualTo("from 2,5 to 5,15");
    }

    [Test]
    public async Task Deleting_the_picture_deletes_the_3d_model()
    {
        // The model part and its render go with the anchor, and so does the drawing, now empty.
        using var saved = LoadAndSave(ws => ws.Pictures.Delete("3D Model 1"));

        await Assert.That(saved.PartsUnder("xl/drawings/")).IsEmpty();
        await Assert.That(saved.PartsUnder("xl/media/")).IsEmpty();
    }

    [Test]
    public async Task Deleting_one_copy_of_a_3d_model_keeps_the_model_the_other_copy_shows()
    {
        // Excel's copies of a 3D model share its model part and its render.
        using var saved = LoadAndSave(Duplicated(), ws => ws.Pictures.Delete("3D Model 2"));

        await Assert.That(Shape(saved)).IsEqualTo(
            "twoCellAnchor > AlternateContent > Choice(am3d) graphicFrame 1 '3D Model 1' | Fallback pic 1 '3D Model 1'");
        await Assert.That(Relationships(saved)).IsEqualTo("image ../media/image1.png, model3d ../media/model3d1.glb");
        await Assert.That(string.Join(", ", saved.PartsUnder("xl/media/").Order()))
            .IsEqualTo("xl/media/image1.png, xl/media/model3d1.glb");
    }

    [Test]
    public async Task Deleting_every_copy_of_a_3d_model_deletes_the_model()
    {
        using var saved = LoadAndSave(Duplicated(), ws =>
        {
            ws.Pictures.Delete("3D Model 1");
            ws.Pictures.Delete("3D Model 2");
        });

        await Assert.That(saved.PartsUnder("xl/drawings/")).IsEmpty();
        await Assert.That(saved.PartsUnder("xl/media/")).IsEmpty();
    }

    [Test]
    public async Task Deleting_a_3d_model_keeps_the_pictures_beside_it()
    {
        // A picture added beside the model keeps the drawing, and its own image, after the model goes.
        using var saved = LoadAndSave(ws =>
        {
            ws.Pictures.Delete("3D Model 1");
            ws.AddPicture(new MemoryStream(Png), "Logo").MoveTo(ws.Cell("B2"));
        });

        await Assert.That(Shape(saved)).IsEqualTo("oneCellAnchor > pic 1 'Logo'");
        await Assert.That(Relationships(saved)).StartsWith("image ");
        await Assert.That(Relationships(saved)).DoesNotContain("model3d");
        await Assert.That(saved.PartsUnder("xl/media/").Length).IsEqualTo(1);
    }

    [Test]
    public async Task Changing_the_image_of_one_copy_of_a_3d_model_moves_its_render_with_it()
    {
        // The changed copy takes an image part of its own (#781). The model's am3d:raster is the same
        // render as the fallback picture, so it moves to the new part too, and the other copy keeps both
        // on the part they shared.
        using var saved = LoadAndSave(Duplicated(), ws =>
        {
            var picture = ws.Picture("3D Model 2");
            var changed = picture.ImageStream.ToArray().Concat(new byte[] { 1, 2, 3 }).ToArray();
            picture.ImageStream.SetLength(0);
            picture.ImageStream.Write(changed);
        });

        await Assert.That(Renders(saved)).IsEqualTo(
            "3D Model 1: raster image1.png, fallback image1.png | 3D Model 2: raster image2.png, fallback image2.png");
    }

    [Test]
    public async Task A_3d_model_grouped_with_a_picture_keeps_the_model_and_loads_the_picture()
    {
        // Excel puts the wrapper inside the group: anchor > grpSp > [mc:AlternateContent, pic]. The
        // picture loads as a grouped picture; the model inside the group is left as Excel wrote it,
        // and a rename of the picture reaches the picture alone (#795).
        using var source = Resource("picture-3d-model-grouped.xlsx");
        using var wb = new XLWorkbook(source);
        await Assert.That(string.Join(", ", wb.Worksheet(1).Pictures.Select(p => p.Name))).IsEqualTo("Picture 3");
        wb.Worksheet(1).Picture("Picture 3").Name = "Photo";

        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        await Assert.That(Shape(saved)).IsEqualTo("twoCellAnchor > grpSp 5 'Group 4'");
        await Assert.That(GroupMembers(saved)).IsEqualTo(
            "AlternateContent > Choice(am3d) graphicFrame 2 '3D Model 1' | Fallback pic 2 '3D Model 1', pic 4 'Photo'");
        await Assert.That(Relationships(saved)).Contains("model3d ../media/model3d1.glb");
    }

    [Test]
    public async Task Grouping_the_3d_model_with_a_picture_keeps_the_model_in_the_group()
    {
        // As Excel does (picture-3d-model-grouped.xlsx), the whole wrapper goes into the group, with
        // the model in its mc:Choice, and not only the fallback picture (#794). Excel 16.0 (build
        // 20430) opened the saved file with the group's first item a 3D model (Shapes type 30,
        // msoModel3D, with a Model3D), at 75 pt, 37.5 pt and 150 by 112.5 pt: the box set below.
        using var saved = GroupedWithLogo();

        // The first save numbered the model 1 and the picture 2, so the group takes 3.
        await Assert.That(Shape(saved)).IsEqualTo("absoluteAnchor > grpSp 3 'Group 3'");
        await Assert.That(GroupMembers(saved)).IsEqualTo(
            "AlternateContent > Choice(am3d) graphicFrame 1 '3D Model 1' | Fallback pic 1 '3D Model 1', pic 2 'Logo'");
        await Assert.That(Relationships(saved)).Contains("model3d ../media/model3d1.glb");
    }

    [Test]
    public async Task A_grouped_3d_model_is_placed_in_the_group_where_its_picture_was()
    {
        // In a group each drawing is placed by its own transform, in the group's child coordinates,
        // which XLibur makes the sheet's. Excel places the model's frame there too, at the box its
        // fallback picture has: off 1270000,635000 ext 1878203 square in picture-3d-model-grouped.xlsx.
        // A frame in an anchor of its own is placed by the anchor, and Excel writes its transform as
        // zeros, so the group must set it, or the model shrinks to nothing at the group's corner.
        using var saved = GroupedWithLogo();

        await Assert.That(GroupedTransforms(saved)).IsEqualTo(
            $"graphicFrame {Box(100, 50, 200, 150)} | pic {Box(100, 50, 200, 150)}");
    }

    [Test]
    public async Task Moving_and_resizing_a_grouped_3d_model_moves_both_forms()
    {
        // Once grouped and saved, the model is written in place in its group. Both forms of it are
        // one drawing, so the model's frame moves and resizes with its fallback picture (#794).
        using var saved = GroupedWithLogo(ws =>
        {
            var model = ws.Picture("3D Model 1").MoveTo(120, 60);
            model.Width = 240;
        });

        await Assert.That(GroupedTransforms(saved)).IsEqualTo(
            $"graphicFrame {Box(120, 60, 240, 150)} | pic {Box(120, 60, 240, 150)}");
    }

    [Test]
    public async Task Deleting_a_grouped_3d_model_deletes_the_model()
    {
        // The model was grouped in this session, so its picture is still loaded and can be deleted.
        // The wrapper goes with it, and the model part, leaving the group its other picture.
        using var saved = GroupedWithLogo(ws => ws.Pictures.Delete("3D Model 1"));

        await Assert.That(GroupMembers(saved)).IsEqualTo("pic 2 'Logo'");
        await Assert.That(Relationships(saved)).DoesNotContain("model3d");
        await Assert.That(saved.PartsUnder("xl/media/").Length).IsEqualTo(1);
    }

    /// <summary>
    /// The 3D model moved to (100, 50) px, sized 200 by 150 px, and grouped with a picture beside it. Grouping needs both
    /// free-floating and saved, so they are saved once before the group is made. An edit after the
    /// group is made is saved once more.
    /// </summary>
    private static MemoryStream GroupedWithLogo(Action<IXLWorksheet>? afterGrouping = null)
    {
        var saved = new MemoryStream();
        using var source = Fixture();
        using var wb = new XLWorkbook(source);
        var ws = wb.Worksheet(1);
        var model = ws.Picture("3D Model 1").MoveTo(100, 50);
        model.Width = 200;
        model.Height = 150;
        var logo = ws.AddPicture(new MemoryStream(Png), "Logo").MoveTo(400, 50);
        using var first = new MemoryStream();
        wb.SaveAs(first);

        ws.Pictures.Group(model, logo);
        using var grouped = new MemoryStream();
        wb.SaveAs(afterGrouping is null ? saved : grouped);
        if (afterGrouping is not null)
        {
            afterGrouping(ws);
            wb.SaveAs(saved);
        }

        return saved;
    }

    /// <summary>A box of pixels at 96 DPI as EMU, e.g. <c>off 952500,476250 ext 1905000x1428750</c>.</summary>
    private static string Box(int left, int top, int width, int height) =>
        $"off {left * 9525},{top * 9525} ext {width * 9525}x{height * 9525}";

    /// <summary>
    /// The transform of each drawing in the branches of the wrapper in the saved drawing's group,
    /// described as <see cref="Box"/> describes one.
    /// </summary>
    private static string GroupedTransforms(Stream saved)
    {
        var wrapper = XDocument.Parse(saved.ReadPart(Drawing1)).Descendants(McNs + "AlternateContent").Single();
        return string.Join(" | ", wrapper.Elements().Select(branch =>
        {
            var drawing = branch.Elements().Single();
            var off = drawing.Descendants(ANs + "off").First();
            var ext = drawing.Descendants(ANs + "ext").First(e => e.Attribute("cx") is not null);
            return $"{drawing.Name.LocalName} off {(string?)off.Attribute("x")},{(string?)off.Attribute("y")} " +
                   $"ext {(string?)ext.Attribute("cx")}x{(string?)ext.Attribute("cy")}";
        }));
    }

    [Test]
    public async Task A_group_inside_the_wrapper_is_left_as_the_file_has_it()
    {
        // Excel was not seen writing a group inside the wrapper (it puts the wrapper inside the group,
        // as above), but another producer may. A picture in it is not the anchor's own, so it is not
        // loaded, as a 3D model inside a group is not: read as the anchor's picture, it would be
        // sized by its own extent in the group, and the save would shrink the group's anchor to it.
        using var source = Fixture().RewritePart(Drawing1, GroupInsideWrapper);
        using var wb = new XLWorkbook(source);
        await Assert.That(wb.Worksheet(1).Pictures.Count).IsEqualTo(0);

        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        await Assert.That(Shape(saved)).IsEqualTo(
            "oneCellAnchor > AlternateContent > Choice(am3d) graphicFrame 2 '3D Model 1' | Fallback grpSp 3 'Group 3'");
        await Assert.That(Extent(saved)).IsEqualTo("3048000x1878203");
        await Assert.That(Relationships(saved)).IsEqualTo("image ../media/image1.png, model3d ../media/model3d1.glb");
    }

    /// <summary>
    /// Rewrites the fixture's anchor as a one-cell anchor the size of a group, and puts the fallback
    /// picture in that group, beside a rectangle.
    /// </summary>
    private static string GroupInsideWrapper(string xml)
    {
        const string group =
            "<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=\"3\" name=\"Group 3\"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>" +
            "<xdr:grpSpPr><a:xfrm><a:off x=\"1270000\" y=\"635000\"/><a:ext cx=\"3048000\" cy=\"1878203\"/>" +
            "<a:chOff x=\"1270000\" y=\"635000\"/><a:chExt cx=\"3048000\" cy=\"1878203\"/></a:xfrm></xdr:grpSpPr>";
        const string rectangle =
            "<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"4\" name=\"Rectangle 4\"/><xdr:cNvSpPr/></xdr:nvSpPr>" +
            "<xdr:spPr><a:xfrm><a:off x=\"3810000\" y=\"635000\"/><a:ext cx=\"508000\" cy=\"508000\"/></a:xfrm>" +
            "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp>";

        xml = Regex.Replace(xml, "<xdr:twoCellAnchor>(<xdr:from>.*?</xdr:from>)<xdr:to>.*?</xdr:to>",
            "<xdr:oneCellAnchor>$1<xdr:ext cx=\"3048000\" cy=\"1878203\"/>", RegexOptions.None, TimeSpan.FromSeconds(1));
        xml = xml.Replace("</xdr:twoCellAnchor>", "</xdr:oneCellAnchor>");
        return xml.Replace("<mc:Fallback><xdr:pic>", "<mc:Fallback>" + group + "<xdr:pic>")
            .Replace("</xdr:pic></mc:Fallback>", "</xdr:pic>" + rectangle + "</xdr:grpSp></mc:Fallback>");
    }

    private static readonly byte[] Png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEklEQVR4nGP4z8CAFTEMLQkAI7c/wTpVXRAAAAAASUVORK5CYII=");

    private static XLWorkbook LoadFixture() => new(Fixture());

    private static MemoryStream Fixture() => Resource("picture-3d-model.xlsx");

    /// <summary>
    /// A 3D model and its copy, made in Excel 16.0 (build 20430) over COM: <c>Shapes.Add3DModel</c> of
    /// the same tetrahedron, then <c>Duplicate</c>. Excel writes two anchors like the one of
    /// <c>picture-3d-model.xlsx</c>, "3D Model 1" (id 2) and "3D Model 2" (id 3), and both embed the
    /// same model part (rId1) and the same render (rId2).
    /// </summary>
    private static MemoryStream Duplicated() => Resource("picture-3d-model-duplicated.xlsx");

    private static MemoryStream Resource(string name)
    {
        using var resource = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath($@"Other\Drawings\{name}"));
        var package = new MemoryStream();
        resource.CopyTo(package);
        package.Position = 0;
        return package;
    }

    private static readonly LoadOptions AtExcelDpi = new() { Dpi = new System.Drawing.Point(120, 120) };

    private static MemoryStream LoadAndSave(Action<IXLWorksheet> edit) => LoadAndSave(Fixture(), edit);

    private static MemoryStream LoadAndSave(MemoryStream source, Action<IXLWorksheet> edit) =>
        LoadAndSave(source, new LoadOptions(), edit);

    private static MemoryStream LoadAndSave(MemoryStream source, LoadOptions options, Action<IXLWorksheet> edit)
    {
        var saved = new MemoryStream();
        using (source)
        using (var wb = new XLWorkbook(source, options))
        {
            edit(wb.Worksheet(1));
            wb.SaveAs(saved);
        }

        return saved;
    }

    /// <summary>
    /// The drawings in the group of the saved drawing's first anchor, described as <see cref="Shape"/>
    /// describes an anchor's content.
    /// </summary>
    private static string GroupMembers(Stream saved)
    {
        var group = XDocument.Parse(saved.ReadPart(Drawing1)).Root!.Elements().First().Element(XdrNs + "grpSp")!;
        return string.Join(", ", group.Elements()
            .Where(e => e.Name.LocalName is not ("nvGrpSpPr" or "grpSpPr"))
            .Select(Describe));
    }

    /// <summary>
    /// For each anchor of the saved drawing, the image file its model's <c>am3d:raster</c> and its
    /// fallback picture show, e.g. <c>3D Model 1: raster image1.png, fallback image1.png</c>.
    /// </summary>
    private static string Renders(Stream saved)
    {
        var targets = XDocument.Parse(saved.ReadPart("xl/drawings/_rels/drawing1.xml.rels")).Root!.Elements()
            .ToDictionary(r => (string)r.Attribute("Id")!, r => ((string)r.Attribute("Target")!).Split('/')[^1]);
        var anchors = XDocument.Parse(saved.ReadPart(Drawing1)).Root!.Elements().Select(anchor =>
        {
            var name = (string?)anchor.Descendants(XdrNs + "cNvPr").First().Attribute("name");
            var raster = anchor.Descendants(Am3dNs + "blip").Single().Attribute(RNs + "embed")!.Value;
            var fallback = anchor.Descendants(ANs + "blip").Single().Attribute(RNs + "embed")!.Value;
            return $"{name}: raster {targets[raster]}, fallback {targets[fallback]}";
        });
        return string.Join(" | ", anchors);
    }

    /// <summary>
    /// The saved drawing's relationships, each as its type's last segment and its target, in target
    /// order, e.g. <c>image ../media/image1.png, model3d ../media/model3d1.glb</c>.
    /// </summary>
    private static string Relationships(Stream saved)
    {
        var relationships = XDocument.Parse(saved.ReadPart("xl/drawings/_rels/drawing1.xml.rels")).Root!.Elements()
            .Select(r => $"{((string)r.Attribute("Type")!).Split('/')[^1]} {(string?)r.Attribute("Target")}")
            .Order();
        return string.Join(", ", relationships);
    }

    /// <summary>
    /// The zero-based column and row of each marker of the saved drawing's first anchor, e.g.
    /// <c>from 7,1</c>.
    /// </summary>
    private static string Markers(Stream saved)
    {
        var anchor = XDocument.Parse(saved.ReadPart(Drawing1)).Root!.Elements().First();
        var markers = anchor.Elements()
            .Where(e => e.Name.LocalName is "from" or "to")
            .Select(e => $"{e.Name.LocalName} {(string?)e.Element(XdrNs + "col")},{(string?)e.Element(XdrNs + "row")}");
        return string.Join(" ", markers);
    }

    /// <summary>The extent of the saved drawing's first anchor in EMU, e.g. <c>381000x190500</c>.</summary>
    private static string Extent(Stream saved)
    {
        var ext = XDocument.Parse(saved.ReadPart(Drawing1)).Root!.Elements().First().Element(XdrNs + "ext");
        return ext is null ? "none" : $"{(string?)ext.Attribute("cx")}x{(string?)ext.Attribute("cy")}";
    }

    /// <summary>
    /// The saved drawing's only anchor, the element that holds its drawing, and each branch of that
    /// element with the drawing it holds, its id and its name.
    /// </summary>
    private static string Shape(Stream saved)
    {
        var anchors = XDocument.Parse(saved.ReadPart(Drawing1)).Root!.Elements().ToList();
        if (anchors.Count != 1)
            return $"{anchors.Count} anchors";

        var anchor = anchors[0];
        var content = anchor.Elements().Single(e => e.Name.LocalName is not ("from" or "to" or "pos" or "ext" or "clientData"));
        return $"{anchor.Name.LocalName} > {Describe(content)}";
    }

    /// <summary>
    /// A drawing as its element, id and name, e.g. <c>pic 1 '3D Model 1'</c>, or a wrapper as each of
    /// its branches with the drawing it holds.
    /// </summary>
    private static string Describe(XElement content)
    {
        if (content.Name != McNs + "AlternateContent")
            return Drawing(content);

        var branches = content.Elements().Select(branch =>
        {
            var requires = (string?)branch.Attribute("Requires");
            var label = requires is null ? branch.Name.LocalName : $"{branch.Name.LocalName}({requires})";
            var drawing = branch.Elements().SingleOrDefault();
            return drawing is null ? label : $"{label} {Drawing(drawing)}";
        });
        return $"AlternateContent > {string.Join(" | ", branches)}";

        static string Drawing(XElement drawing)
        {
            var properties = drawing.Descendants(XdrNs + "cNvPr").First();
            return $"{drawing.Name.LocalName} {(string?)properties.Attribute("id")} '{(string?)properties.Attribute("name")}'";
        }
    }
}
