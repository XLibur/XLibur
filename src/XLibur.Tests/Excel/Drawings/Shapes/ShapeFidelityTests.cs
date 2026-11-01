using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;
using XLibur.Excel.Drawings;
using XLibur.Tests.Utils;
using static XLibur.Tests.Excel.Drawings.Shapes.ShapeCreationTests;

namespace XLibur.Tests.Excel.Drawings.Shapes;

/// <summary>
/// A loaded shape is written back by patching the <c>xdr:sp</c> it was read from, so an edit changes
/// exactly what it names and nothing beside it (spec 15, acceptance criteria 3, 6 and 7).
/// </summary>
/// <remarks>
/// The fixture was made by Excel over COM: a plain text box and rectangle, a text box with a
/// gradient, a shadow, a rotation and two paragraphs of mixed runs ("Styled"), an arc with
/// arrowheads, two rectangles joined by a connector, and an empty text box.
/// </remarks>
public class ShapeFidelityTests
{
    [Test]
    public async Task A_save_that_edits_no_shape_leaves_the_drawing_part_as_it_was()
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
    public async Task Editing_one_run_rewrites_only_its_text()
    {
        var (before, after) = SaveEdited(ws => ws.Shapes.Shape("Styled").TextBody[0][1].Text = " changed");

        await Assert.That(XmlChangeSet.Between(before, after).ToString()).IsEqualTo(XmlChangeSet.Expect(
            "modified /xdr:wsDr[1]/xdr:twoCellAnchor[3]/xdr:sp[1]/xdr:txBody[1]/a:p[1]/a:r[2]/a:t[1] text: ' plain' -> ' changed'"));
    }

    [Test]
    public async Task Formatting_one_run_changes_only_its_run_properties()
    {
        var (before, after) = SaveEdited(ws =>
        {
            var run = ws.Shapes.Shape("Styled").TextBody[0][0];
            run.Italic = true;
            run.FontName = "Arial";
        });

        await Assert.That(ShapeChanges(before, after, "Styled")).IsEqualTo(XmlChangeSet.Expect(
            "modified /xdr:sp[1]/xdr:txBody[1]/a:p[1]/a:r[1]/a:rPr[1] +@i='1'",
            "added /xdr:sp[1]/xdr:txBody[1]/a:p[1]/a:r[1]/a:rPr[1]/a:latin[1]"));
    }

    [Test]
    public async Task Changing_the_text_leaves_the_gradient_shadow_and_rotation_alone()
    {
        var (before, after) = SaveEdited(ws => ((IXLTextBox)ws.Shapes.Shape("Styled")).Text = "Replaced");

        await Assert.That(XmlChangeSet.Between(SpPr(before, "Styled"), SpPr(after, "Styled")).ToString())
            .IsEqualTo(XmlChangeSet.Expect());

        // The new run takes the formatting of the first run the body had.
        await Assert.That(ShapeChanges(before, after, "Styled")).IsEqualTo(XmlChangeSet.Expect(
            "modified /xdr:sp[1]/xdr:txBody[1]/a:p[1]/a:r[1]/a:t[1] text: 'Bold' -> 'Replaced'",
            "removed /xdr:sp[1]/xdr:txBody[1]/a:p[1]/a:r[2] (subtree)",
            "removed /xdr:sp[1]/xdr:txBody[1]/a:p[2] (subtree)"));
    }

    [Test]
    public async Task A_fill_colour_replaces_a_gradient_and_keeps_the_line_and_shadow()
    {
        var (before, after) = SaveEdited(ws => ws.Shapes.Shape("Styled").FillColor = XLColor.FromHtml("#00FF00"));

        await Assert.That(ShapeChanges(before, after, "Styled")).IsEqualTo(XmlChangeSet.Expect(
            "added /xdr:sp[1]/xdr:spPr[1]/a:solidFill[1] (subtree)",
            "removed /xdr:sp[1]/xdr:spPr[1]/a:gradFill[1] (subtree)"));

        var children = XElement.Parse(SpPr(after, "Styled")).Elements().Select(e => e.Name.LocalName);
        await Assert.That(string.Join(",", children)).IsEqualTo("xfrm,prstGeom,solidFill,ln,effectLst");
    }

    [Test]
    public async Task A_line_width_keeps_the_arrowheads()
    {
        var (before, after) = SaveEdited(ws => ws.Shapes.Shape("Arc").LineWidthPt = 3);

        await Assert.That(ShapeChanges(before, after, "Arc")).IsEqualTo(XmlChangeSet.Expect(
            "modified /xdr:sp[1]/xdr:spPr[1]/a:ln[1] @w: '25400' -> '38100'"));
    }

    [Test]
    public async Task Renaming_a_shape_changes_only_its_name()
    {
        var (before, after) = SaveEdited(ws => ws.Shapes.Shape("Arc").Name = "Bow");

        await Assert.That(XmlChangeSet.Between(before, after).ToString()).IsEqualTo(XmlChangeSet.Expect(
            "modified /xdr:wsDr[1]/xdr:twoCellAnchor[4]/xdr:sp[1]/xdr:nvSpPr[1]/xdr:cNvPr[1] @name: 'Arc' -> 'Bow'"));
    }

    [Test]
    public async Task Moving_a_shape_writes_only_its_markers()
    {
        var (before, after) = SaveEdited(ws => ws.Shapes.Shape("Left").MoveTo(ws.Cell("B20"), ws.Cell("C22")));

        await Assert.That(XmlChangeSet.Between(before, after).ToString()).IsEqualTo(XmlChangeSet.Expect(
            "modified /xdr:wsDr[1]/xdr:twoCellAnchor[5]/xdr:from[1]/xdr:colOff[1] text: '152400' -> '0'",
            "modified /xdr:wsDr[1]/xdr:twoCellAnchor[5]/xdr:from[1]/xdr:row[1] text: '15' -> '19'",
            "modified /xdr:wsDr[1]/xdr:twoCellAnchor[5]/xdr:from[1]/xdr:rowOff[1] text: '50800' -> '0'",
            "modified /xdr:wsDr[1]/xdr:twoCellAnchor[5]/xdr:to[1]/xdr:colOff[1] text: '304800' -> '0'",
            "modified /xdr:wsDr[1]/xdr:twoCellAnchor[5]/xdr:to[1]/xdr:row[1] text: '18' -> '21'",
            "modified /xdr:wsDr[1]/xdr:twoCellAnchor[5]/xdr:to[1]/xdr:rowOff[1] text: '10160' -> '0'"));
    }

    [Test]
    public async Task Changing_the_placement_moves_the_shape_into_a_new_anchor_intact()
    {
        // As for a picture: a new placement is a new form of anchor, and the shape moves into it.
        var (before, after) = SaveEdited(ws => ws.Shapes.Shape("Left").MoveTo(30, 40));

        await Assert.That(ShapeChanges(before, after, "Left")).IsEqualTo(XmlChangeSet.Expect(
            "modified /xdr:sp[1]/xdr:spPr[1]/a:xfrm[1]/a:off[1] @x: '762000' -> '285750', @y: '2794000' -> '381000'"));
        var anchor = Shape(after, "Left").Parent!;
        await Assert.That(anchor.Name.LocalName).IsEqualTo("absoluteAnchor");
        await Assert.That(anchor.Element(Xdr + "pos")!.Attribute("x")!.Value).IsEqualTo("285750");
        await Assert.That(ConnectorEnds(after)).IsEqualTo(("6", "7"));
    }

    [Test]
    public async Task Moving_a_one_cell_shape_leaves_its_extent_alone()
    {
        using var source = TestHelper.OpenResource(@"TryToLoad\textbox_shapemissing_onload_2377.xlsx");
        var before = source.ReadPart(Drawing1);
        using var wb = new XLWorkbook(source);
        var ws = wb.Worksheets.First();
        ws.Shapes.Single(s => s.Placement == XLPicturePlacement.Move).MoveTo(ws.Cell("D9"), 0, 0);
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        var changes = XmlChangeSet.Between(before, saved.ReadPart(Drawing1)).Describe();
        await Assert.That(changes.Any(c => c.Contains("/xdr:from[1]"))).IsTrue();
        await Assert.That(changes.Where(c => !c.Contains("/xdr:from[1]"))).IsEmpty();
    }

    [Test]
    public async Task Replacing_the_text_keeps_the_paragraph_properties_the_body_had()
    {
        var (before, after) = SaveEdited(ws => ws.Shapes.Shape("Arc").TextBody.Text = "Bow");

        // The empty body's paragraph had algn="l"; the rebuilt one still has it.
        var paragraph = XElement.Parse(ShapeXml(after, "Arc")).Descendants(A + "p").Single();
        await Assert.That((string?)paragraph.Element(A + "pPr")?.Attribute("algn")).IsEqualTo("l");
    }

    [Test]
    public async Task A_deleted_or_renamed_loaded_shape_frees_its_name_for_a_picture()
    {
        using var wb = TestHelper.LoadWorkbook(Fixture);
        var ws = wb.Worksheet(1);
        ws.Shapes.Delete("Arc");
        ws.Shapes.Shape("Left").Name = "Port";
        using var image = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Images\SampleImagePng.png"));

        await Assert.That(ws.AddPicture(image, "Arc").Name).IsEqualTo("Arc");
        await Assert.That(ws.AddPicture(image, "Left").Name).IsEqualTo("Left");
        await Assert.That(() => ws.AddPicture(image, "Right")).Throws<ArgumentException>();
    }

    [Test]
    public async Task New_runs_added_together_each_take_the_formatting_the_paragraph_had()
    {
        var (_, after) = SaveEdited(ws =>
        {
            var paragraph = ws.Shapes.Shape("Styled").TextBody[1];
            paragraph.AddText(" bold").SetBold();
            paragraph.AddText(" plain");
        });

        var runs = XElement.Parse(ShapeXml(after, "Styled")).Descendants(A + "p").ElementAt(1).Elements(A + "r").ToList();
        await Assert.That(runs[1].Element(A + "rPr")!.Attribute("b")!.Value).IsEqualTo("1");
        await Assert.That(runs[2].Element(A + "rPr")!.Attribute("b")).IsNull();
    }

    [Test]
    public async Task Editing_a_connected_shape_keeps_the_ids_the_connector_names()
    {
        var (before, after) = SaveEdited(ws =>
        {
            ws.Shapes.Shape("Left").FillColor = XLColor.Red;
            ws.Shapes.Shape("Right").Name = "Right side";
        });

        await Assert.That(ConnectorEnds(after)).IsEqualTo(("6", "7"));
        await Assert.That(ShapeId(after, "Left")).IsEqualTo("6");
        await Assert.That(ShapeId(after, "Right side")).IsEqualTo("7");
        await Assert.That(ConnectorEnds(before)).IsEqualTo(("6", "7"));
    }

    [Test]
    public async Task Text_added_to_a_loaded_paragraph_takes_the_formatting_of_the_run_before_it()
    {
        var (before, after) = SaveEdited(ws => ws.Shapes.Shape("Styled").TextBody[1].AddText(" more"));

        await Assert.That(ShapeChanges(before, after, "Styled")).IsEqualTo(XmlChangeSet.Expect(
            "added /xdr:sp[1]/xdr:txBody[1]/a:p[2]/a:r[2] (subtree)"));
        var added = XElement.Parse(ShapeXml(after, "Styled")).Descendants(A + "p").ElementAt(1).Elements(A + "r").Last();
        await Assert.That(added.ToString(System.Xml.Linq.SaveOptions.DisableFormatting)).Contains("<a:rPr lang=\"en-AU\" sz=\"1100\" /><a:t> more</a:t>");
    }

    [Test]
    public async Task The_first_paragraph_added_to_an_empty_text_box_goes_into_its_empty_paragraph()
    {
        var (before, after) = SaveEdited(ws => ws.Shapes.Shape("Empty").TextBody.AddParagraph("Now filled"));

        await Assert.That(ShapeChanges(before, after, "Empty")).IsEqualTo(XmlChangeSet.Expect(
            "added /xdr:sp[1]/xdr:txBody[1]/a:p[1]/a:r[1] (subtree)"));
    }

    [Test]
    public async Task Deleting_a_shape_removes_its_anchor_and_nothing_else()
    {
        var (before, after) = SaveEdited(ws => ws.Shapes.Delete("Empty"));

        await Assert.That(XmlChangeSet.Between(before, after).ToString()).IsEqualTo(XmlChangeSet.Expect(
            "removed /xdr:wsDr[1]/xdr:twoCellAnchor[8] (subtree)"));
    }

    [Test]
    public async Task A_shape_added_to_a_loaded_drawing_takes_the_next_free_id()
    {
        var (before, after) = SaveEdited(ws => ws.Shapes.AddTextBox("New", ws.Cell("K2")).Name = "New one");

        await Assert.That(XmlChangeSet.Between(before, after).ToString()).IsEqualTo(XmlChangeSet.Expect(
            "added /xdr:wsDr[1]/xdr:oneCellAnchor[1] (subtree)"));
        await Assert.That(ShapeId(after, "New one")).IsEqualTo("10");
    }

    [Test]
    public async Task A_second_save_after_an_edit_leaves_the_drawing_part_as_the_first_wrote_it()
    {
        using var source = TestHelper.OpenResource(Fixture);
        using var wb = new XLWorkbook(source);
        var ws = wb.Worksheet(1);
        ws.Shapes.Shape("Styled").TextBody[0][0].Text = "Heavy";
        ws.Shapes.AddTextBox("New", ws.Cell("K2"));
        using var first = new MemoryStream();
        wb.SaveAs(first);
        using var second = new MemoryStream();
        wb.SaveAs(second);

        await Assert.That(Convert.ToHexString(second.PartBytes(Drawing1))).IsEqualTo(Convert.ToHexString(first.PartBytes(Drawing1)));
        await Assert.That(ws.Shapes.Shape("TextBox 8").Id).IsEqualTo(10);
    }

    [Test]
    public async Task The_text_box_regression_file_keeps_its_shapes_when_one_is_edited()
    {
        using var wb = TestHelper.LoadWorkbook(@"TryToLoad\textbox_shapemissing_onload_2377.xlsx");
        var ws = wb.Worksheets.First();
        var box = ws.Shapes.Single(s => s.TextBody.Text == "TEXTBOX");
        box.TextBody[0][0].Text = "EDITED";

        using var saved = new MemoryStream();
        wb.SaveAs(saved);
        using var reloaded = new XLWorkbook(saved);

        await Assert.That(reloaded.Worksheets.First().Shapes.Select(s => s.TextBody.Text))
            .IsEquivalentTo(new[] { "SHAPE", "EDITED" });
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

    private static string ShapeChanges(string before, string after, string name) =>
        XmlChangeSet.Between(ShapeXml(before, name), ShapeXml(after, name)).ToString();

    private static string ShapeXml(string drawing, string name) => Shape(drawing, name).ToString();

    private static string SpPr(string drawing, string name) => Shape(drawing, name).Element(Xdr + "spPr")!.ToString();

    private static XElement Shape(string drawing, string name) =>
        XDocument.Parse(drawing).Descendants(Xdr + "sp")
            .Single(sp => (string?)sp.Descendants(Xdr + "cNvPr").Single().Attribute("name") == name);

    private static string ShapeId(string drawing, string name) =>
        Shape(drawing, name).Descendants(Xdr + "cNvPr").Single().Attribute("id")!.Value;

    private static (string Start, string End) ConnectorEnds(string drawing)
    {
        var connector = XDocument.Parse(drawing).Descendants(Xdr + "cxnSp").Single();
        return (connector.Descendants(A + "stCxn").Single().Attribute("id")!.Value,
            connector.Descendants(A + "endCxn").Single().Attribute("id")!.Value);
    }
}
