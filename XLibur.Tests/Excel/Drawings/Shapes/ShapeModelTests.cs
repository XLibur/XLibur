using System;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.Drawings;

namespace XLibur.Tests.Excel.Drawings.Shapes;

/// <summary>The shape model in memory: <c>ws.Shapes</c>, a shape's geometry and formatting, and its text (spec 15).</summary>
public class ShapeModelTests
{
    [Test]
    public async Task A_text_box_added_at_a_cell_holds_its_text_and_moves_with_the_cell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var box = ws.Shapes.AddTextBox("Hello", ws.Cell("B2"));

        await Assert.That(ws.Shapes.Count).IsEqualTo(1);
        await Assert.That(box.Text).IsEqualTo("Hello");
        await Assert.That(box.ShapeType).IsEqualTo(XLShapeType.TextBox);
        await Assert.That(box.PresetGeometry).IsEqualTo("rect");
        await Assert.That(box.Placement).IsEqualTo(XLPicturePlacement.Move);
        await Assert.That(box.TopLeftCell.Address.ToString()).IsEqualTo("B2");
        await Assert.That(box.Name).IsEqualTo("TextBox 1");
        await Assert.That(box.Worksheet).IsSameReferenceAs(ws);
    }

    [Test]
    public async Task A_free_floating_text_box_sits_where_it_was_put()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var box = ws.Shapes.AddTextBox("Hi", 30, 40).WithSize(220, 60);

        await Assert.That(box.Placement).IsEqualTo(XLPicturePlacement.FreeFloating);
        await Assert.That((box.Left, box.Top, box.Width, box.Height)).IsEqualTo((30, 40, 220, 60));
    }

    [Test]
    [Arguments(XLShapeType.Rectangle, "rect")]
    [Arguments(XLShapeType.Ellipse, "ellipse")]
    [Arguments(XLShapeType.WedgeRectCallout, "wedgeRectCallout")]
    [Arguments(XLShapeType.RoundedCallout, "wedgeRoundRectCallout")]
    [Arguments(XLShapeType.Star5, "star5")]
    public async Task A_preset_shape_has_its_geometry(XLShapeType type, string preset)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var shape = ws.Shapes.Add(type, ws.Cell("C3"));

        await Assert.That(shape.ShapeType).IsEqualTo(type);
        await Assert.That(shape.PresetGeometry).IsEqualTo(preset);
        await Assert.That(shape is IXLTextBox).IsFalse();
    }

    [Test]
    public async Task Any_drawingml_preset_can_be_added_and_reads_back_as_custom()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var moon = ws.Shapes.AddPreset("moon", ws.Cell("A1"));

        await Assert.That(moon.ShapeType).IsEqualTo(XLShapeType.Custom);
        await Assert.That(moon.PresetGeometry).IsEqualTo("moon");
        await Assert.That(() => ws.Shapes.AddPreset("notAShape", ws.Cell("A1"))).Throws<ArgumentException>();
        await Assert.That(() => ws.Shapes.Add(XLShapeType.Custom, ws.Cell("A1"))).Throws<ArgumentException>();
    }

    [Test]
    public async Task A_new_shape_is_named_after_its_type_and_avoids_names_in_use()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.AddPicture(Image(), "Oval 2");

        var first = ws.Shapes.Add(XLShapeType.Ellipse, ws.Cell("A1"));
        var second = ws.Shapes.Add(XLShapeType.Ellipse, ws.Cell("A1"));

        await Assert.That(first.Name).IsEqualTo("Oval 1");
        await Assert.That(second.Name).IsEqualTo("Oval 3");
        await Assert.That(() => ws.AddPicture(Image(), "oval 3"))
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task Shapes_are_found_and_deleted_by_name_in_any_case()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var a = ws.Shapes.Add(XLShapeType.Rectangle, ws.Cell("A1"));
        var b = ws.Shapes.Add(XLShapeType.Rectangle, ws.Cell("A1"));
        var c = ws.Shapes.Add(XLShapeType.Rectangle, ws.Cell("A1"));
        a.Name = "Twin";
        b.Name = "twin";

        await Assert.That(ws.Shapes.Shape("TWIN")).IsSameReferenceAs(a);
        await Assert.That(ws.Shapes.Contains("tWiN")).IsTrue();

        c.Delete();
        await Assert.That(ws.Shapes.Count).IsEqualTo(2);

        ws.Shapes.Delete("Twin");
        await Assert.That(ws.Shapes.Count).IsEqualTo(0);
        await Assert.That(ws.Shapes.TryGetShape("twin", out _)).IsFalse();
        await Assert.That(() => ws.Shapes.Delete("twin")).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task A_fill_colour_and_no_fill_exclude_each_other_and_so_do_line_colour_and_no_line()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var shape = ws.Shapes.Add(XLShapeType.Rectangle, ws.Cell("A1"));

        shape.NoFill = true;
        shape.FillColor = XLColor.Red;
        shape.LineColor = XLColor.Blue;
        shape.NoLine = true;

        await Assert.That(shape.NoFill).IsFalse();
        await Assert.That(shape.FillColor).IsEqualTo(XLColor.Red);
        await Assert.That(shape.LineColor).IsNull();
        await Assert.That(shape.NoLine).IsTrue();
        await Assert.That(() => shape.LineWidthPt = -1).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task Text_is_built_by_paragraph_and_run()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var shape = ws.Shapes.Add(XLShapeType.WedgeRectCallout, ws.Cell("E4"));

        var paragraph = shape.TextBody.AddParagraph();
        paragraph.Alignment = XLTextAlignment.Center;
        paragraph.AddText("Revised ").SetBold().SetFontSize(14);
        paragraph.AddText("2026-08-01");
        shape.TextBody.AddParagraph("Second");

        await Assert.That(shape.TextBody.Count).IsEqualTo(2);
        await Assert.That(shape.TextBody.Text).IsEqualTo("Revised 2026-08-01\nSecond");
        await Assert.That(shape.TextBody[0].Count).IsEqualTo(2);
        await Assert.That(shape.TextBody[0][0].Bold).IsTrue();
        await Assert.That(shape.TextBody[0][0].FontSize).IsEqualTo(14);
        await Assert.That(shape.TextBody[0][1].Bold).IsNull();
        await Assert.That(() => shape.TextBody[0][0].FontSize = 0).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task Setting_the_text_replaces_every_paragraph()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var box = ws.Shapes.AddTextBox("one", ws.Cell("A1"));

        box.Text = "first\r\nsecond\n";

        await Assert.That(box.TextBody.Count).IsEqualTo(3);
        await Assert.That(box.Text).IsEqualTo("first\nsecond\n");

        box.TextBody.Clear();
        await Assert.That(box.TextBody.Count).IsEqualTo(0);
        await Assert.That(box.Text).IsEqualTo("");
    }

    [Test]
    public async Task Text_that_XML_cannot_hold_is_refused_where_it_is_set()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var box = ws.Shapes.AddTextBox("ok", ws.Cell("A1"));

        // A line break inside a paragraph reads as \v, which cannot be written back.
        await Assert.That(() => box.Text = "a\vb").Throws<ArgumentException>();
        await Assert.That(() => box.TextBody.AddParagraph("a\u0001b")).Throws<ArgumentException>();
        await Assert.That(() => box.TextBody[0].AddText("\0")).Throws<ArgumentException>();
        await Assert.That(() => box.TextBody[0][0].Text = "a\vb").Throws<ArgumentException>();
        await Assert.That(() => box.Text = "tab\tis fine").ThrowsNothing();
    }

    private static System.IO.Stream Image() =>
        TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Images\SampleImagePng.png"));

    [Test]
    public async Task A_shape_sized_by_its_cells_cannot_be_given_a_size()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var shape = ws.Shapes.Add(XLShapeType.Rectangle, ws.Cell("A1")).MoveTo(ws.Cell("B2"), ws.Cell("D5"));

        await Assert.That(shape.Placement).IsEqualTo(XLPicturePlacement.MoveAndSize);
        await Assert.That(shape.BottomRightCell.Address.ToString()).IsEqualTo("D5");
        await Assert.That(() => shape.Width = 10).Throws<ArgumentException>();
        await Assert.That(() => shape.MoveTo(ws.Cell("A1"), wb.AddWorksheet().Cell("B2"))).Throws<ArgumentException>();
    }
}
