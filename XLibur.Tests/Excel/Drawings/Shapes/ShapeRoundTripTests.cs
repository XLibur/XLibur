using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.Drawings;

namespace XLibur.Tests.Excel.Drawings.Shapes;

/// <summary>
/// A shape saved by XLibur loads back with its type, geometry, fill, line and text (spec 15,
/// acceptance criterion 2), and the shapes Excel wrote load into the model.
/// </summary>
public class ShapeRoundTripTests
{
    public static IEnumerable<XLShapeType> PresetTypes() =>
        Enum.GetValues<XLShapeType>().Where(type => type != XLShapeType.Custom);

    [Test]
    [MethodDataSource(nameof(PresetTypes))]
    public async Task Every_shape_type_reads_back_as_it_was_written(XLShapeType type)
    {
        using var reloaded = SaveAndReload(ws => Format(ws.Shapes.Add(type, ws.Cell("C3"))));

        var shape = reloaded.Worksheet(1).Shapes.Single();
        await Assert.That(shape.ShapeType).IsEqualTo(type);
        await Assert.That(shape is IXLTextBox).IsEqualTo(type == XLShapeType.TextBox);
        await AssertFormatted(shape);
    }

    [Test]
    public async Task A_custom_preset_reads_back_with_its_geometry()
    {
        using var reloaded = SaveAndReload(ws => Format(ws.Shapes.AddPreset("moon", ws.Cell("C3"))));

        var shape = reloaded.Worksheet(1).Shapes.Single();
        await Assert.That(shape.ShapeType).IsEqualTo(XLShapeType.Custom);
        await Assert.That(shape.PresetGeometry).IsEqualTo("moon");
        await AssertFormatted(shape);
    }

    [Test]
    public async Task Each_placement_reads_back_with_its_geometry()
    {
        using var reloaded = SaveAndReload(ws =>
        {
            ws.Shapes.AddTextBox("free", 30, 40).WithSize(120, 50).Name = "Free";
            ws.Shapes.AddTextBox("move", ws.Cell("C3")).MoveTo(ws.Cell("C3"), 5, 6).WithSize(110, 45).Name = "Move";
            ws.Shapes.Add(XLShapeType.Rectangle, ws.Cell("A1")).MoveTo(ws.Cell("B2"), ws.Cell("D5")).Name = "Size";
        });

        var shapes = reloaded.Worksheet(1).Shapes;
        var free = shapes.Shape("Free");
        await Assert.That((free.Placement, free.Left, free.Top, free.Width, free.Height))
            .IsEqualTo((XLPicturePlacement.FreeFloating, 30, 40, 120, 50));

        var move = shapes.Shape("Move");
        await Assert.That((move.Placement, move.TopLeftCell.Address.ToString(), move.Left, move.Top, move.Width, move.Height))
            .IsEqualTo((XLPicturePlacement.Move, "C3", 5, 6, 110, 45));

        var size = shapes.Shape("Size");
        await Assert.That((size.Placement, size.TopLeftCell.Address.ToString(), size.BottomRightCell.Address.ToString()))
            .IsEqualTo((XLPicturePlacement.MoveAndSize, "B2", "D5"));
        await Assert.That((size.Width, size.Height)).IsEqualTo((128, 60));
    }

    [Test]
    public async Task No_fill_and_no_line_read_back()
    {
        using var reloaded = SaveAndReload(ws =>
        {
            var shape = ws.Shapes.Add(XLShapeType.Ellipse, ws.Cell("A1"));
            shape.NoFill = true;
            shape.NoLine = true;
        });

        var shape = reloaded.Worksheet(1).Shapes.Single();
        await Assert.That((shape.NoFill, shape.NoLine, shape.FillColor, shape.LineColor)).IsEqualTo((true, true, (XLColor?)null, (XLColor?)null));
    }

    [Test]
    public async Task An_empty_shape_reads_back_with_no_paragraph_and_takes_the_first_one_added()
    {
        using var reloaded = SaveAndReload(ws => ws.Shapes.Add(XLShapeType.Rectangle, ws.Cell("A1")));
        var shape = reloaded.Worksheet(1).Shapes.Single();
        await Assert.That(shape.TextBody.Count).IsEqualTo(0);

        shape.TextBody.AddParagraph("first");
        using var again = SaveAndReload(reloaded);

        var reread = again.Worksheet(1).Shapes.Single();
        await Assert.That(reread.TextBody.Text).IsEqualTo("first");
        await Assert.That(reread.TextBody.Count).IsEqualTo(1);
    }

    [Test]
    public async Task The_shapes_Excel_wrote_load_into_the_model()
    {
        using var wb = new XLWorkbook(TestHelper.OpenResource(ShapeCreationTests.Fixture));
        var shapes = wb.Worksheet(1).Shapes;

        // The connector is not a shape XLibur models; it stays as the file has it.
        await Assert.That(shapes.Select(s => s.Name))
            .IsEquivalentTo(new[] { "Plain TextBox", "Plain Rectangle", "Styled", "Arc", "Left", "Right", "Empty" });
        await Assert.That(shapes.Select(s => s.Id)).IsEquivalentTo(new[] { 2, 3, 4, 5, 6, 7, 9 });

        var styled = (IXLTextBox)shapes.Shape("Styled");
        await Assert.That(styled.Text).IsEqualTo("Bold plain\nSecond");
        await Assert.That(styled.TextBody[0][0].Bold).IsTrue();
        await Assert.That(styled.TextBody[0][0].FontColor).IsEqualTo(XLColor.FromHtml("#FF0000"));
        await Assert.That(styled.TextBody[0][0].FontSize).IsEqualTo(11);
        await Assert.That(styled.TextBody[0][1].Bold).IsNull();
        await Assert.That(styled.TextBody[1].Alignment).IsEqualTo(XLTextAlignment.Center);
        await Assert.That(styled.FillColor).IsNull();
        await Assert.That(styled.LineWidthPt).IsEqualTo(0.75);
        await Assert.That(styled.Placement).IsEqualTo(XLPicturePlacement.MoveAndSize);

        await Assert.That(shapes.Shape("Arc").ShapeType).IsEqualTo(XLShapeType.Custom);
        await Assert.That(shapes.Shape("Arc").LineWidthPt).IsEqualTo(2);
        await Assert.That(shapes.Shape("Plain Rectangle").ShapeType).IsEqualTo(XLShapeType.Rectangle);
        await Assert.That(shapes.Shape("Plain TextBox").FillColor).IsEqualTo(XLColor.FromTheme(XLThemeColor.Background1));
        await Assert.That(((IXLTextBox)shapes.Shape("Empty")).Text).IsEqualTo("");
        await Assert.That(shapes.Shape("Empty").TextBody.Count).IsEqualTo(0);
    }

    private static void Format(IXLShape shape)
    {
        shape.WithSize(150, 80);
        shape.FillColor = XLColor.Red;
        shape.LineColor = XLColor.Blue;
        shape.LineWidthPt = 1.5;
        shape.LineDash = XLLineDashStyle.Dash;
        shape.TextBody.VerticalAlignment = XLTextVerticalAlignment.Center;
        shape.TextBody.WrapText = false;
        var paragraph = shape.TextBody.AddParagraph();
        paragraph.Alignment = XLTextAlignment.Right;
        var run = paragraph.AddText("Hi").SetBold().SetItalic().SetFontSize(12.5).SetFontColor(XLColor.Green).SetFontName("Arial");
        run.Underline = XLShapeTextUnderline.Double;
        run.Strikethrough = true;
        paragraph.AddText(" there");
        shape.TextBody.AddParagraph("Second");
    }

    private static async Task AssertFormatted(IXLShape shape)
    {
        await Assert.That((shape.Placement, shape.TopLeftCell.Address.ToString(), shape.Width, shape.Height))
            .IsEqualTo((XLPicturePlacement.Move, "C3", 150, 80));
        await Assert.That(shape.FillColor).IsEqualTo(XLColor.Red);
        await Assert.That(shape.LineColor).IsEqualTo(XLColor.Blue);
        await Assert.That(shape.LineWidthPt).IsEqualTo(1.5);
        await Assert.That(shape.LineDash).IsEqualTo(XLLineDashStyle.Dash);
        await Assert.That(shape.TextBody.VerticalAlignment).IsEqualTo(XLTextVerticalAlignment.Center);
        await Assert.That(shape.TextBody.WrapText).IsFalse();
        await Assert.That(shape.TextBody.Text).IsEqualTo("Hi there\nSecond");
        await Assert.That(shape.TextBody[0].Alignment).IsEqualTo(XLTextAlignment.Right);

        var run = shape.TextBody[0][0];
        await Assert.That((run.Bold, run.Italic, run.Underline, run.Strikethrough, run.FontSize, run.FontName))
            .IsEqualTo((true, true, XLShapeTextUnderline.Double, true, 12.5, "Arial"));
        await Assert.That(run.FontColor).IsEqualTo(XLColor.Green);
        await Assert.That(shape.TextBody[0][1].Bold).IsNull();
    }

    private static XLWorkbook SaveAndReload(Action<IXLWorksheet> build)
    {
        using var wb = new XLWorkbook();
        build(wb.AddWorksheet());
        return SaveAndReload(wb);
    }

    private static XLWorkbook SaveAndReload(XLWorkbook wb)
    {
        var saved = new MemoryStream();
        wb.SaveAs(saved);
        saved.Position = 0;
        return new XLWorkbook(saved);
    }
}
