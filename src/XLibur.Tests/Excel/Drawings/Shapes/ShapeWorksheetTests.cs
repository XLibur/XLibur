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
/// Shapes take part in what the worksheet does: they follow the cells they hang off when rows and
/// columns are inserted or deleted, a copy of the sheet copies them, and a workbook opened as a
/// template saves them whole.
/// </summary>
public class ShapeWorksheetTests
{
    [Test]
    public async Task A_shape_follows_its_cells_when_rows_and_columns_are_inserted()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var moving = ws.Shapes.AddTextBox("moves", ws.Cell("B10"));
        var sizing = ws.Shapes.Add(XLShapeType.Rectangle, ws.Cell("A1")).MoveTo(ws.Cell("C10"), ws.Cell("D12"));
        var floating = ws.Shapes.AddTextBox("stays", 30, 40);

        ws.Row(5).InsertRowsAbove(1);
        ws.Column(1).InsertColumnsBefore(2);

        await Assert.That(moving.TopLeftCell.Address.ToString()).IsEqualTo("D11");
        await Assert.That((sizing.TopLeftCell.Address.ToString(), sizing.BottomRightCell.Address.ToString()))
            .IsEqualTo(("E11", "F13"));
        await Assert.That((floating.Left, floating.Top)).IsEqualTo((30, 40));
    }

    [Test]
    public async Task A_loaded_shape_moves_in_the_file_when_rows_are_inserted_above_it()
    {
        using var source = TestHelper.OpenResource(Fixture);
        using var wb = new XLWorkbook(source);
        wb.Worksheet(1).Row(1).InsertRowsAbove(2);
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        var left = XDocument.Parse(saved.ReadPart(Drawing1)).Descendants(Xdr + "sp")
            .Single(sp => (string?)sp.Descendants(Xdr + "cNvPr").Single().Attribute("name") == "Left").Parent!;
        await Assert.That(left.Element(Xdr + "from")!.Element(Xdr + "row")!.Value).IsEqualTo("17");
        await Assert.That(left.Element(Xdr + "to")!.Element(Xdr + "row")!.Value).IsEqualTo("20");
    }

    [Test]
    public async Task Copying_a_worksheet_copies_its_shapes_with_their_text_and_formatting()
    {
        using var wb = TestHelper.LoadWorkbook(Fixture);
        var ws = wb.Worksheet(1);
        var added = ws.Shapes.Add(XLShapeType.Ellipse, ws.Cell("K2"));
        added.FillColor = XLColor.Red;
        added.TextBody.AddParagraph("new one");
        ws.Shapes.Shape("Arc").LineWidthPt = 4;

        ws.CopyTo("Copy");
        using var saved = new MemoryStream();
        wb.SaveAs(saved);
        using var reloaded = new XLWorkbook(saved);
        var copy = reloaded.Worksheet("Copy").Shapes;

        await Assert.That(copy.Select(s => s.Name)).IsEquivalentTo(ws.Shapes.Select(s => s.Name));
        await Assert.That(((IXLTextBox)copy.Shape("Styled")).Text).IsEqualTo("Bold plain\nSecond");
        await Assert.That(copy.Shape("Styled").TextBody[0][0].Bold).IsTrue();
        await Assert.That(copy.Shape("Arc").LineWidthPt).IsEqualTo(4);
        await Assert.That(copy.Shape("Oval 8").FillColor).IsEqualTo(XLColor.Red);
        await Assert.That(copy.Shape("Oval 8").TextBody.Text).IsEqualTo("new one");
        await Assert.That(copy.Shape("Left").TopLeftCell.Address.ToString()).IsEqualTo("B16");

        // What XLibur does not model comes with the copy: the gradient and the arrowheads.
        var copyDrawing = saved.ReadPart("xl/drawings/drawing2.xml");
        await Assert.That(copyDrawing).Contains("gradFill");
        await Assert.That(copyDrawing).Contains("headEnd");
    }

    [Test]
    public async Task A_workbook_opened_as_a_template_saves_its_shapes_whole()
    {
        using var file = new TemporaryFile();
        using (var resource = TestHelper.OpenResource(Fixture))
        using (var output = File.Create(file.Path))
            resource.CopyTo(output);

        using var wb = XLWorkbook.OpenFromTemplate(file.Path);
        wb.Worksheet(1).Shapes.Shape("Styled").TextBody[1][0].Text = "Changed";
        using var saved = new MemoryStream();
        wb.SaveAs(saved);
        using var reloaded = new XLWorkbook(saved);
        var shapes = reloaded.Worksheet(1).Shapes;

        await Assert.That(shapes.Count).IsEqualTo(7);
        await Assert.That(((IXLTextBox)shapes.Shape("Styled")).Text).IsEqualTo("Bold plain\nChanged");
        await Assert.That(shapes.Shape("Styled").TextBody[0][0].FontColor).IsEqualTo(XLColor.FromHtml("#FF0000"));
        await Assert.That(saved.ReadPart(Drawing1)).Contains("gradFill");
    }
}
