using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;
using XLibur.Tests.Utils;

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
        using var saved = LoadAndSave(ws => ws.Pictures.Delete("3D Model 1"));

        await Assert.That(Shape(saved)).IsEqualTo("0 anchors");
    }

    private static XLWorkbook LoadFixture() => new(Fixture());

    private static MemoryStream Fixture()
    {
        using var resource = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Drawings\picture-3d-model.xlsx"));
        var package = new MemoryStream();
        resource.CopyTo(package);
        package.Position = 0;
        return package;
    }

    private static MemoryStream LoadAndSave(Action<IXLWorksheet> edit)
    {
        var saved = new MemoryStream();
        using var wb = LoadFixture();
        edit(wb.Worksheet(1));
        wb.SaveAs(saved);
        return saved;
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
        if (content.Name != McNs + "AlternateContent")
            return $"{anchor.Name.LocalName} > {Drawing(content)}";

        var branches = content.Elements().Select(branch =>
        {
            var requires = (string?)branch.Attribute("Requires");
            var label = requires is null ? branch.Name.LocalName : $"{branch.Name.LocalName}({requires})";
            var drawing = branch.Elements().SingleOrDefault();
            return drawing is null ? label : $"{label} {Drawing(drawing)}";
        });
        return $"{anchor.Name.LocalName} > AlternateContent > {string.Join(" | ", branches)}";

        static string Drawing(XElement drawing)
        {
            var properties = drawing.Descendants(XdrNs + "cNvPr").First();
            return $"{drawing.Name.LocalName} {(string?)properties.Attribute("id")} '{(string?)properties.Attribute("name")}'";
        }
    }
}
