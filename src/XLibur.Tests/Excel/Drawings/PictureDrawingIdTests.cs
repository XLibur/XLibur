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
/// The drawing ids a save gives the pictures it writes, in a drawing that holds a group shape.
/// </summary>
/// <remarks>
/// A save never renumbers a drawing: a connector names the shapes it joins by their ids, and a
/// loaded picture is written back into its own anchor under the id it was loaded with (spec 17).
/// A new picture takes the id after the highest the drawing holds at that moment (issue #797
/// keeps this while it stops scanning the drawing for each picture).
/// </remarks>
public class PictureDrawingIdTests
{
    private const string Drawing1 = "xl/drawings/drawing1.xml";

    private static readonly XNamespace XdrNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

    private static readonly byte[] Png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEklEQVR4nGP4z8CAFTEMLQkAI7c/wTpVXRAAAAAASUVORK5CYII=");

    [Test]
    public async Task New_pictures_take_the_ids_after_the_highest_one_the_drawing_holds()
    {
        // The fixture's group is id 10, holding Picture 1 (2), Picture 2 (3) and a connector (5).
        // The free pictures are written first, then the grouped ones, so Grouped, which the
        // drawing holds in the group ahead of the free pictures, takes the id after theirs.
        using var saved = WithNewPictures();

        await Assert.That(Ids(saved)).IsEqualTo(
            "Group 9 10, Picture 1 2, Picture 2 3, Straight Arrow Connector 4 5, Grouped 13, Free 1 11, Free 2 12");
    }

    [Test]
    public async Task Deleting_the_pictures_with_the_highest_ids_frees_them_for_the_pictures_written_after()
    {
        // Deleting Grouped (13) and Free 2 (12) leaves 11 the highest. Moving Free 1 writes it back
        // into its own anchor under its own id, and the picture added after it takes the next one.
        using var source = WithNewPictures();
        using var saved = LoadAndSave(source, ws =>
        {
            ws.Pictures.Delete("Grouped");
            ws.Pictures.Delete("Free 2");
            ws.Picture("Free 1").MoveTo(ws.Cell("K2"));
            ws.AddPicture(new MemoryStream(Png), "Free 3").MoveTo(ws.Cell("K20"));
        });

        await Assert.That(Ids(saved)).IsEqualTo(
            "Group 9 10, Picture 1 2, Picture 2 3, Straight Arrow Connector 4 5, Free 1 11, Free 3 12");
    }

    [Test]
    public async Task A_moved_3d_model_takes_no_id_from_the_picture_written_after_it()
    {
        // A 3D model's anchor is kept and given the new position; its drawing keeps the id Excel
        // gave it, 2, so the picture added after it takes the id after the group's, as if the
        // model had not moved.
        using var model = Resource("picture-3d-model.xlsx");
        using var withPictures = LoadAndSave(model, ws =>
        {
            ws.AddPicture(new MemoryStream(Png), "Pic A").MoveTo(500, 20);
            ws.AddPicture(new MemoryStream(Png), "Pic B").MoveTo(500, 300);
        });
        using var grouped = LoadAndSave(withPictures, ws =>
            ws.Pictures.Group(ws.Picture("Pic A"), ws.Picture("Pic B")));

        using var saved = LoadAndSave(grouped, ws =>
        {
            ws.Picture("3D Model 1").MoveTo(ws.Cell("B30"));
            ws.AddPicture(new MemoryStream(Png), "Pic C").MoveTo(ws.Cell("K2"));
        });

        await Assert.That(Ids(saved)).IsEqualTo(
            "3D Model 1 2, 3D Model 1 2, Group 5 5, Pic A 3, Pic B 4, Pic C 6");
    }

    /// <summary>
    /// The grouped-pictures fixture with two free pictures and a picture added to its group, saved.
    /// </summary>
    private static MemoryStream WithNewPictures()
    {
        using var source = Resource(@"GroupedPictures.xlsx");
        return LoadAndSave(source, ws =>
        {
            ws.AddPicture(new MemoryStream(Png), "Free 1").MoveTo(ws.Cell("K2"));
            ws.AddPicture(new MemoryStream(Png), "Free 2").MoveTo(ws.Cell("K10"));
            var pictures = (XLPictures)ws.Pictures;
            pictures.AddToGroup((XLPicture)ws.Picture("Picture 1"), new MemoryStream(Png), "Grouped");
        });
    }

    private static MemoryStream LoadAndSave(MemoryStream source, Action<IXLWorksheet> edit)
    {
        var saved = new MemoryStream();
        source.Position = 0;
        using (var wb = new XLWorkbook(source))
        {
            edit(wb.Worksheet(1));
            wb.SaveAs(saved);
        }

        return saved;
    }

    private static MemoryStream Resource(string name)
    {
        using var resource = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath($@"Other\Drawings\{name}"));
        var package = new MemoryStream();
        resource.CopyTo(package);
        package.Position = 0;
        return package;
    }

    /// <summary>Every drawing of the saved drawing part as its name and id, in document order.</summary>
    private static string Ids(Stream saved) =>
        string.Join(", ", XDocument.Parse(saved.ReadPart(Drawing1)).Descendants(XdrNs + "cNvPr")
            .Select(e => $"{(string?)e.Attribute("name")} {(string?)e.Attribute("id")}"));
}
