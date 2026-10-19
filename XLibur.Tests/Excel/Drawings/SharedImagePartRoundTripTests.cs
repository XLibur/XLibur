using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;
using XLibur.Excel.Drawings;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Drawings;

/// <summary>
/// Pictures that share one image part each keep their own anchor through a load and save
/// (issue #781). Excel shares the part when a picture is copied and pasted: every copy's
/// <c>a:blip</c> embeds the same relationship.
/// </summary>
public class SharedImagePartRoundTripTests
{
    private const string Drawing1 = "xl/drawings/drawing1.xml";
    private const string Drawing1Rels = "xl/drawings/_rels/drawing1.xml.rels";

    private static readonly XNamespace XdrNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace RNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    // The three anchors of the Excel fixture, as Excel wrote them.
    private const string Picture1Anchor = "twoCellAnchor Picture 1 from 0,0 to 5,22";
    private const string Picture2Anchor = "twoCellAnchor editAs=oneCell Picture 2 from 5,0 to 10,22";
    private const string Picture3Anchor = "twoCellAnchor editAs=absolute Picture 3 from 11,0 to 16,22";

    [Test]
    public async Task A_load_and_save_keeps_each_picture_on_its_own_anchor()
    {
        using var wb = LoadFixture();
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        await Assert.That(Anchors(saved)).IsEqualTo(Lines(Picture1Anchor, Picture2Anchor, Picture3Anchor));
        await AssertOneSharedImagePart(saved);
    }

    [Test]
    public async Task A_second_save_keeps_each_picture_on_its_own_anchor()
    {
        // The first save renumbers the drawing ids, and the second save edits what the first wrote.
        using var wb = LoadFixture();
        using var first = new MemoryStream();
        using var second = new MemoryStream();
        wb.SaveAs(first);
        wb.SaveAs(second);

        await Assert.That(Anchors(second)).IsEqualTo(Lines(Picture1Anchor, Picture2Anchor, Picture3Anchor));
        await AssertOneSharedImagePart(second);
    }

    [Test]
    public async Task A_saved_workbook_reloads_with_every_picture()
    {
        using var wb = LoadFixture();
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        using var reloaded = new XLWorkbook(saved);
        var pictures = reloaded.Worksheet(1).Pictures.Select(p => $"{p.Name} {p.Placement}").ToArray();
        await Assert.That(Lines(pictures)).IsEqualTo(Lines(
            "Picture 1 MoveAndSize",
            "Picture 2 Move",
            "Picture 3 FreeFloating"));
    }

    [Test]
    [Arguments("Picture 1", Picture2Anchor, Picture3Anchor)]
    [Arguments("Picture 2", Picture1Anchor, Picture3Anchor)]
    [Arguments("Picture 3", Picture1Anchor, Picture2Anchor)]
    public async Task Deleting_one_picture_keeps_the_others_and_their_shared_image(
        string deleted, string firstKept, string secondKept)
    {
        using var wb = LoadFixture();
        wb.Worksheet(1).Pictures.Delete(deleted);
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        await Assert.That(Anchors(saved)).IsEqualTo(Lines(firstKept, secondKept));
        await AssertOneSharedImagePart(saved);

        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet(1).Pictures.Count).IsEqualTo(2);
    }

    [Test]
    public async Task Deleting_a_picture_after_a_save_keeps_the_others()
    {
        using var wb = LoadFixture();
        using var first = new MemoryStream();
        wb.SaveAs(first);
        wb.Worksheet(1).Pictures.Delete("Picture 1");
        using var second = new MemoryStream();
        wb.SaveAs(second);

        await Assert.That(Anchors(second)).IsEqualTo(Lines(Picture2Anchor, Picture3Anchor));
        await AssertOneSharedImagePart(second);
    }

    [Test]
    public async Task Deleting_every_picture_drops_the_shared_image()
    {
        using var wb = LoadFixture();
        var ws = wb.Worksheet(1);
        ws.Pictures.Delete("Picture 1");
        ws.Pictures.Delete("Picture 2");
        ws.Pictures.Delete("Picture 3");
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        await Assert.That(saved.PartsUnder("xl/media/").Length).IsEqualTo(0);
    }

    [Test]
    public async Task Grouping_two_pictures_moves_those_two_into_the_group()
    {
        using var wb = LoadFixture();
        var ws = wb.Worksheet(1);
        var picture1 = ws.Picture("Picture 1");
        picture1.Placement = XLPicturePlacement.FreeFloating;
        ws.Pictures.Group(picture1, ws.Picture("Picture 3"));
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        var drawing = XDocument.Parse(saved.ReadPart(Drawing1));
        var grouped = drawing.Descendants(XdrNs + "grpSp")
            .SelectMany(g => g.Elements(XdrNs + "pic"))
            .Select(p => (string?)p.Descendants(XdrNs + "cNvPr").First().Attribute("name"));
        await Assert.That(string.Join(", ", grouped)).IsEqualTo("Picture 1, Picture 3");
        await Assert.That(Anchors(saved)).Contains(Picture2Anchor);
    }

    private static XLWorkbook LoadFixture() =>
        new(TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\ExcelProducedWorkbookWithImages.xlsx")));

    /// <summary>Every picture embeds the same image, and the package holds that image once.</summary>
    private static async Task AssertOneSharedImagePart(Stream saved)
    {
        var embeds = XDocument.Parse(saved.ReadPart(Drawing1))
            .Descendants(ANs + "blip")
            .Select(b => (string?)b.Attribute(RNs + "embed"))
            .Distinct()
            .ToArray();
        await Assert.That(embeds.Length).IsEqualTo(1);

        var imageRels = XDocument.Parse(saved.ReadPart(Drawing1Rels)).Root!.Elements()
            .Where(r => ((string?)r.Attribute("Type"))?.EndsWith("/image") ?? false)
            .Select(r => (string?)r.Attribute("Id"))
            .ToArray();
        await Assert.That(string.Join(",", imageRels)).IsEqualTo(embeds[0]);
        await Assert.That(saved.PartsUnder("xl/media/").Length).IsEqualTo(1);
    }

    private static string Lines(params string[] lines) => string.Join("\n", lines);

    /// <summary>
    /// Each anchor of the saved drawing as its element and <c>editAs</c>, the picture's name and the
    /// zero-based column and row of its markers, e.g. <c>twoCellAnchor Picture 1 from 0,0 to 5,22</c>.
    /// </summary>
    private static string Anchors(Stream saved)
    {
        static string Marker(XElement anchor, string name)
        {
            var marker = anchor.Element(XdrNs + name);
            return marker is null
                ? "none"
                : $"{(string?)marker.Element(XdrNs + "col")},{(string?)marker.Element(XdrNs + "row")}";
        }

        return Lines(XDocument.Parse(saved.ReadPart(Drawing1)).Root!.Elements()
            .Select(anchor =>
            {
                var editAs = (string?)anchor.Attribute("editAs");
                var name = (string?)anchor.Descendants(XdrNs + "cNvPr").FirstOrDefault()?.Attribute("name");
                return $"{anchor.Name.LocalName}{(editAs is null ? "" : $" editAs={editAs}")} {name} " +
                       $"from {Marker(anchor, "from")} to {Marker(anchor, "to")}";
            })
            .ToArray());
    }
}
