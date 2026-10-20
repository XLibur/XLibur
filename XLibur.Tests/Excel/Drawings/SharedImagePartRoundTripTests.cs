using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
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
    public async Task A_save_that_fails_leaves_the_next_save_finding_each_picture_on_its_own_anchor()
    {
        // The encrypted save writes the drawing, renumbering its ids, then fails to create the file
        // over a directory. The next save copies the source again, which still has the old ids.
        using var wb = LoadFixture();
        using var saved = new MemoryStream();
        var directory = Directory.CreateTempSubdirectory().FullName;
        try
        {
            Action failedSave = () => wb.SaveAs(directory, new XLibur.Excel.SaveOptions { Password = "secret" });
            await Assert.That(failedSave).Throws<Exception>();

            wb.SaveAs(saved);
        }
        finally
        {
            Directory.Delete(directory);
        }

        await Assert.That(Anchors(saved)).IsEqualTo(Lines(Picture1Anchor, Picture2Anchor, Picture3Anchor));
    }

    [Test]
    public async Task A_picture_inside_alternate_content_is_written_back_into_its_own_anchor()
    {
        // Excel wraps a drawing in mc:AlternateContent when it also has a newer form, as a 3D model
        // has, so the xdr:pic is not a direct child of its anchor.
        using var source = AlternateContentFixture();
        using var wb = new XLWorkbook(source);
        await Assert.That(wb.Worksheet(1).Pictures.Count).IsEqualTo(1);

        using var first = new MemoryStream();
        using var second = new MemoryStream();
        wb.SaveAs(first);
        wb.SaveAs(second);

        var drawing = XDocument.Parse(second.ReadPart(Drawing1));
        await Assert.That(drawing.Root!.Elements().Count()).IsEqualTo(1);
        await Assert.That(drawing.Descendants(XdrNs + "pic").Count()).IsEqualTo(1);
    }

    [Test]
    public async Task Changing_the_image_of_one_copy_changes_that_picture_only()
    {
        // The three pictures share one image part. A copy whose image was changed takes a part of
        // its own, so the change neither reaches the other copies nor is overwritten by them.
        using var wb = LoadFixture();
        var picture2 = wb.Worksheet(1).Picture("Picture 2");
        var original = picture2.ImageStream.ToArray();
        var changed = original.Concat(new byte[] { 1, 2, 3 }).ToArray();
        picture2.ImageStream.SetLength(0);
        picture2.ImageStream.Write(changed);

        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        using var reloaded = new XLWorkbook(saved);
        var images = reloaded.Worksheet(1).Pictures
            .Select(p => $"{p.Name} {(p.ImageStream.ToArray().SequenceEqual(changed) ? "changed" : p.ImageStream.ToArray().SequenceEqual(original) ? "original" : "other")}")
            .ToArray();
        await Assert.That(Lines(images)).IsEqualTo(Lines(
            "Picture 1 original",
            "Picture 2 changed",
            "Picture 3 original"));
        await Assert.That(saved.PartsUnder("xl/media/").Length).IsEqualTo(2);
    }

    [Test]
    public async Task Changing_two_copies_to_the_same_image_puts_them_on_one_new_part()
    {
        // The first changed copy takes a new part. The second, changed to the same bytes in a
        // stream of its own, finds that part by its bytes and shows it too.
        using var wb = LoadFixture();
        var ws = wb.Worksheet(1);
        var original = ws.Picture("Picture 2").ImageStream.ToArray();
        var changed = original.Concat(new byte[] { 1, 2, 3 }).ToArray();
        foreach (var name in new[] { "Picture 2", "Picture 3" })
        {
            var image = ws.Picture(name).ImageStream;
            image.SetLength(0);
            image.Write(changed);
        }

        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        // The image link of Picture 1, 2 and 3, in anchor order.
        var embeds = XDocument.Parse(saved.ReadPart(Drawing1))
            .Descendants(ANs + "blip")
            .Select(b => (string?)b.Attribute(RNs + "embed"))
            .ToArray();
        await Assert.That(embeds[1]).IsNotEqualTo(embeds[0]);
        await Assert.That(embeds[2]).IsEqualTo(embeds[1]);
        await Assert.That(saved.PartsUnder("xl/media/").Length).IsEqualTo(2);

        using var reloaded = new XLWorkbook(saved);
        var images = reloaded.Worksheet(1).Pictures
            .Select(p => $"{p.Name} {(p.ImageStream.ToArray().SequenceEqual(changed) ? "changed" : "original")}")
            .ToArray();
        await Assert.That(Lines(images)).IsEqualTo(Lines(
            "Picture 1 original",
            "Picture 2 changed",
            "Picture 3 changed"));
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

    private static readonly byte[] Png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEklEQVR4nGP4z8CAFTEMLQkAI7c/wTpVXRAAAAAASUVORK5CYII=");

    /// <summary>One picture whose <c>xdr:pic</c> sits in the <c>mc:Choice</c> of an <c>mc:AlternateContent</c>.</summary>
    private static MemoryStream AlternateContentFixture()
    {
        var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Sheet1");
            ws.AddPicture(new MemoryStream(Png), "Logo").MoveTo(ws.Cell("B2"));
            wb.SaveAs(package);
        }

        return package.RewritePart(Drawing1, xml => Regex.Replace(xml, "<xdr:pic>.*</xdr:pic>", match =>
            "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">" +
            "<mc:Choice xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" Requires=\"a14\">" +
            match.Value + "</mc:Choice><mc:Fallback /></mc:AlternateContent>", RegexOptions.Singleline));
    }

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
