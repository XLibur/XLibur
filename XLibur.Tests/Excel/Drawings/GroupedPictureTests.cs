using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using XLibur.Excel;
using XLibur.Excel.Drawings;
using XLibur.Excel.IO;
using XLibur.Tests.Utils;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Tests.Excel.Drawings;

[TestFixture]
public class GroupedPictureTests
{
    private const string GroupedPicturesResource = @"Other\Drawings\GroupedPictures.xlsx";

    // The fixture's "Map" sheet has a twoCellAnchor → grpSp containing two pictures
    // (Picture 1: child ext 2_000_000 EMU, Picture 2: child ext 1_500_000 EMU) plus a
    // connector. The group is scaled 2× (ext 10_000_000 vs chExt 5_000_000 horizontally,
    // 8_000_000 vs 4_000_000 vertically), so the sheet-space sizes are the child extents × 2.

    private static MemoryStream OpenFixture()
    {
        var ms = new MemoryStream();
        using var src = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(GroupedPicturesResource));
        src.CopyTo(ms);
        ms.Position = 0;
        return ms;
    }

    [Test]
    public void GroupedPicturesAreLoadedWithGroupScaledGeometry()
    {
        using var stream = OpenFixture();
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheet("Map");

        Assert.That(ws.Pictures.Count, Is.EqualTo(2));

        var picture1 = (XLPicture)ws.Pictures.Single(p => p.Name == "Picture 1");
        var picture2 = (XLPicture)ws.Pictures.Single(p => p.Name == "Picture 2");

        Assert.That(picture1.IsInGroup, Is.True);
        Assert.That(picture2.IsInGroup, Is.True);

        // Both pictures are scaled by the same group factor, so their relative sizes are preserved:
        // Picture 1 (child 2_000_000) is larger than Picture 2 (child 1_500_000).
        Assert.That(picture1.Width, Is.GreaterThan(0));
        Assert.That(picture1.Width, Is.GreaterThan(picture2.Width));
        Assert.That(picture1.Height, Is.GreaterThan(picture2.Height));

        // Picture 1's sheet-space extent is twice its child extent (2_000_000 → 4_000_000 EMU).
        var expectedPx1 = DrawingPartReader.ConvertFromEnglishMetricUnits(4_000_000, wb.DpiX);
        Assert.That(picture1.Width, Is.EqualTo(expectedPx1));
    }

    [Test]
    public void UneditedRoundTripPreservesGroupPicturesAndShapes()
    {
        using var output = new MemoryStream();
        using (var stream = OpenFixture())
        using (var wb = new XLWorkbook(stream))
        {
            wb.SaveAs(output);
        }

        output.Position = 0;
        using var package = SpreadsheetDocument.Open(output, false);
        var drawingsPart = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart;
        Assert.That(drawingsPart, Is.Not.Null);
        var drawing = drawingsPart!.WorksheetDrawing;

        var groups = drawing.Descendants<Xdr.GroupShape>().ToList();
        Assert.That(groups.Count, Is.EqualTo(1), "group preserved");

        // Assert on the group node so the shapes are verified to remain *inside* the group rather
        // than having been moved out to the top level during the round-trip.
        var group = groups[0];
        Assert.That(group.Descendants<Xdr.Picture>().Count(), Is.EqualTo(2), "both pictures preserved inside the group");
        Assert.That(group.Descendants<Xdr.ConnectionShape>().Count(), Is.EqualTo(1), "connector preserved inside the group");

        // An unedited grouped picture must keep its exact child-space extent (no rounding drift).
        var extents = group.Descendants<Xdr.Picture>()
            .Select(p => p.ShapeProperties!.Transform2D!.Extents!)
            .Select(e => (e.Cx!.Value, e.Cy!.Value))
            .OrderByDescending(t => t.Item1)
            .ToList();
        Assert.That(extents[0], Is.EqualTo((2_000_000L, 2_000_000L)));
        Assert.That(extents[1], Is.EqualTo((1_500_000L, 1_500_000L)));

        // Both image relationships still resolve to image parts.
        var embeds = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>()
            .Select(b => b.Embed?.Value).Where(v => v is not null).ToList();
        Assert.That(embeds.Count, Is.EqualTo(2));
        foreach (var embed in embeds)
            Assert.That(drawingsPart.GetPartById(embed!), Is.InstanceOf<ImagePart>());
    }

    [Test]
    public void ResizingGroupedPictureRoundTrips()
    {
        using var output = new MemoryStream();
        int newWidth, newHeight;

        using (var stream = OpenFixture())
        using (var wb = new XLWorkbook(stream))
        {
            var picture1 = (XLPicture)wb.Worksheet("Map").Pictures.Single(p => p.Name == "Picture 1");
            newWidth = picture1.Width + 150;
            newHeight = picture1.Height + 90;
            picture1.Width = newWidth;
            picture1.Height = newHeight;
            wb.SaveAs(output);
        }

        output.Position = 0;
        using (var wb = new XLWorkbook(output))
        {
            var picture1 = (XLPicture)wb.Worksheet("Map").Pictures.Single(p => p.Name == "Picture 1");

            // Round-trips through the group transform involve EMU<->pixel conversions, so allow a
            // small rounding tolerance.
            Assert.That(picture1.Width, Is.EqualTo(newWidth).Within(2));
            Assert.That(picture1.Height, Is.EqualTo(newHeight).Within(2));
        }

        // The group, the second picture and the connector all survive the edit.
        output.Position = 0;
        using (var package = SpreadsheetDocument.Open(output, false))
        {
            var drawing = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing;
            var groups = drawing.Descendants<Xdr.GroupShape>().ToList();
            Assert.That(groups.Count, Is.EqualTo(1));

            // The picture stays inside the group after the resize, alongside its sibling and connector.
            var group = groups[0];
            Assert.That(group.Descendants<Xdr.Picture>().Count(), Is.EqualTo(2));
            Assert.That(group.Descendants<Xdr.ConnectionShape>().Count(), Is.EqualTo(1));
        }
    }

    // The nested fixture's "Map" sheet has an outer group (2× scale) containing Picture 1
    // (child ext 2_000_000) and an inner group (a further 2× scale) containing Picture 2
    // (child ext 500_000) and a connector. So Picture 1's sheet extent is 2_000_000 × 2 =
    // 4_000_000 EMU, and Picture 2's is 500_000 × 2 × 2 = 2_000_000 EMU.
    private const string NestedGroupPicturesResource = @"Other\Drawings\NestedGroupPictures.xlsx";

    private static MemoryStream OpenNestedFixture()
    {
        var ms = new MemoryStream();
        using var src = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(NestedGroupPicturesResource));
        src.CopyTo(ms);
        ms.Position = 0;
        return ms;
    }

    [Test]
    public void NestedGroupPicturesLoadWithComposedScale()
    {
        using var stream = OpenNestedFixture();
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheet("Map");

        Assert.That(ws.Pictures.Count, Is.EqualTo(2), "pictures at both nesting levels are loaded");

        var picture1 = (XLPicture)ws.Pictures.Single(p => p.Name == "Picture 1");
        var picture2 = (XLPicture)ws.Pictures.Single(p => p.Name == "Picture 2");
        Assert.That(picture1.IsInGroup, Is.True);
        Assert.That(picture2.IsInGroup, Is.True);

        // Composed scale: Picture 1 → 4_000_000 EMU, Picture 2 → 2_000_000 EMU (exactly 2:1).
        Assert.That(picture1.Width, Is.EqualTo(DrawingPartReader.ConvertFromEnglishMetricUnits(4_000_000, wb.DpiX)));
        Assert.That(picture2.Width, Is.EqualTo(DrawingPartReader.ConvertFromEnglishMetricUnits(2_000_000, wb.DpiX)));
    }

    [Test]
    public void UneditedNestedRoundTripPreservesStructure()
    {
        using var output = new MemoryStream();
        using (var stream = OpenNestedFixture())
        using (var wb = new XLWorkbook(stream))
        {
            wb.SaveAs(output);
        }

        output.Position = 0;
        using var package = SpreadsheetDocument.Open(output, false);
        var drawing = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing;

        Assert.That(drawing.Descendants<Xdr.GroupShape>().Count(), Is.EqualTo(2), "outer + inner group preserved");
        Assert.That(drawing.Descendants<Xdr.Picture>().Count(), Is.EqualTo(2), "both pictures preserved");
        Assert.That(drawing.Descendants<Xdr.ConnectionShape>().Count(), Is.EqualTo(1), "nested connector preserved");

        // Unedited pictures keep their exact child-space extents at their respective depths.
        var extents = drawing.Descendants<Xdr.Picture>()
            .Select(p => p.ShapeProperties!.Transform2D!.Extents!.Cx!.Value)
            .OrderByDescending(cx => cx)
            .ToList();
        Assert.That(extents, Is.EqualTo(new[] { 2_000_000L, 500_000L }));
    }

    [Test]
    public void ResizingDeeplyNestedPictureRoundTrips()
    {
        using var output = new MemoryStream();
        int newWidth, newHeight;

        using (var stream = OpenNestedFixture())
        using (var wb = new XLWorkbook(stream))
        {
            var picture2 = (XLPicture)wb.Worksheet("Map").Pictures.Single(p => p.Name == "Picture 2");
            newWidth = picture2.Width + 120;
            newHeight = picture2.Height + 120;
            picture2.Width = newWidth;
            picture2.Height = newHeight;
            wb.SaveAs(output);
        }

        output.Position = 0;
        using (var wb = new XLWorkbook(output))
        {
            var picture2 = (XLPicture)wb.Worksheet("Map").Pictures.Single(p => p.Name == "Picture 2");
            Assert.That(picture2.Width, Is.EqualTo(newWidth).Within(2));
            Assert.That(picture2.Height, Is.EqualTo(newHeight).Within(2));
        }

        // Both groups, both pictures and the connector survive the deep edit.
        output.Position = 0;
        using (var package = SpreadsheetDocument.Open(output, false))
        {
            var drawing = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing;
            Assert.That(drawing.Descendants<Xdr.GroupShape>().Count(), Is.EqualTo(2));
            Assert.That(drawing.Descendants<Xdr.Picture>().Count(), Is.EqualTo(2));
            Assert.That(drawing.Descendants<Xdr.ConnectionShape>().Count(), Is.EqualTo(1));
        }
    }
}
