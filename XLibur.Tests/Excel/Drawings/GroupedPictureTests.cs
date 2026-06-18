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

        Assert.That(drawing.Descendants<Xdr.GroupShape>().Count(), Is.EqualTo(1), "group preserved");
        Assert.That(drawing.Descendants<Xdr.Picture>().Count(), Is.EqualTo(2), "both pictures preserved");
        Assert.That(drawing.Descendants<Xdr.ConnectionShape>().Count(), Is.EqualTo(1), "connector preserved");

        // An unedited grouped picture must keep its exact child-space extent (no rounding drift).
        var extents = drawing.Descendants<Xdr.Picture>()
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
            Assert.That(drawing.Descendants<Xdr.GroupShape>().Count(), Is.EqualTo(1));
            Assert.That(drawing.Descendants<Xdr.Picture>().Count(), Is.EqualTo(2));
            Assert.That(drawing.Descendants<Xdr.ConnectionShape>().Count(), Is.EqualTo(1));
        }
    }
}
