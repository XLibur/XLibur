using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using XLibur.Excel.Drawings;
using XLibur.Excel.IO;
using XLibur.Tests.Utils;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;

namespace XLibur.Tests.Excel.Drawings;

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
    public async Task GroupedPicturesAreLoadedWithGroupScaledGeometry()
    {
        using var stream = OpenFixture();
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheet("Map");

        await Assert.That(ws.Pictures.Count).IsEqualTo(2);

        var picture1 = (XLPicture)ws.Pictures.Single(p => p.Name == "Picture 1");
        var picture2 = (XLPicture)ws.Pictures.Single(p => p.Name == "Picture 2");

        await Assert.That(picture1.IsInGroup).IsTrue();
        await Assert.That(picture2.IsInGroup).IsTrue();

        // Both pictures are scaled by the same group factor, so their relative sizes are preserved:
        // Picture 1 (child 2_000_000) is larger than Picture 2 (child 1_500_000).
        await Assert.That(picture1.Width).IsGreaterThan(0);
        await Assert.That(picture1.Width).IsGreaterThan(picture2.Width);
        await Assert.That(picture1.Height).IsGreaterThan(picture2.Height);

        // Picture 1's sheet-space extent is twice its child extent (2_000_000 → 4_000_000 EMU).
        var expectedPx1 = DrawingPartReader.ConvertFromEnglishMetricUnits(4_000_000, wb.DpiX);
        await Assert.That(picture1.Width).IsEqualTo(expectedPx1);
    }

    [Test]
    public async Task UneditedRoundTripPreservesGroupPicturesAndShapes()
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
        await Assert.That(drawingsPart).IsNotNull();
        var drawing = drawingsPart!.WorksheetDrawing;

        var groups = drawing.Descendants<Xdr.GroupShape>().ToList();
        await Assert.That(groups.Count).IsEqualTo(1).Because("group preserved");

        // Assert on the group node so the shapes are verified to remain *inside* the group rather
        // than having been moved out to the top level during the round-trip.
        var group = groups[0];
        await Assert.That(group.Descendants<Xdr.Picture>().Count()).IsEqualTo(2).Because("both pictures preserved inside the group");
        await Assert.That(group.Descendants<Xdr.ConnectionShape>().Count()).IsEqualTo(1).Because("connector preserved inside the group");

        // An unedited grouped picture must keep its exact child-space extent (no rounding drift).
        var extents = group.Descendants<Xdr.Picture>()
            .Select(p => p.ShapeProperties!.Transform2D!.Extents!)
            .Select(e => (e.Cx!.Value, e.Cy!.Value))
            .OrderByDescending(t => t.Item1)
            .ToList();
        await Assert.That(extents[0]).IsEqualTo((2_000_000L, 2_000_000L));
        await Assert.That(extents[1]).IsEqualTo((1_500_000L, 1_500_000L));

        // Both image relationships still resolve to image parts.
        var embeds = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>()
            .Select(b => b.Embed?.Value).Where(v => v is not null).ToList();
        await Assert.That(embeds.Count).IsEqualTo(2);
        foreach (var embed in embeds)
            await Assert.That(drawingsPart.GetPartById(embed!)).IsAssignableTo<ImagePart>();
    }

    [Test]
    public async Task ResizingGroupedPictureRoundTrips()
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
            await Assert.That(picture1.Width).IsEqualTo(newWidth).Within(2);
            await Assert.That(picture1.Height).IsEqualTo(newHeight).Within(2);
        }

        // The group, the second picture and the connector all survive the edit.
        output.Position = 0;
        using (var package = SpreadsheetDocument.Open(output, false))
        {
            var drawing = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing;
            var groups = drawing.Descendants<Xdr.GroupShape>().ToList();
            await Assert.That(groups.Count).IsEqualTo(1);

            // The picture stays inside the group after the resize, alongside its sibling and connector.
            var group = groups[0];
            await Assert.That(group.Descendants<Xdr.Picture>().Count()).IsEqualTo(2);
            await Assert.That(group.Descendants<Xdr.ConnectionShape>().Count()).IsEqualTo(1);
        }
    }

    [Test]
    public async Task GroupedPictureLeftTopReflectSheetPosition()
    {
        using var stream = OpenFixture();
        using var wb = new XLWorkbook(stream);

        var picture1 = (XLPicture)wb.Worksheet("Map").Pictures.Single(p => p.Name == "Picture 1");

        // Group off (1_000_000, 1_000_000), 2× scale, child off (1_000_000, 1_000_000):
        // sheet pos = (off − chOff·scale) + childOff·scale = −1_000_000 + 1_000_000·2 = 1_000_000 EMU.
        await Assert.That(picture1.Left).IsEqualTo(DrawingPartReader.ConvertFromEnglishMetricUnits(1_000_000, wb.DpiX));
        await Assert.That(picture1.Top).IsEqualTo(DrawingPartReader.ConvertFromEnglishMetricUnits(1_000_000, wb.DpiY));
    }

    [Test]
    public async Task MovingGroupedPictureRoundTrips()
    {
        using var output = new MemoryStream();
        int newLeft, newTop;

        using (var stream = OpenFixture())
        using (var wb = new XLWorkbook(stream))
        {
            var picture1 = (XLPicture)wb.Worksheet("Map").Pictures.Single(p => p.Name == "Picture 1");
            newLeft = picture1.Left + 200;
            newTop = picture1.Top + 150;
            picture1.Left = newLeft;
            picture1.Top = newTop;
            wb.SaveAs(output);
        }

        output.Position = 0;
        using (var wb = new XLWorkbook(output))
        {
            var picture1 = (XLPicture)wb.Worksheet("Map").Pictures.Single(p => p.Name == "Picture 1");
            await Assert.That(picture1.Left).IsEqualTo(newLeft).Within(2);
            await Assert.That(picture1.Top).IsEqualTo(newTop).Within(2);
        }

        // The picture stays inside the group, and the sibling + connector are untouched.
        output.Position = 0;
        using (var package = SpreadsheetDocument.Open(output, false))
        {
            var group = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing
                .Descendants<Xdr.GroupShape>().Single();
            await Assert.That(group.Descendants<Xdr.Picture>().Count()).IsEqualTo(2);
            await Assert.That(group.Descendants<Xdr.ConnectionShape>().Count()).IsEqualTo(1);
        }
    }

    [Test]
    public async Task RemovingGroupedPictureKeepsTheRestOfTheGroup()
    {
        using var output = new MemoryStream();
        using (var stream = OpenFixture())
        using (var wb = new XLWorkbook(stream))
        {
            var ws = wb.Worksheet("Map");
            ws.Pictures.Single(p => p.Name == "Picture 2").Delete();
            await Assert.That(ws.Pictures.Count).IsEqualTo(1).Because("deleted picture removed from the collection");
            wb.SaveAs(output);
        }

        // Reopening: only Picture 1 remains, still inside the group.
        output.Position = 0;
        using (var wb = new XLWorkbook(output))
        {
            var pictures = wb.Worksheet("Map").Pictures;
            await Assert.That(pictures.Count).IsEqualTo(1);
            await Assert.That(pictures.Single().Name).IsEqualTo("Picture 1");
        }

        // Only the deleted xdr:pic is gone; the group, the surviving picture and the connector stay.
        output.Position = 0;
        using (var package = SpreadsheetDocument.Open(output, false))
        {
            var drawingsPart = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart!;
            var drawing = drawingsPart.WorksheetDrawing;
            var group = drawing.Descendants<Xdr.GroupShape>().Single();

            await Assert.That(group.Descendants<Xdr.Picture>().Count()).IsEqualTo(1);
            await Assert.That(group.Descendants<Xdr.ConnectionShape>().Count()).IsEqualTo(1);

            // The surviving picture's image part is kept; the removed picture's is dropped.
            var embeds = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>()
                .Select(b => b.Embed?.Value).Where(v => v is not null).ToList();
            await Assert.That(embeds.Count).IsEqualTo(1);
            await Assert.That(drawingsPart.GetPartById(embeds[0]!)).IsAssignableTo<ImagePart>();
            await Assert.That(drawingsPart.Parts.Count(p => p.OpenXmlPart is ImagePart)).IsEqualTo(1);
        }
    }

    [Test]
    public async Task AddingPictureToGroupRoundTrips()
    {
        using var output = new MemoryStream();
        int width, height, left, top;

        using (var stream = OpenFixture())
        using (var wb = new XLWorkbook(stream))
        {
            var ws = wb.Worksheet("Map");
            var pictures = (XLPictures)ws.Pictures;
            var sibling = (XLPicture)ws.Pictures.Single(p => p.Name == "Picture 1");
            using var image = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Images\SampleImagePng.png"));

            var added = pictures.AddToGroup(sibling, image, "Added Picture");
            added.Width = 300;
            added.Height = 200;
            added.Left = 400;
            added.Top = 250;
            (width, height, left, top) = (added.Width, added.Height, added.Left, added.Top);

            await Assert.That(pictures.Count).IsEqualTo(3);
            wb.SaveAs(output);
        }

        output.Position = 0;
        using (var wb = new XLWorkbook(output))
        {
            var pictures = wb.Worksheet("Map").Pictures;
            await Assert.That(pictures.Count).IsEqualTo(3);

            var added = (XLPicture)pictures.Single(p => p.Name == "Added Picture");
            await Assert.That(added.IsInGroup).IsTrue();
            await Assert.That(added.Width).IsEqualTo(width).Within(2);
            await Assert.That(added.Height).IsEqualTo(height).Within(2);
            await Assert.That(added.Left).IsEqualTo(left).Within(2);
            await Assert.That(added.Top).IsEqualTo(top).Within(2);
        }

        // The new picture went inside the group with the two originals and the connector, and got its
        // own image part.
        output.Position = 0;
        using (var package = SpreadsheetDocument.Open(output, false))
        {
            var drawingsPart = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart!;
            var group = drawingsPart.WorksheetDrawing.Descendants<Xdr.GroupShape>().Single();
            await Assert.That(group.Descendants<Xdr.Picture>().Count()).IsEqualTo(3);
            await Assert.That(group.Descendants<Xdr.ConnectionShape>().Count()).IsEqualTo(1);
            await Assert.That(drawingsPart.Parts.Count(p => p.OpenXmlPart is ImagePart)).IsEqualTo(3);
        }
    }

    private static void AddFreeFloatingPicture(IXLWorksheet ws, string name, int left, int top)
    {
        using var image = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Images\SampleImagePng.png"));
        ws.AddPicture(image, name).MoveTo(left, top);
    }

    [Test]
    public async Task GroupingFreeFloatingPicturesCreatesAGroup()
    {
        // Build a workbook with two free-floating pictures and save, so they exist in the drawing.
        using var seeded = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Map");
            AddFreeFloatingPicture(ws, "Pic A", 100, 100);
            AddFreeFloatingPicture(ws, "Pic B", 500, 300);
            wb.SaveAs(seeded);
        }

        using var output = new MemoryStream();
        seeded.Position = 0;
        using (var wb = new XLWorkbook(seeded))
        {
            var ws = wb.Worksheet("Map");
            var pictures = (XLPictures)ws.Pictures;
            var a = (XLPicture)ws.Pictures.Single(p => p.Name == "Pic A");
            var b = (XLPicture)ws.Pictures.Single(p => p.Name == "Pic B");
            pictures.Group(a, b);
            wb.SaveAs(output);
        }

        // Both pictures now live in a single group; no top-level picture anchors remain.
        output.Position = 0;
        using (var package = SpreadsheetDocument.Open(output, false))
        {
            var drawing = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing;
            var group = drawing.Descendants<Xdr.GroupShape>().Single();
            await Assert.That(group.Descendants<Xdr.Picture>().Count()).IsEqualTo(2);
            await Assert.That(drawing.Descendants<Xdr.Picture>().Count()).IsEqualTo(2).Because("no picture left outside the group");
            await Assert.That(drawing.Elements<Xdr.AbsoluteAnchor>().Count()).IsEqualTo(1).Because("only the group's anchor remains");
        }

        // XLibur reloads them as grouped pictures with their positions preserved.
        output.Position = 0;
        using (var wb = new XLWorkbook(output))
        {
            var ws = wb.Worksheet("Map");
            await Assert.That(ws.Pictures.Count).IsEqualTo(2);
            var a = (XLPicture)ws.Pictures.Single(p => p.Name == "Pic A");
            await Assert.That(a.IsInGroup).IsTrue();
            await Assert.That(a.Left).IsEqualTo(100).Within(2);
            await Assert.That(a.Top).IsEqualTo(100).Within(2);
        }
    }

    [Test]
    public async Task PublicGroupApiExposesMembershipAndMutation()
    {
        using var output = new MemoryStream();
        using (var stream = OpenFixture())
        using (var wb = new XLWorkbook(stream))
        {
            var ws = wb.Worksheet("Map");

            // IXLWorksheet.PictureGroups and IXLPicture.Group expose the group.
            await Assert.That(ws.PictureGroups.Count()).IsEqualTo(1);
            var picture1 = ws.Pictures.Single(p => p.Name == "Picture 1");
            await Assert.That(picture1.IsInGroup).IsTrue();

            var group = picture1.Group;
            await Assert.That(group).IsNotNull();
            await Assert.That(group!.Worksheet).IsSameReferenceAs(ws);
            await Assert.That(group.Pictures.Count()).IsEqualTo(2);

            // IXLPictureGroup.Add and Remove mutate membership.
            using var image = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Images\SampleImagePng.png"));
            var added = group.Add(image, "Group Added");
            added.Width = 200;
            added.Height = 150;
            added.Left = 50;
            added.Top = 60;
            await Assert.That(group.Pictures.Count()).IsEqualTo(3);

            group.Remove(ws.Pictures.Single(p => p.Name == "Picture 2"));
            await Assert.That(group.Pictures.Count()).IsEqualTo(2);

            wb.SaveAs(output);
        }

        output.Position = 0;
        using (var wb = new XLWorkbook(output))
        {
            var ws = wb.Worksheet("Map");
            await Assert.That(ws.PictureGroups.Count()).IsEqualTo(1);
            var group = ws.PictureGroups.Single();
            var names = group.Pictures.Select(p => p.Name).ToList();
            await Assert.That(names).IsEquivalentTo(new[] { "Picture 1", "Group Added" });
        }
    }

    [Test]
    public async Task AddingToANewlyCreatedGroupBeforeSaveIsNotDropped()
    {
        using var seeded = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Map");
            AddFreeFloatingPicture(ws, "Pic A", 100, 100);
            AddFreeFloatingPicture(ws, "Pic B", 500, 300);
            wb.SaveAs(seeded);
        }

        using var output = new MemoryStream();
        seeded.Position = 0;
        using (var wb = new XLWorkbook(seeded))
        {
            var ws = wb.Worksheet("Map");
            var a = (XLPicture)ws.Pictures.Single(p => p.Name == "Pic A");
            var b = (XLPicture)ws.Pictures.Single(p => p.Name == "Pic B");

            // Create the group and add a third picture to it, all before any save: the group's
            // drawing id isn't assigned until save, so the added picture inherits a null id.
            var group = ws.Pictures.Group(a, b);
            using var image = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Images\SampleImagePng.png"));
            var added = group.Add(image, "Pic C");
            added.Width = 150;
            added.Height = 120;
            added.Left = 700;
            added.Top = 200;

            wb.SaveAs(output);
        }

        output.Position = 0;
        using (var package = SpreadsheetDocument.Open(output, false))
        {
            var group = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing
                .Descendants<Xdr.GroupShape>().Single();
            await Assert.That(group.Descendants<Xdr.Picture>().Count()).IsEqualTo(3).Because("the picture added to the group before its first save must not be dropped");
        }
    }

    [Test]
    public async Task GroupRejectsInvalidArguments()
    {
        using var seeded = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Map");
            AddFreeFloatingPicture(ws, "Pic A", 100, 100);
            wb.SaveAs(seeded);
        }

        seeded.Position = 0;
        using var reopened = new XLWorkbook(seeded);
        var pictures = reopened.Worksheet("Map").Pictures;
        var a = pictures.Single(p => p.Name == "Pic A");

        await Assert.That(() => pictures.Group(null!)).Throws<ArgumentNullException>();
        await Assert.That(() => pictures.Group(a, null!)).Throws<ArgumentException>();
        await Assert.That(() => pictures.Group(a, a)).Throws<ArgumentException>();
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
    public async Task NestedGroupPicturesLoadWithComposedScale()
    {
        using var stream = OpenNestedFixture();
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheet("Map");

        await Assert.That(ws.Pictures.Count).IsEqualTo(2).Because("pictures at both nesting levels are loaded");

        var picture1 = (XLPicture)ws.Pictures.Single(p => p.Name == "Picture 1");
        var picture2 = (XLPicture)ws.Pictures.Single(p => p.Name == "Picture 2");
        await Assert.That(picture1.IsInGroup).IsTrue();
        await Assert.That(picture2.IsInGroup).IsTrue();

        // Composed scale: Picture 1 → 4_000_000 EMU, Picture 2 → 2_000_000 EMU (exactly 2:1).
        await Assert.That(picture1.Width).IsEqualTo(DrawingPartReader.ConvertFromEnglishMetricUnits(4_000_000, wb.DpiX));
        await Assert.That(picture2.Width).IsEqualTo(DrawingPartReader.ConvertFromEnglishMetricUnits(2_000_000, wb.DpiX));
    }

    [Test]
    public async Task UneditedNestedRoundTripPreservesStructure()
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

        await Assert.That(drawing.Descendants<Xdr.GroupShape>().Count()).IsEqualTo(2).Because("outer + inner group preserved");
        await Assert.That(drawing.Descendants<Xdr.Picture>().Count()).IsEqualTo(2).Because("both pictures preserved");
        await Assert.That(drawing.Descendants<Xdr.ConnectionShape>().Count()).IsEqualTo(1).Because("nested connector preserved");

        // Unedited pictures keep their exact child-space extents at their respective depths.
        var extents = drawing.Descendants<Xdr.Picture>()
            .Select(p => p.ShapeProperties!.Transform2D!.Extents!.Cx!.Value)
            .OrderByDescending(cx => cx)
            .ToList();
        await Assert.That(extents).IsEquivalentTo(new[] { 2_000_000L, 500_000L }, CollectionOrdering.Matching);
    }

    [Test]
    public async Task ResizingDeeplyNestedPictureRoundTrips()
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
            await Assert.That(picture2.Width).IsEqualTo(newWidth).Within(2);
            await Assert.That(picture2.Height).IsEqualTo(newHeight).Within(2);
        }

        // Both groups, both pictures and the connector survive the deep edit.
        output.Position = 0;
        using (var package = SpreadsheetDocument.Open(output, false))
        {
            var drawing = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing;
            await Assert.That(drawing.Descendants<Xdr.GroupShape>().Count()).IsEqualTo(2);
            await Assert.That(drawing.Descendants<Xdr.Picture>().Count()).IsEqualTo(2);
            await Assert.That(drawing.Descendants<Xdr.ConnectionShape>().Count()).IsEqualTo(1);
        }
    }

    [Test]
    public async Task MovingDeeplyNestedPictureRoundTrips()
    {
        using var output = new MemoryStream();
        int newLeft, newTop;

        using (var stream = OpenNestedFixture())
        using (var wb = new XLWorkbook(stream))
        {
            var picture2 = (XLPicture)wb.Worksheet("Map").Pictures.Single(p => p.Name == "Picture 2");
            newLeft = picture2.Left + 100;
            newTop = picture2.Top + 80;
            picture2.Left = newLeft;
            picture2.Top = newTop;
            wb.SaveAs(output);
        }

        output.Position = 0;
        using (var wb = new XLWorkbook(output))
        {
            var picture2 = (XLPicture)wb.Worksheet("Map").Pictures.Single(p => p.Name == "Picture 2");
            await Assert.That(picture2.Left).IsEqualTo(newLeft).Within(2);
            await Assert.That(picture2.Top).IsEqualTo(newTop).Within(2);
        }

        output.Position = 0;
        using (var package = SpreadsheetDocument.Open(output, false))
        {
            var drawing = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing;
            await Assert.That(drawing.Descendants<Xdr.GroupShape>().Count()).IsEqualTo(2);
            await Assert.That(drawing.Descendants<Xdr.Picture>().Count()).IsEqualTo(2);
            await Assert.That(drawing.Descendants<Xdr.ConnectionShape>().Count()).IsEqualTo(1);
        }
    }
}
