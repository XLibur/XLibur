using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using XLibur.Excel;
using XLibur.Tests.Utils;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Tests.Excel.Drawings;

[TestFixture]
public class GroupedPictureTests
{
    private const string GroupedPicturesResource = @"Other\Drawings\GroupedPictures.xlsx";

    // Regression test for pictures nested inside a group shape (xdr:grpSp).
    // Previously, loading a worksheet whose drawing contained a group of pictures would read
    // only the first picture and, on save, replace the whole group anchor with that single
    // regenerated picture — silently dropping the sibling pictures, connectors and shapes
    // (e.g. "Picture 2" vanished and "Picture 1" was resized). Such pictures are now skipped
    // by the loader so the original drawing XML is preserved verbatim on round-trip.

    [Test]
    public void GroupedPicturesAreNotLoadedIntoTheModel()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(GroupedPicturesResource));
        using var wb = new XLWorkbook(stream);

        var ws = wb.Worksheet("Map");

        // The two pictures live inside a group shape and are intentionally not surfaced as
        // editable pictures, because the model cannot round-trip them without data loss.
        Assert.That(ws.Pictures.Count, Is.EqualTo(0));
    }

    [Test]
    public void RoundTripPreservesGroupedPicturesAndShapes()
    {
        using var input = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(GroupedPicturesResource));
        using var output = new MemoryStream();

        using (var wb = new XLWorkbook(input))
        {
            wb.SaveAs(output);
        }

        output.Position = 0;

        using var package = SpreadsheetDocument.Open(output, false);
        var drawingsPart = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart;
        Assert.That(drawingsPart, Is.Not.Null);

        var drawing = drawingsPart!.WorksheetDrawing;

        // The group and both of its pictures survive...
        Assert.That(drawing.Descendants<Xdr.GroupShape>().Count(), Is.EqualTo(1), "group shape should be preserved");
        Assert.That(drawing.Descendants<Xdr.Picture>().Count(), Is.EqualTo(2), "both grouped pictures should be preserved");

        // ...along with the connector inside the group...
        Assert.That(drawing.Descendants<Xdr.ConnectionShape>().Count(), Is.EqualTo(1), "grouped connector should be preserved");

        // ...and both image relationships still resolve to image parts.
        var embeds = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>()
            .Select(b => b.Embed?.Value)
            .Where(v => v is not null)
            .ToList();
        Assert.That(embeds.Count, Is.EqualTo(2));
        foreach (var embed in embeds)
            Assert.That(drawingsPart.GetPartById(embed!), Is.InstanceOf<ImagePart>());
    }
}
