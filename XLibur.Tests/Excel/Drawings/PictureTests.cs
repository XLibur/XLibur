using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;
using XLibur.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Tests.Excel.Drawings;

[TestFixture]
public class PictureTests
{
    [TestCase("Other.Drawings.picture-webp.xlsx")]
    public void Can_load_and_save_workbook_with_image_type(string resourceWithImageType)
    {
        TestHelper.LoadSaveAndCompare(resourceWithImageType, resourceWithImageType);
    }

    [Test]
    public void Can_load_picture_with_empty_name()
    {
        // Empty name attribute on cNvPr is valid per ECMA-376 (xsd:string, no minLength).
        // Excel can produce such files. Verify they load without throwing.
        using var xlsxStream = CreateXlsxWithEmptyPictureName();
        using var wb = new XLWorkbook(xlsxStream);

        var ws = wb.Worksheets.First();
        Assert.That(ws.Pictures.Count, Is.EqualTo(1));

        var pic = ws.Pictures.First();
        Assert.That(pic.Name, Does.StartWith("Picture"));
    }

    [Test]
    public void Non_picture_shapes_are_preserved_after_roundtrip()
    {
        // Issue #2377: textboxes and shapes (non-picture anchors) were lost after load/save
        // because the DrawingsPart was deleted when there were no pictures.
        using var stream = TestHelper.GetStreamFromResource(
            TestHelper.GetResourcePath(@"TryToLoad\textbox_shapemissing_onload_2377.xlsx"));
        using var wb = new XLWorkbook(stream);
        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        // Verify the drawing part still exists with both shapes
        ms.Position = 0;
        using var savedDoc = SpreadsheetDocument.Open(ms, false);
        var worksheetPart = savedDoc.WorkbookPart!.WorksheetParts.First();
        Assert.That(worksheetPart.DrawingsPart, Is.Not.Null, "DrawingsPart should be preserved");

        var drawing = worksheetPart.DrawingsPart!.WorksheetDrawing!;
        var twoCellAnchors = drawing.Elements<Xdr.TwoCellAnchor>().ToList();
        var oneCellAnchors = drawing.Elements<Xdr.OneCellAnchor>().ToList();

        Assert.That(twoCellAnchors, Has.Count.EqualTo(1), "TwoCellAnchor (rectangle shape) should be preserved");
        Assert.That(oneCellAnchors, Has.Count.EqualTo(1), "OneCellAnchor (textbox) should be preserved");

        // Verify the shape text content is preserved
        var shapeText = twoCellAnchors[0].Descendants<Xdr.Shape>().First()
            .Descendants<Xdr.TextBody>().First().InnerText;
        Assert.That(shapeText, Is.EqualTo("SHAPE"));

        var textboxText = oneCellAnchors[0].Descendants<Xdr.Shape>().First()
            .Descendants<Xdr.TextBody>().First().InnerText;
        Assert.That(textboxText, Is.EqualTo("TEXTBOX"));
    }

    /// <summary>
    /// Creates a minimal xlsx with a single picture whose cNvPr name attribute is empty string.
    /// </summary>
    private static MemoryStream CreateXlsxWithEmptyPictureName()
    {
        var ms = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            workbookPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" }));

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add a drawing part
            var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
            worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });

            // Add an image part with a real PNG
            using var imageStream = Assembly.GetExecutingAssembly()
                .GetManifestResourceStream("XLibur.Tests.Resource.Images.ImageHandling.png");
            var imagePart = drawingsPart.AddImagePart(ImagePartType.Png);
            imagePart.FeedData(imageStream);
            var imageRelId = drawingsPart.GetIdOfPart(imagePart);

            // Create a TwoCellAnchor with a picture that has an empty name
            var worksheetDrawing = new Xdr.WorksheetDrawing();
            worksheetDrawing.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            worksheetDrawing.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            var twoCellAnchor = new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("0"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("0"),
                    new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId("5"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("5"),
                    new Xdr.RowOffset("0")),
                new Xdr.Picture(
                    new Xdr.NonVisualPictureProperties(
                        new Xdr.NonVisualDrawingProperties { Id = 1, Name = "" },  // Empty name!
                        new Xdr.NonVisualPictureDrawingProperties(
                            new A.PictureLocks { NoChangeAspect = true })),
                    new Xdr.BlipFill(
                        new A.Blip { Embed = imageRelId },
                        new A.Stretch(new A.FillRectangle())),
                    new Xdr.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0, Y = 0 },
                            new A.Extents { Cx = 1000000, Cy = 1000000 }),
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })),
                new Xdr.ClientData());

            worksheetDrawing.Append(twoCellAnchor);
            drawingsPart.WorksheetDrawing = worksheetDrawing;
        }

        ms.Position = 0;
        return ms;
    }
}
