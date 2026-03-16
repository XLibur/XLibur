using XLibur.Excel.ContentManagers;
using XLibur.Extensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using static XLibur.Excel.XLWorkbook;
using Drawing = DocumentFormat.OpenXml.Spreadsheet.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

internal static class PictureWriter
{
    internal static void WriteDrawings(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet,
        WorksheetPart worksheetPart,
        SaveContext context)
    {
        if (worksheetPart.DrawingsPart != null)
        {
            var xlPictures = (Drawings.XLPictures)xlWorksheet.Pictures;
            foreach (var removedPicture in xlPictures.Deleted)
            {
                var anchor = XLWorkbook.GetAnchorFromImageId(worksheetPart.DrawingsPart, removedPicture);
                if (anchor is not null)
                    worksheetPart.DrawingsPart.WorksheetDrawing!.RemoveChild(anchor);

                worksheetPart.DrawingsPart.DeletePart(removedPicture);
            }

            xlPictures.Deleted.Clear();
        }

        foreach (var pic in xlWorksheet.Pictures)
        {
            AddPictureAnchor(worksheetPart, pic, context);
        }

        if (xlWorksheet.Pictures.Count > 0)
            RebaseNonVisualDrawingPropertiesIds(worksheetPart);

        var tableParts = worksheet.Elements<TableParts>().First();
        if (xlWorksheet.Pictures.Count > 0 && !worksheet.OfType<Drawing>().Any())
        {
            var worksheetDrawing = new Drawing { Id = worksheetPart.GetIdOfPart(worksheetPart.DrawingsPart!) };
            worksheetDrawing.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet.InsertBefore(worksheetDrawing, tableParts);
            cm.SetElement(XLWorksheetContents.Drawing, worksheet.Elements<Drawing>().First());
        }

        // Instead of saving a file with an empty Drawings.xml file, rather remove the .xml file
        var hasCharts = worksheetPart.DrawingsPart is not null && worksheetPart.DrawingsPart.Parts.Any();
        if (worksheetPart.DrawingsPart is not null && // There is a drawing part for the sheet that could be deleted
            xlWorksheet
                .LegacyDrawingId is null && // and sheet doesn't contain any form controls or comments or other shapes
            xlWorksheet.Pictures.Count == 0 && // and also no pictures.
            !hasCharts && // and no charts
                          // Check for non-picture shapes (textboxes, rectangles, etc.) last to avoid
                          // loading the DrawingsPart DOM unnecessarily — DOM loading causes re-serialization
                          // that changes the XML even when no modifications are made.
            !(worksheetPart.DrawingsPart.WorksheetDrawing?.HasChildren ?? false))
        {
            var id = worksheetPart.GetIdOfPart(worksheetPart.DrawingsPart);
            worksheet.RemoveChild(worksheet.OfType<Drawing>().FirstOrDefault(p => p.Id == id));
            worksheetPart.DeletePart(worksheetPart.DrawingsPart);
            cm.SetElement(XLWorksheetContents.Drawing, null);
        }
    }

    internal static void WriteLegacyDrawing(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet)
    {
        // Does worksheet have any comments (stored in legacy VML drawing)
        if (!string.IsNullOrEmpty(xlWorksheet.LegacyDrawingId))
        {
            worksheet.RemoveAllChildren<LegacyDrawing>();
            var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.LegacyDrawing);
            worksheet.InsertAfter(new LegacyDrawing { Id = xlWorksheet.LegacyDrawingId },
                previousElement);

            cm.SetElement(XLWorksheetContents.LegacyDrawing, worksheet.Elements<LegacyDrawing>().First());
        }
        else
        {
            worksheet.RemoveAllChildren<LegacyDrawing>();
            cm.SetElement(XLWorksheetContents.LegacyDrawing, null);
        }
    }

    // http://polymathprogrammer.com/2009/10/22/english-metric-units-and-open-xml/
    // http://archive.oreilly.com/pub/post/what_is_an_emu.html
    // https://en.wikipedia.org/wiki/Office_Open_XML_file_formats#DrawingML
    private static long ConvertToEnglishMetricUnits(int pixels, double resolution)
    {
        return Convert.ToInt64(914400L * pixels / resolution);
    }

    private static void AddPictureAnchor(WorksheetPart worksheetPart, Drawings.IXLPicture picture, SaveContext context)
    {
        var pic = (Drawings.XLPicture)picture;
        var drawingsPart = worksheetPart.DrawingsPart ??
                           worksheetPart.AddNewPart<DrawingsPart>(context.RelIdGenerator.GetNext(RelType.Workbook));

        drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

        var worksheetDrawing = drawingsPart.WorksheetDrawing;

        // Add namespaces
        if (!worksheetDrawing.NamespaceDeclarations.Any(nd =>
                nd.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/main")))
            worksheetDrawing.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

        if (!worksheetDrawing.NamespaceDeclarations.Any(nd =>
                nd.Value.Equals("http://schemas.openxmlformats.org/officeDocument/2006/relationships")))
            worksheetDrawing.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        // Overwrite actual image binary data
        ImagePart imagePart;
        if (drawingsPart.HasPartWithId(pic.RelId!))
            imagePart = (ImagePart)drawingsPart.GetPartById(pic.RelId!);
        else
        {
            pic.RelId = context.RelIdGenerator.GetNext(RelType.Workbook);
            imagePart = drawingsPart.AddImagePart(pic.Format.ToOpenXml(), pic.RelId);
        }

        pic.ImageStream.Position = 0;
        imagePart.FeedData(pic.ImageStream);

        // Clear current anchors
        var existingAnchor = XLWorkbook.GetAnchorFromImageId(drawingsPart, pic.RelId!);

        var wb = pic.Worksheet.Workbook;
        var extentsCx = ConvertToEnglishMetricUnits(pic.Width, wb.DpiX);
        var extentsCy = ConvertToEnglishMetricUnits(pic.Height, wb.DpiY);

        var nvps = worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>();
        var nvpId = nvps.Any()
            ? (UInt32Value)worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>().Max(p => p.Id!.Value) + 1
            : 1U;

        Xdr.FromMarker fMark;
        switch (pic.Placement)
        {
            case Drawings.XLPicturePlacement.FreeFloating:
                var absoluteAnchor = new Xdr.AbsoluteAnchor(
                    new Xdr.Position
                    {
                        X = ConvertToEnglishMetricUnits(pic.Left, wb.DpiX),
                        Y = ConvertToEnglishMetricUnits(pic.Top, wb.DpiY)
                    },
                    new Xdr.Extent
                    {
                        Cx = extentsCx,
                        Cy = extentsCy
                    },
                    new Xdr.Picture(
                        new Xdr.NonVisualPictureProperties(
                            new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = pic.Name },
                            new Xdr.NonVisualPictureDrawingProperties(new PictureLocks { NoChangeAspect = true })
                        ),
                        new Xdr.BlipFill(
                            new Blip
                            {
                                Embed = drawingsPart.GetIdOfPart(imagePart),
                                CompressionState = BlipCompressionValues.Print
                            },
                            new Stretch(new FillRectangle())
                        ),
                        new Xdr.ShapeProperties(
                            new Transform2D(
                                new Offset { X = 0, Y = 0 },
                                new Extents { Cx = extentsCx, Cy = extentsCy }
                            ),
                            new PresetGeometry { Preset = ShapeTypeValues.Rectangle }
                        )
                    ),
                    new Xdr.ClientData()
                );

                AttachAnchor(absoluteAnchor, existingAnchor);
                break;

            case Drawings.XLPicturePlacement.MoveAndSize:
                var moveAndSizeFromMarker = pic.Markers[Drawings.XLMarkerPosition.TopLeft];
                if (moveAndSizeFromMarker == null)
                    moveAndSizeFromMarker = new Drawings.XLMarker(picture.Worksheet.Cell("A1"));
                fMark = new Xdr.FromMarker
                {
                    ColumnId = new Xdr.ColumnId((moveAndSizeFromMarker.ColumnNumber - 1).ToInvariantString()),
                    RowId = new Xdr.RowId((moveAndSizeFromMarker.RowNumber - 1).ToInvariantString()),
                    ColumnOffset =
                        new Xdr.ColumnOffset(ConvertToEnglishMetricUnits(moveAndSizeFromMarker.Offset.X, wb.DpiX)
                            .ToInvariantString()),
                    RowOffset = new Xdr.RowOffset(ConvertToEnglishMetricUnits(moveAndSizeFromMarker.Offset.Y, wb.DpiY)
                        .ToInvariantString())
                };

                var moveAndSizeToMarker = pic.Markers[Drawings.XLMarkerPosition.BottomRight];
                if (moveAndSizeToMarker == null)
                    moveAndSizeToMarker = new Drawings.XLMarker(picture.Worksheet.Cell("A1"),
                        new System.Drawing.Point(picture.Width, picture.Height));
                var tMark = new Xdr.ToMarker
                {
                    ColumnId = new Xdr.ColumnId((moveAndSizeToMarker.ColumnNumber - 1).ToInvariantString()),
                    RowId = new Xdr.RowId((moveAndSizeToMarker.RowNumber - 1).ToInvariantString()),
                    ColumnOffset =
                        new Xdr.ColumnOffset(ConvertToEnglishMetricUnits(moveAndSizeToMarker.Offset.X, wb.DpiX)
                            .ToInvariantString()),
                    RowOffset = new Xdr.RowOffset(ConvertToEnglishMetricUnits(moveAndSizeToMarker.Offset.Y, wb.DpiY)
                        .ToInvariantString())
                };

                var twoCellAnchor = new Xdr.TwoCellAnchor(
                    fMark,
                    tMark,
                    new Xdr.Picture(
                        new Xdr.NonVisualPictureProperties(
                            new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = pic.Name },
                            new Xdr.NonVisualPictureDrawingProperties(new PictureLocks { NoChangeAspect = true })
                        ),
                        new Xdr.BlipFill(
                            new Blip
                            {
                                Embed = drawingsPart.GetIdOfPart(imagePart),
                                CompressionState = BlipCompressionValues.Print
                            },
                            new Stretch(new FillRectangle())
                        ),
                        new Xdr.ShapeProperties(
                            new Transform2D(
                                new Offset { X = 0, Y = 0 },
                                new Extents { Cx = extentsCx, Cy = extentsCy }
                            ),
                            new PresetGeometry { Preset = ShapeTypeValues.Rectangle }
                        )
                    ),
                    new Xdr.ClientData()
                );

                AttachAnchor(twoCellAnchor, existingAnchor);
                break;

            case Drawings.XLPicturePlacement.Move:
                var moveFromMarker = pic.Markers[Drawings.XLMarkerPosition.TopLeft];
                if (moveFromMarker == null) moveFromMarker = new Drawings.XLMarker(picture.Worksheet.Cell("A1"));
                fMark = new Xdr.FromMarker
                {
                    ColumnId = new Xdr.ColumnId((moveFromMarker.ColumnNumber - 1).ToInvariantString()),
                    RowId = new Xdr.RowId((moveFromMarker.RowNumber - 1).ToInvariantString()),
                    ColumnOffset = new Xdr.ColumnOffset(ConvertToEnglishMetricUnits(moveFromMarker.Offset.X, wb.DpiX)
                        .ToInvariantString()),
                    RowOffset = new Xdr.RowOffset(ConvertToEnglishMetricUnits(moveFromMarker.Offset.Y, wb.DpiY)
                        .ToInvariantString())
                };

                var oneCellAnchor = new Xdr.OneCellAnchor(
                    fMark,
                    new Xdr.Extent
                    {
                        Cx = extentsCx,
                        Cy = extentsCy
                    },
                    new Xdr.Picture(
                        new Xdr.NonVisualPictureProperties(
                            new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = pic.Name },
                            new Xdr.NonVisualPictureDrawingProperties(new PictureLocks { NoChangeAspect = true })
                        ),
                        new Xdr.BlipFill(
                            new Blip
                            {
                                Embed = drawingsPart.GetIdOfPart(imagePart),
                                CompressionState = BlipCompressionValues.Print
                            },
                            new Stretch(new FillRectangle())
                        ),
                        new Xdr.ShapeProperties(
                            new Transform2D(
                                new Offset { X = 0, Y = 0 },
                                new Extents { Cx = extentsCx, Cy = extentsCy }
                            ),
                            new PresetGeometry { Preset = ShapeTypeValues.Rectangle }
                        )
                    ),
                    new Xdr.ClientData()
                );

                AttachAnchor(oneCellAnchor, existingAnchor);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(picture), pic.Placement, "Unsupported picture placement.");
        }

        return;

        void AttachAnchor(OpenXmlElement pictureAnchor, OpenXmlElement? existingAnchorX)
        {
            if (existingAnchorX is not null)
            {
                worksheetDrawing.ReplaceChild(pictureAnchor, existingAnchorX);
            }
            else
            {
                worksheetDrawing.Append(pictureAnchor);
            }
        }
    }

    private static void RebaseNonVisualDrawingPropertiesIds(WorksheetPart worksheetPart)
    {
        var worksheetDrawing = worksheetPart.DrawingsPart!.WorksheetDrawing;

        uint id = 1;
        foreach (var nvdpr in worksheetDrawing!.Descendants<Xdr.NonVisualDrawingProperties>())
            nvdpr.Id = id++;
    }
}
