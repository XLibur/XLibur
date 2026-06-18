using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ContentManagers;
using XLibur.Excel.Drawings;
using XLibur.Extensions;
using static XLibur.Excel.XLWorkbook;
using Drawing = DocumentFormat.OpenXml.Spreadsheet.Drawing;
using Point = System.Drawing.Point;
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
            var xlPictures = (XLPictures)xlWorksheet.Pictures;
            foreach (var removedPicture in xlPictures.Deleted)
            {
                var anchor = GetAnchorFromImageId(worksheetPart.DrawingsPart, removedPicture);

                // Deleting a picture that lives in a group would otherwise dismantle the whole
                // group (and its sibling shapes). Removing a single picture from a group is not
                // supported yet, so leave the group — and its image part — intact.
                if (anchor is not null && anchor.Descendants<Xdr.GroupShape>().Any())
                    continue;

                if (anchor is not null)
                    worksheetPart.DrawingsPart.WorksheetDrawing!.RemoveChild(anchor);

                worksheetPart.DrawingsPart.DeletePart(removedPicture);
            }

            xlPictures.Deleted.Clear();

            // Pictures deleted from inside a group are removed in place so the group and its other
            // shapes survive. Guard on Count so we don't materialize (and thus re-serialize) the
            // drawing DOM for sheets that have nothing to remove.
            if (xlPictures.DeletedFromGroups.Count > 0)
            {
                RemoveGroupedPictures(worksheetPart.DrawingsPart, xlPictures.DeletedFromGroups);
                xlPictures.DeletedFromGroups.Clear();
            }

            // Newly requested groups must be built before the picture loop, which then sees their
            // members as ordinary grouped pictures.
            if (xlPictures.PendingGroups.Count > 0)
            {
                CreateGroups(worksheetPart.DrawingsPart, xlPictures.PendingGroups);
                xlPictures.PendingGroups.Clear();
            }
        }

        var groupedPictures = new List<XLPicture>();
        foreach (var pic in xlWorksheet.Pictures)
        {
            var xlPic = (XLPicture)pic;
            if (xlPic.IsInGroup)
                groupedPictures.Add(xlPic);
            else
                AddPictureAnchor(worksheetPart, pic, context);
        }

        foreach (var groupedPicture in groupedPictures)
        {
            if (groupedPicture.GroupInfo!.IsNew)
                InsertGroupedPicture(worksheetPart, groupedPicture, context);
            else
                UpdateGroupedPicture(worksheetPart, groupedPicture);
        }

        // Rebasing renumbers every NonVisualDrawingProperties id. That would break connector
        // start/end connection references (a:stCxn/@id, a:endCxn/@id) inside a group, so only do
        // it for the historical picture-only case where the drawing contains no group shapes.
        if (xlWorksheet.Pictures.Count > 0 && groupedPictures.Count == 0 &&
            !(worksheetPart.DrawingsPart?.WorksheetDrawing?.Descendants<Xdr.GroupShape>().Any() ?? false))
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
        var hasNewCharts = xlWorksheet.Charts.Any(c => ((XLChart)c).IsNew);
        if (worksheetPart.DrawingsPart is not null && // There is a drawing part for the sheet that could be deleted
            xlWorksheet
                .LegacyDrawingId is null && // and sheet doesn't contain any form controls or comments or other shapes
            xlWorksheet.Pictures.Count == 0 && // and also no pictures.
            !hasCharts && // and no existing chart parts
            !hasNewCharts && // and no new charts pending write
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

    private static void AddPictureAnchor(WorksheetPart worksheetPart, IXLPicture picture, SaveContext context)
    {
        var pic = (XLPicture)picture;
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
        var existingAnchor = GetAnchorFromImageId(drawingsPart, pic.RelId!);

        // Never overwrite an anchor that hosts a group shape (xdr:grpSp): replacing it with a
        // single regenerated picture would discard the other pictures/connectors/shapes in the
        // group. Such pictures are not loaded into the model (see DrawingPartReader.LoadDrawings),
        // so this is a defensive guard to keep grouped drawings intact on save.
        if (existingAnchor is not null && existingAnchor.Descendants<Xdr.GroupShape>().Any())
            return;

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
            case XLPicturePlacement.FreeFloating:
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

            case XLPicturePlacement.MoveAndSize:
                var moveAndSizeFromMarker = pic.Markers[XLMarkerPosition.TopLeft];
                if (moveAndSizeFromMarker == null)
                    moveAndSizeFromMarker = new XLMarker(picture.Worksheet.Cell("A1"));
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

                var moveAndSizeToMarker = pic.Markers[XLMarkerPosition.BottomRight];
                if (moveAndSizeToMarker == null)
                    moveAndSizeToMarker = new XLMarker(picture.Worksheet.Cell("A1"),
                        new Point(picture.Width, picture.Height));
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

            case XLPicturePlacement.Move:
                var moveFromMarker = pic.Markers[XLMarkerPosition.TopLeft];
                if (moveFromMarker == null) moveFromMarker = new XLMarker(picture.Worksheet.Cell("A1"));
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

    /// <summary>
    /// Update a picture that lives inside a group shape in place: re-feed its (possibly replaced)
    /// image data and, if the picture was resized, write the new size back into its child-space
    /// extent. The surrounding group — its other pictures, connectors and shapes — is left
    /// untouched. The picture is matched to its <c>xdr:pic</c> element by drawing id, because the
    /// save DOM is an independent re-parse of the package and object references from load don't
    /// survive.
    /// </summary>
    /// <summary>
    /// Build the groups requested via <c>XLPictures.Group(...)</c>. For each pending group a new
    /// group shape is created in an absolute anchor matching its bounding box, and every member's
    /// existing top-level <c>xdr:pic</c> is moved into it (its child <c>off</c>/<c>ext</c> set to the
    /// member's absolute sheet position/size, since the group uses an identity child coordinate
    /// space). The members' now-empty top-level anchors are removed.
    /// </summary>
    private static void CreateGroups(DrawingsPart drawingsPart, ICollection<XLPendingGroup> pendingGroups)
    {
        var worksheetDrawing = drawingsPart.WorksheetDrawing;
        if (worksheetDrawing is null)
            return;

        foreach (var pending in pendingGroups)
        {
            uint maxId = 0;
            foreach (var nvdpr in worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>())
                maxId = Math.Max(maxId, nvdpr.Id?.Value ?? 0);
            var groupId = maxId + 1;

            var groupShape = new Xdr.GroupShape(
                new Xdr.NonVisualGroupShapeProperties(
                    new Xdr.NonVisualDrawingProperties { Id = groupId, Name = $"Group {groupId}" },
                    new Xdr.NonVisualGroupShapeDrawingProperties()
                ),
                new Xdr.GroupShapeProperties(
                    new TransformGroup(
                        new Offset { X = pending.OffsetX, Y = pending.OffsetY },
                        new Extents { Cx = pending.ExtentCx, Cy = pending.ExtentCy },
                        new ChildOffset { X = pending.OffsetX, Y = pending.OffsetY },
                        new ChildExtents { Cx = pending.ExtentCx, Cy = pending.ExtentCy }
                    )
                )
            );

            foreach (var member in pending.Members)
            {
                var anchor = string.IsNullOrEmpty(member.RelId)
                    ? null
                    : GetAnchorFromImageId(drawingsPart, member.RelId!);
                var picElement = anchor?.Descendants<Xdr.Picture>().FirstOrDefault();
                if (anchor is null || picElement is null)
                    continue;

                var wb = member.Worksheet.Workbook;
                var transform = picElement.ShapeProperties?.Transform2D;
                if (transform is not null)
                {
                    // Identity child space: child coordinates are the member's absolute sheet EMU.
                    transform.Offset ??= new Offset();
                    transform.Offset.X = ConvertToEnglishMetricUnits(member.Left, wb.DpiX);
                    transform.Offset.Y = ConvertToEnglishMetricUnits(member.Top, wb.DpiY);
                    transform.Extents ??= new Extents();
                    transform.Extents.Cx = ConvertToEnglishMetricUnits(member.Width, wb.DpiX);
                    transform.Extents.Cy = ConvertToEnglishMetricUnits(member.Height, wb.DpiY);
                }

                picElement.Remove();
                groupShape.Append(picElement);
                anchor.Remove();

                member.GroupInfo = new XLPictureGroup
                {
                    ScaleX = 1.0,
                    ScaleY = 1.0,
                    OffsetX = 0.0,
                    OffsetY = 0.0,
                    GroupId = groupId,
                    GroupKey = member.GroupInfo?.GroupKey ?? 0,
                    LoadedWidthPx = member.Width,
                    LoadedHeightPx = member.Height,
                    LoadedLeftPx = member.Left,
                    LoadedTopPx = member.Top,
                };
            }

            var absoluteAnchor = new Xdr.AbsoluteAnchor(
                new Xdr.Position { X = pending.OffsetX, Y = pending.OffsetY },
                new Xdr.Extent { Cx = pending.ExtentCx, Cy = pending.ExtentCy },
                groupShape,
                new Xdr.ClientData()
            );
            worksheetDrawing.Append(absoluteAnchor);
        }
    }

    /// <summary>
    /// Insert a newly added picture into its target group. Allocates a drawing-wide unique id and a
    /// new image part, builds the <c>xdr:pic</c> with its child <c>off</c>/<c>ext</c> derived from the
    /// requested sheet geometry via the inverse group transform, and appends it to the group element.
    /// The model is then reset so a subsequent save treats it as an existing grouped picture.
    /// </summary>
    private static void InsertGroupedPicture(WorksheetPart worksheetPart, XLPicture pic, SaveContext context)
    {
        var drawingsPart = worksheetPart.DrawingsPart;
        var group = pic.GroupInfo;
        if (drawingsPart?.WorksheetDrawing is null || group?.GroupId is null)
            return;

        var worksheetDrawing = drawingsPart.WorksheetDrawing;

        Xdr.GroupShape? groupElement = null;
        foreach (var candidate in worksheetDrawing.Descendants<Xdr.GroupShape>())
        {
            if (candidate.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value == group.GroupId.Value)
            {
                groupElement = candidate;
                break;
            }
        }

        if (groupElement is null)
            return;

        // A drawing id must be unique across the whole drawing (pictures, connectors, shapes, groups).
        uint maxId = 0;
        foreach (var nvdpr in worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>())
            maxId = Math.Max(maxId, nvdpr.Id?.Value ?? 0);
        var newId = maxId + 1;

        var relId = context.RelIdGenerator.GetNext(RelType.Workbook);
        var imagePart = drawingsPart.AddImagePart(pic.Format.ToOpenXml(), relId);
        pic.ImageStream.Position = 0;
        imagePart.FeedData(pic.ImageStream);

        var wb = pic.Worksheet.Workbook;
        var sheetEmuCx = ConvertToEnglishMetricUnits(pic.Width, wb.DpiX);
        var sheetEmuCy = ConvertToEnglishMetricUnits(pic.Height, wb.DpiY);
        var sheetEmuX = ConvertToEnglishMetricUnits(pic.Left, wb.DpiX);
        var sheetEmuY = ConvertToEnglishMetricUnits(pic.Top, wb.DpiY);

        var childCx = group.ScaleX == 0 ? sheetEmuCx : (long)Math.Round(sheetEmuCx / group.ScaleX);
        var childCy = group.ScaleY == 0 ? sheetEmuCy : (long)Math.Round(sheetEmuCy / group.ScaleY);
        var childX = group.ScaleX == 0 ? sheetEmuX : (long)Math.Round((sheetEmuX - group.OffsetX) / group.ScaleX);
        var childY = group.ScaleY == 0 ? sheetEmuY : (long)Math.Round((sheetEmuY - group.OffsetY) / group.ScaleY);

        var picElement = new Xdr.Picture(
            new Xdr.NonVisualPictureProperties(
                new Xdr.NonVisualDrawingProperties { Id = newId, Name = pic.Name },
                new Xdr.NonVisualPictureDrawingProperties(new PictureLocks { NoChangeAspect = true })
            ),
            new Xdr.BlipFill(
                new Blip { Embed = relId },
                new Stretch(new FillRectangle())
            ),
            new Xdr.ShapeProperties(
                new Transform2D(
                    new Offset { X = childX, Y = childY },
                    new Extents { Cx = childCx, Cy = childCy }
                ),
                new PresetGeometry { Preset = ShapeTypeValues.Rectangle }
            )
        );

        groupElement.Append(picElement);

        // Reset the model so further edits/saves treat this as an existing grouped picture.
        pic.Id = (int)newId;
        pic.RelId = relId;
        pic.GroupInfo = new XLPictureGroup
        {
            ScaleX = group.ScaleX,
            ScaleY = group.ScaleY,
            OffsetX = group.OffsetX,
            OffsetY = group.OffsetY,
            GroupId = group.GroupId,
            GroupKey = group.GroupKey,
            LoadedWidthPx = pic.Width,
            LoadedHeightPx = pic.Height,
            LoadedLeftPx = pic.Left,
            LoadedTopPx = pic.Top,
        };
    }

    private static void UpdateGroupedPicture(WorksheetPart worksheetPart, XLPicture pic)
    {
        var drawingsPart = worksheetPart.DrawingsPart;
        var group = pic.GroupInfo;
        if (drawingsPart?.WorksheetDrawing is null || group is null)
            return;

        var worksheetDrawing = drawingsPart.WorksheetDrawing;

        Xdr.Picture? picElement = null;
        foreach (var candidate in worksheetDrawing.Descendants<Xdr.Picture>())
        {
            var id = candidate.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value;
            if (id == (uint)pic.Id && candidate.Ancestors<Xdr.GroupShape>().Any())
            {
                picElement = candidate;
                break;
            }
        }

        if (picElement is null)
            return;

        // Re-feed the image bytes into the existing part. If the image was not replaced these are
        // the same bytes that were read, so the part is unchanged.
        if (!string.IsNullOrEmpty(pic.RelId) && drawingsPart.HasPartWithId(pic.RelId!))
        {
            var imagePart = (ImagePart)drawingsPart.GetPartById(pic.RelId!);
            pic.ImageStream.Position = 0;
            imagePart.FeedData(pic.ImageStream);
        }

        var sizeChanged = pic.Width != group.LoadedWidthPx || pic.Height != group.LoadedHeightPx;
        var positionChanged = pic.Left != group.LoadedLeftPx || pic.Top != group.LoadedTopPx;

        // Leave the geometry untouched when nothing changed, so an unedited picture round-trips
        // without rounding drift. The group's own bounding box (ext/chExt) is kept fixed.
        if (!sizeChanged && !positionChanged)
            return;

        var transform = picElement.ShapeProperties?.Transform2D;
        if (transform is null)
            return;

        var wb = pic.Worksheet.Workbook;

        if (sizeChanged)
        {
            var sheetEmuCx = ConvertToEnglishMetricUnits(pic.Width, wb.DpiX);
            var sheetEmuCy = ConvertToEnglishMetricUnits(pic.Height, wb.DpiY);

            // Convert the sheet-space size back to the group's child coordinate space.
            var childCx = group.ScaleX == 0 ? sheetEmuCx : (long)Math.Round(sheetEmuCx / group.ScaleX);
            var childCy = group.ScaleY == 0 ? sheetEmuCy : (long)Math.Round(sheetEmuCy / group.ScaleY);

            transform.Extents ??= new Extents();
            transform.Extents.Cx = childCx;
            transform.Extents.Cy = childCy;
        }

        if (positionChanged)
        {
            var sheetEmuX = ConvertToEnglishMetricUnits(pic.Left, wb.DpiX);
            var sheetEmuY = ConvertToEnglishMetricUnits(pic.Top, wb.DpiY);

            // Invert the composed affine (sheet = offset + child·scale) to get the child a:off.
            var childOffX = group.ScaleX == 0 ? sheetEmuX : (long)Math.Round((sheetEmuX - group.OffsetX) / group.ScaleX);
            var childOffY = group.ScaleY == 0 ? sheetEmuY : (long)Math.Round((sheetEmuY - group.OffsetY) / group.ScaleY);

            transform.Offset ??= new Offset();
            transform.Offset.X = childOffX;
            transform.Offset.Y = childOffY;
        }
    }

    /// <summary>
    /// Remove pictures that were deleted from inside a group. Only the matching <c>xdr:pic</c>
    /// element is removed (located by drawing id) — the group and its remaining pictures, connectors
    /// and shapes are left intact. The image part is deleted only when no other blip still references
    /// it. The group's bounding box is kept fixed.
    /// </summary>
    private static void RemoveGroupedPictures(DrawingsPart drawingsPart,
        ICollection<(int Id, string? RelId)> removed)
    {
        var worksheetDrawing = drawingsPart.WorksheetDrawing;
        if (worksheetDrawing is null)
            return;

        foreach (var (id, relId) in removed)
        {
            Xdr.Picture? picElement = null;
            foreach (var candidate in worksheetDrawing.Descendants<Xdr.Picture>())
            {
                var candidateId = candidate.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value;
                if (candidateId == (uint)id && candidate.Ancestors<Xdr.GroupShape>().Any())
                {
                    picElement = candidate;
                    break;
                }
            }

            picElement?.Remove();

            // Drop the image part only if nothing else references it any more.
            if (!string.IsNullOrEmpty(relId) && drawingsPart.HasPartWithId(relId!))
            {
                var stillReferenced = worksheetDrawing.Descendants<Blip>()
                    .Any(b => b.Embed?.Value == relId);
                if (!stillReferenced)
                    drawingsPart.DeletePart(relId!);
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
