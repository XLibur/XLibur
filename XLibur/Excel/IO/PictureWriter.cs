using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ContentManagers;
using XLibur.Excel.Drawings;
using XLibur.Excel.IO.DrawingML;
using XLibur.Extensions;
using static XLibur.Excel.IO.OpenXmlConst;
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
            ProcessDeletedPicturesAndGroups(worksheetPart.DrawingsPart, (XLPictures)xlWorksheet.Pictures);

        var written = new List<(XLPicture Picture, Xdr.Picture Element)>();
        var groupedPictureCount = WritePictureAnchors(worksheetPart, xlWorksheet, context, written);

        // Rebasing renumbers every NonVisualDrawingProperties id. That would break connector
        // start/end connection references (a:stCxn/@id, a:endCxn/@id) inside a group, so only do
        // it for the historical picture-only case where the drawing contains no group shapes.
        if (xlWorksheet.Pictures.Count > 0 && groupedPictureCount == 0 &&
            !(worksheetPart.DrawingsPart?.WorksheetDrawing?.Descendants<Xdr.GroupShape>().Any() ?? false))
            RebaseNonVisualDrawingPropertiesIds(worksheetPart);

        // The next save edits what this one wrote and finds each picture's anchor by its drawing
        // id and image link, so the model takes the ones the drawing now holds, once the save has
        // succeeded (XLWorkbook.KeepWhatTheSaveWrote).
        foreach (var (picture, element) in written)
        {
            context.PictureIdsForNextSave.Add((picture,
                (int)element.NonVisualPictureProperties!.NonVisualDrawingProperties!.Id!.Value,
                element.BlipFill!.Blip!.Embed!.Value!));
        }

        if (xlWorksheet.Pictures.Count > 0 && !worksheet.OfType<Drawing>().Any())
        {
            var worksheetDrawing = new Drawing { Id = worksheetPart.GetIdOfPart(worksheetPart.DrawingsPart!) };
            worksheetDrawing.AddNamespaceDeclaration("r", RelationshipsNs);
            worksheet.InsertAfter(worksheetDrawing, DrawingPartScaffold.ElementBeforeDrawing(worksheet, cm));
            cm.SetElement(XLWorksheetContents.Drawing, worksheetDrawing);
        }

        RemoveEmptyDrawingPart(worksheet, cm, xlWorksheet, worksheetPart);
    }

    /// <summary>
    /// Apply pending deletions and group requests to the drawing before the picture-writing pass:
    /// remove deleted (non-grouped) pictures, remove pictures deleted from inside groups in place,
    /// and build any newly requested groups so their members appear as ordinary grouped pictures.
    /// </summary>
    private static void ProcessDeletedPicturesAndGroups(DrawingsPart drawingsPart, XLPictures xlPictures)
    {
        if (xlPictures.Deleted.Count > 0 && drawingsPart.WorksheetDrawing is { } worksheetDrawing)
        {
            // The anchor found holds the picture itself, never a group shape, so a group and its
            // other shapes are never removed with it.
            foreach (var (id, relId) in xlPictures.Deleted)
            {
                FindPictureAnchor(worksheetDrawing, id, relId)?.Remove();

                // Other pictures may still show the image: Excel shares one part between copies.
                if (drawingsPart.HasPartWithId(relId))
                    DeleteImagePartIfUnreferenced(drawingsPart, worksheetDrawing, relId);
            }
        }

        xlPictures.Deleted.Clear();

        // Pictures deleted from inside a group are removed in place so the group and its other
        // shapes survive. Guard on Count so we don't materialize (and thus re-serialize) the
        // drawing DOM for sheets that have nothing to remove.
        if (xlPictures.DeletedFromGroups.Count > 0)
        {
            RemoveGroupedPictures(drawingsPart, xlPictures.DeletedFromGroups);
            xlPictures.DeletedFromGroups.Clear();
        }

        // Newly requested groups must be built before the picture loop, which then sees their
        // members as ordinary grouped pictures.
        if (xlPictures.PendingGroups.Count > 0)
        {
            CreateGroups(drawingsPart, xlPictures.PendingGroups);
            xlPictures.PendingGroups.Clear();
        }
    }

    /// <summary>
    /// Write every picture's anchor: free-floating pictures get a fresh anchor, grouped pictures are
    /// inserted into or updated within their group. Returns the number of grouped pictures handled.
    /// </summary>
    private static int WritePictureAnchors(WorksheetPart worksheetPart, XLWorksheet xlWorksheet, SaveContext context,
        List<(XLPicture Picture, Xdr.Picture Element)> written)
    {
        var groupedPictures = new List<XLPicture>();
        var fedImages = new Dictionary<string, MemoryStream>();
        foreach (var pic in xlWorksheet.Pictures)
        {
            var xlPic = (XLPicture)pic;
            if (xlPic.IsInGroup)
                groupedPictures.Add(xlPic);
            else
                written.Add((xlPic, AddPictureAnchor(worksheetPart, pic, context, fedImages)));
        }

        foreach (var groupedPicture in groupedPictures)
        {
            if (groupedPicture.GroupInfo!.IsNew)
                InsertGroupedPicture(worksheetPart, groupedPicture, context);
            else
                UpdateGroupedPicture(worksheetPart, groupedPicture);
        }

        return groupedPictures.Count;
    }

    /// <summary>
    /// Remove the sheet's DrawingsPart (and its workbook reference) when it would otherwise be saved
    /// empty — no pictures, charts, legacy shapes, or other drawing children. Avoids writing an empty
    /// Drawings.xml part.
    /// </summary>
    private static void RemoveEmptyDrawingPart(Worksheet worksheet, XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet, WorksheetPart worksheetPart)
    {
        var hasCharts = worksheetPart.DrawingsPart is not null && worksheetPart.DrawingsPart.Parts.Any();
        var hasNewCharts = xlWorksheet.Charts.Any(c => ((XLChart)c).IsNew);
        if (worksheetPart.DrawingsPart is not null && // There is a drawing part for the sheet that could be deleted
            xlWorksheet
                .LegacyDrawingId is null && // and sheet doesn't contain any form controls or comments or other shapes
            xlWorksheet.Pictures.Count == 0 && // and also no pictures.
            !hasCharts && // and no existing chart parts
            !hasNewCharts && // and no new charts pending write
                             // Check for non-picture shapes (textboxes, rectangles, etc.) last, and
                             // check them by streaming the part rather than through WorksheetDrawing.
                             // Reading that property attaches the DOM, and the SDK then writes the
                             // attached tree back over the part's original bytes on save even though
                             // nothing was modified — which is the very re-serialization the ordering
                             // here was meant to avoid.
            !DrawingPartProbe.HasAnyChild(worksheetPart.DrawingsPart))
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

    // A group scale at or extremely near zero is degenerate: dividing by it would
    // overflow the EMU long, so callers fall back to the untransformed sheet value.
    private const double DegenerateScaleEpsilon = 1e-9;

    private static bool IsDegenerateScale(double scale) => Math.Abs(scale) < DegenerateScaleEpsilon;

    private static void EnsureDrawingNamespaces(Xdr.WorksheetDrawing worksheetDrawing)
    {
        if (!worksheetDrawing.NamespaceDeclarations.Any(nd =>
                nd.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/main")))
            worksheetDrawing.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

        if (!worksheetDrawing.NamespaceDeclarations.Any(nd =>
                nd.Value.Equals(RelationshipsNs)))
            worksheetDrawing.AddNamespaceDeclaration("r", RelationshipsNs);
    }

    /// <summary>
    /// Write a free-floating picture's anchor, editing or replacing its own anchor from the loaded
    /// drawing. Returns the picture written, whose drawing id and image link the picture takes after
    /// the save.
    /// </summary>
    private static Xdr.Picture AddPictureAnchor(WorksheetPart worksheetPart, IXLPicture picture,
        SaveContext context, Dictionary<string, MemoryStream> fedImages)
    {
        var pic = (XLPicture)picture;
        var drawingsPart = worksheetPart.DrawingsPart ??
                           worksheetPart.AddNewPart<DrawingsPart>(context.RelIdGenerator.GetNext(RelType.Workbook));

        drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

        var worksheetDrawing = drawingsPart.WorksheetDrawing;

        EnsureDrawingNamespaces(worksheetDrawing);

        // Found by the image link the picture was loaded with, before the picture can take a part
        // of its own below. The anchor holds the picture itself, never a group shape, so a group is
        // never overwritten.
        var existing = FindPicture(worksheetDrawing, pic.Id, pic.RelId);

        var imagePart = FeedImagePart(drawingsPart, pic, context, fedImages);
        var embedId = drawingsPart.GetIdOfPart(imagePart);

        if (pic.KeptAnchor is { } kept && existing is { Anchor: Xdr.TwoCellAnchor twoCellAnchor } &&
            TryWriteKeptAnchor(twoCellAnchor, existing.Value.Picture, kept, pic, embedId))
            return existing.Value.Picture;

        var wb = pic.Worksheet.Workbook;
        var extentsCx = DrawingUnits.PixelsToEmu(pic.Width, wb.DpiX);
        var extentsCy = DrawingUnits.PixelsToEmu(pic.Height, wb.DpiY);

        // The same xdr:pic whichever anchor form ends up holding it: the anchor decides where the
        // drawing goes, never what it is.
        var pictureElement = new Xdr.Picture(
            new Xdr.NonVisualPictureProperties(
                new Xdr.NonVisualDrawingProperties { Id = NextDrawingId(worksheetDrawing), Name = pic.Name },
                new Xdr.NonVisualPictureDrawingProperties(new PictureLocks { NoChangeAspect = true })
            ),
            new Xdr.BlipFill(
                new Blip
                {
                    Embed = embedId,
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
        );

        // A marker left null here is not an oversight: DrawingAnchorFactory falls back to A1, and a
        // picture nobody moved arrives with both of them unset. See its remarks.
        var anchor = DrawingAnchorFactory.Create(
            pic.Placement,
            new DrawingAnchorGeometry
            {
                Worksheet = picture.Worksheet,
                LeftPx = pic.Left,
                TopPx = pic.Top,
                WidthPx = pic.Width,
                HeightPx = pic.Height,
                FromMarker = pic.Markers[XLMarkerPosition.TopLeft],
                ToMarker = pic.Markers[XLMarkerPosition.BottomRight]
            },
            pictureElement);

        if (existing is not { } found)
        {
            worksheetDrawing.Append(anchor);
            return pictureElement;
        }

        if (found.Picture.Parent == found.Anchor)
        {
            worksheetDrawing.ReplaceChild(anchor, found.Anchor);
            return pictureElement;
        }

        // The picture is one branch of mc:AlternateContent, as the render of a 3D model is. The
        // wrapper and its other branches are the drawing; the regenerated anchor only says where
        // it goes.
        WriteAnchorPosition(worksheetDrawing, found.Anchor, anchor);
        WritePictureNameAndImage(found.Anchor, found.Picture, pic.Name, embedId);
        return found.Picture;
    }

    /// <summary>
    /// Gives <paramref name="existing"/> the position of the regenerated <paramref name="generated"/>
    /// and keeps everything else it holds: the drawing, in whatever wrapper Excel put it, and the
    /// client data.
    /// </summary>
    /// <remarks>
    /// The anchor is edited in place when it keeps its form. When the placement changed it, as a
    /// <c>MoveTo(cell)</c> makes a one-cell anchor out of a two-cell one, the generated anchor takes
    /// the existing anchor's place and its contents.
    /// </remarks>
    private static void WriteAnchorPosition(Xdr.WorksheetDrawing worksheetDrawing,
        OpenXmlCompositeElement existing, OpenXmlCompositeElement generated)
    {
        var position = generated.ChildElements.Where(IsAnchorPosition).ToList();
        foreach (var element in position)
            element.Remove();

        if (existing.GetType() == generated.GetType())
        {
            foreach (var element in existing.ChildElements.Where(IsAnchorPosition).ToList())
                element.Remove();

            for (var i = position.Count - 1; i >= 0; i--)
                existing.PrependChild(position[i]);

            if (existing is Xdr.TwoCellAnchor twoCellAnchor)
                twoCellAnchor.EditAs = ((Xdr.TwoCellAnchor)generated).EditAs;

            return;
        }

        generated.RemoveAllChildren();
        foreach (var element in position)
            generated.AppendChild(element);

        foreach (var element in existing.ChildElements.Where(e => !IsAnchorPosition(e)).ToList())
        {
            element.Remove();
            generated.AppendChild(element);
        }

        worksheetDrawing.ReplaceChild(generated, existing);
    }

    /// <summary>The markers, position and extent that say where an anchor puts its drawing.</summary>
    private static bool IsAnchorPosition(OpenXmlElement element) =>
        element is Xdr.FromMarker or Xdr.ToMarker or Xdr.Position or Xdr.Extent;

    /// <summary>
    /// Writes the name and image link the model owns into a picture written back into its anchor.
    /// </summary>
    /// <remarks>
    /// The name goes to every drawing in the anchor with the picture's id. Excel gives each branch
    /// of <c>mc:AlternateContent</c> the same id and name, as one drawing in two forms, so a 3D
    /// model renamed through its fallback picture is renamed in the form Excel shows.
    /// </remarks>
    private static void WritePictureNameAndImage(OpenXmlElement anchor, Xdr.Picture picture, string name,
        string embedId)
    {
        if (picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value is { } id)
        {
            foreach (var properties in anchor.Descendants<Xdr.NonVisualDrawingProperties>())
            {
                if (properties.Id?.Value == id)
                    properties.Name = name;
            }
        }

        if (picture.BlipFill?.Blip is { } blip)
            blip.Embed = embedId;
    }

    /// <summary>
    /// The image part <paramref name="pic"/> is written with, holding its image. Each part is fed
    /// once per save, by the first picture that shows it; <paramref name="fedImages"/> holds, by
    /// relationship id, the image each part was fed this save. The picture's own link is left as it
    /// was: it takes the link the save wrote once the save has succeeded, as it takes its id.
    /// </summary>
    /// <remarks>
    /// Pictures that share a part, as Excel's copies of a picture do (#781), feed it once. A copy
    /// whose image differs from what its part was fed, because the image was changed after the load,
    /// shows a part that holds its own image instead: one this save already fed with the same image,
    /// or a new one. So a changed image is not overwritten by the copies that kept theirs, and does
    /// not reach them.
    /// </remarks>
    private static ImagePart FeedImagePart(DrawingsPart drawingsPart, XLPicture pic, SaveContext context,
        Dictionary<string, MemoryStream> fedImages)
    {
        if (pic.RelId is { } relId && drawingsPart.GetPartOrNull<ImagePart>(relId) is { } loaded)
        {
            if (!fedImages.TryGetValue(relId, out var fed))
                return Feed(loaded, relId);

            if (SameImage(fed, pic.ImageStream))
                return loaded;

            foreach (var (otherRelId, otherImage) in fedImages)
            {
                if (SameImage(otherImage, pic.ImageStream))
                    return drawingsPart.GetPartOrNull<ImagePart>(otherRelId)!;
            }
        }

        var newRelId = context.RelIdGenerator.GetNext(RelType.Workbook);
        return Feed(drawingsPart.AddImagePart(pic.Format.ToOpenXml(), newRelId), newRelId);

        ImagePart Feed(ImagePart part, string partRelId)
        {
            pic.ImageStream.Position = 0;
            part.FeedData(pic.ImageStream);
            fedImages[partRelId] = pic.ImageStream;
            return part;
        }
    }

    private static bool SameImage(MemoryStream left, MemoryStream right) =>
        ReferenceEquals(left, right) || Bytes(left).SequenceEqual(Bytes(right));

    private static ReadOnlySpan<byte> Bytes(MemoryStream stream) =>
        stream.TryGetBuffer(out var buffer) ? buffer.AsSpan() : stream.ToArray();

    /// <summary>
    /// The top-level anchor that holds the picture with drawing id <paramref name="id"/> and image
    /// link <paramref name="relId"/>, or <c>null</c>. The link alone is not enough: pictures that
    /// share an image part, as a copied picture does in Excel, all have the same link (issue #781).
    /// </summary>
    /// <remarks>
    /// The picture is looked for anywhere in the anchor, as the loader finds it, and not only as its
    /// child: Excel puts it inside <c>mc:AlternateContent</c> when the drawing also has a newer form.
    /// An anchor that holds a group shape is skipped, because its pictures are written in place
    /// (<see cref="UpdateGroupedPicture"/>) and the anchor holds the group's other shapes too.
    /// </remarks>
    private static OpenXmlElement? FindPictureAnchor(Xdr.WorksheetDrawing worksheetDrawing, int id, string? relId) =>
        FindPicture(worksheetDrawing, id, relId)?.Anchor;

    /// <summary>
    /// The picture with drawing id <paramref name="id"/> and image link <paramref name="relId"/>, and
    /// the top-level anchor that holds it, or <c>null</c>. See <see cref="FindPictureAnchor"/>.
    /// </summary>
    private static (OpenXmlCompositeElement Anchor, Xdr.Picture Picture)? FindPicture(
        Xdr.WorksheetDrawing worksheetDrawing, int id, string? relId)
    {
        if (string.IsNullOrEmpty(relId))
            return null;

        foreach (var anchor in worksheetDrawing.ChildElements.OfType<OpenXmlCompositeElement>())
        {
            if (anchor.Elements<Xdr.GroupShape>().Any())
                continue;

            foreach (var picture in anchor.Descendants<Xdr.Picture>())
            {
                if (picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value == (uint)id &&
                    picture.BlipFill?.Blip?.Embed?.Value == relId)
                    return (anchor, picture);
            }
        }

        return null;
    }

    /// <summary>
    /// Write a picture the caller has not placed back into the two-cell anchor it was loaded from,
    /// rather than regenerating it. See <see cref="XLKeptAnchor"/> for why the model cannot.
    /// </summary>
    /// <remarks>
    /// The anchor is edited, not replaced, as a slicer's is: only each marker's column and row, which
    /// a row or column edit may have moved, and the picture's name and image link, which the model
    /// owns. The cell offsets, <c>a:xfrm</c> and anything Excel added that XLibur does not model stay
    /// as the file had them. Returns <c>false</c> when the anchor is not the one that was loaded, or
    /// when the rows and columns under it were resized so the kept markers would no longer place the
    /// picture where Excel showed it; the caller then regenerates it.
    /// </remarks>
    private static bool TryWriteKeptAnchor(Xdr.TwoCellAnchor anchor, Xdr.Picture picture, XLKeptAnchor kept,
        XLPicture pic, string embedId)
    {
        if (!kept.IsLayoutUnchanged ||
            anchor.EditAs?.Value != kept.Placement.ToOpenXml() ||
            anchor.FromMarker is not { } from ||
            anchor.ToMarker is not { } to)
            return false;

        WriteMarkerCell(from, kept.From);
        WriteMarkerCell(to, kept.To);
        WritePictureNameAndImage(anchor, picture, pic.Name, embedId);
        return true;
    }

    /// <summary>Writes a marker's column and row, zero-based, leaving its offsets alone.</summary>
    private static void WriteMarkerCell(Xdr.MarkerType marker, XLMarker cell)
    {
        marker.ColumnId = new Xdr.ColumnId((cell.ColumnNumber - 1).ToInvariantString());
        marker.RowId = new Xdr.RowId((cell.RowNumber - 1).ToInvariantString());
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
            var groupId = NextDrawingId(worksheetDrawing);

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

            MoveMembersIntoGroup(drawingsPart, pending, groupShape, groupId);

            // Propagate the new group id to every member sharing this group's key — including
            // pictures added to the group before its first save, which carry a null group id and
            // would otherwise be skipped during insertion.
            PropagateGroupIdToMembers(pending, groupId);

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
    /// Move each member's existing top-level <c>xdr:pic</c> into <paramref name="groupShape"/>, setting its
    /// child <c>off</c>/<c>ext</c> to the member's absolute sheet geometry (identity child space) and removing
    /// the now-empty top-level anchor.
    /// </summary>
    private static void MoveMembersIntoGroup(DrawingsPart drawingsPart, XLPendingGroup pending,
        Xdr.GroupShape groupShape, uint groupId)
    {
        foreach (var member in pending.Members)
        {
            var anchor = FindPictureAnchor(drawingsPart.WorksheetDrawing!, member.Id, member.RelId);
            var picElement = anchor?.Descendants<Xdr.Picture>().FirstOrDefault();
            if (anchor is null || picElement is null)
                continue;

            var wb = member.Worksheet.Workbook;
            var transform = picElement.ShapeProperties?.Transform2D;
            if (transform is not null)
            {
                // Identity child space: child coordinates are the member's absolute sheet EMU.
                transform.Offset ??= new Offset();
                transform.Offset.X = DrawingUnits.PixelsToEmu(member.Left, wb.DpiX);
                transform.Offset.Y = DrawingUnits.PixelsToEmu(member.Top, wb.DpiY);
                transform.Extents ??= new Extents();
                transform.Extents.Cx = DrawingUnits.PixelsToEmu(member.Width, wb.DpiX);
                transform.Extents.Cy = DrawingUnits.PixelsToEmu(member.Height, wb.DpiY);
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
    }

    /// <summary>
    /// Propagate <paramref name="groupId"/> to every picture on the worksheet sharing the group's key,
    /// including members added before the group's first save (which carry a null group id).
    /// </summary>
    private static void PropagateGroupIdToMembers(XLPendingGroup pending, uint groupId)
    {
        var groupKey = pending.Members.Count > 0 ? pending.Members[0].GroupInfo?.GroupKey : null;
        if (groupKey is null)
            return;

        foreach (var pic in (IEnumerable<XLPicture>)(XLPictures)pending.Members[0].Worksheet.Pictures)
        {
            if (pic.GroupInfo is { } info && info.GroupKey == groupKey && info.GroupId != groupId)
                pic.GroupInfo = info with { GroupId = groupId };
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

        var groupElement = FindGroupShapeById(worksheetDrawing, group.GroupId.Value);
        if (groupElement is null)
            return;

        var newId = NextDrawingId(worksheetDrawing);

        var relId = context.RelIdGenerator.GetNext(RelType.Workbook);
        var imagePart = drawingsPart.AddImagePart(pic.Format.ToOpenXml(), relId);
        pic.ImageStream.Position = 0;
        imagePart.FeedData(pic.ImageStream);

        var wb = pic.Worksheet.Workbook;
        var sheetEmuCx = DrawingUnits.PixelsToEmu(pic.Width, wb.DpiX);
        var sheetEmuCy = DrawingUnits.PixelsToEmu(pic.Height, wb.DpiY);
        var sheetEmuX = DrawingUnits.PixelsToEmu(pic.Left, wb.DpiX);
        var sheetEmuY = DrawingUnits.PixelsToEmu(pic.Top, wb.DpiY);

        var childCx = IsDegenerateScale(group.ScaleX) ? sheetEmuCx : (long)Math.Round(sheetEmuCx / group.ScaleX);
        var childCy = IsDegenerateScale(group.ScaleY) ? sheetEmuCy : (long)Math.Round(sheetEmuCy / group.ScaleY);
        var childX = IsDegenerateScale(group.ScaleX) ? sheetEmuX : (long)Math.Round((sheetEmuX - group.OffsetX) / group.ScaleX);
        var childY = IsDegenerateScale(group.ScaleY) ? sheetEmuY : (long)Math.Round((sheetEmuY - group.OffsetY) / group.ScaleY);

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

    private static Xdr.GroupShape? FindGroupShapeById(Xdr.WorksheetDrawing worksheetDrawing, uint groupId)
    {
        foreach (var candidate in worksheetDrawing.Descendants<Xdr.GroupShape>())
        {
            if (candidate.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value == groupId)
                return candidate;
        }

        return null;
    }

    /// <summary>
    /// The next free drawing id. A drawing id must be unique across the whole drawing (pictures,
    /// connectors, shapes, groups).
    /// </summary>
    private static uint NextDrawingId(Xdr.WorksheetDrawing worksheetDrawing)
    {
        uint maxId = 0;
        foreach (var nvdpr in worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>())
            maxId = Math.Max(maxId, nvdpr.Id?.Value ?? 0);
        return maxId + 1;
    }

    private static void UpdateGroupedPicture(WorksheetPart worksheetPart, XLPicture pic)
    {
        var drawingsPart = worksheetPart.DrawingsPart;
        var group = pic.GroupInfo;
        if (drawingsPart?.WorksheetDrawing is null || group is null)
            return;

        var worksheetDrawing = drawingsPart.WorksheetDrawing;

        var picElement = FindGroupedPictureById(worksheetDrawing, pic.Id);
        if (picElement is null)
            return;

        RefeedGroupedPictureImage(drawingsPart, pic);
        UpdateGroupedPictureGeometry(picElement, pic, group);
    }

    /// <summary>
    /// Re-feed the image bytes into the existing part. If the image was not replaced these are
    /// the same bytes that were read, so the part is unchanged.
    /// </summary>
    private static void RefeedGroupedPictureImage(DrawingsPart drawingsPart, XLPicture pic)
    {
        if (drawingsPart.GetPartOrNull<ImagePart>(pic.RelId) is not { } imagePart)
            return;

        pic.ImageStream.Position = 0;
        imagePart.FeedData(pic.ImageStream);
    }

    private static void UpdateGroupedPictureGeometry(Xdr.Picture picElement, XLPicture pic, XLPictureGroup group)
    {
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
            ApplyGroupedChildExtents(transform, pic, group, wb);

        if (positionChanged)
            ApplyGroupedChildOffset(transform, pic, group, wb);
    }

    /// <summary>
    /// Locate the <c>xdr:pic</c> inside a group shape that carries the given drawing id. The save DOM
    /// is an independent re-parse, so the picture is matched by id rather than object reference.
    /// </summary>
    private static Xdr.Picture? FindGroupedPictureById(Xdr.WorksheetDrawing worksheetDrawing, int id)
    {
        foreach (var candidate in worksheetDrawing.Descendants<Xdr.Picture>())
        {
            var candidateId = candidate.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value;
            if (candidateId == (uint)id && candidate.Ancestors<Xdr.GroupShape>().Any())
                return candidate;
        }

        return null;
    }

    private static void ApplyGroupedChildExtents(Transform2D transform, XLPicture pic, XLPictureGroup group, XLWorkbook wb)
    {
        var sheetEmuCx = DrawingUnits.PixelsToEmu(pic.Width, wb.DpiX);
        var sheetEmuCy = DrawingUnits.PixelsToEmu(pic.Height, wb.DpiY);

        // Convert the sheet-space size back to the group's child coordinate space.
        var childCx = IsDegenerateScale(group.ScaleX) ? sheetEmuCx : (long)Math.Round(sheetEmuCx / group.ScaleX);
        var childCy = IsDegenerateScale(group.ScaleY) ? sheetEmuCy : (long)Math.Round(sheetEmuCy / group.ScaleY);

        transform.Extents ??= new Extents();
        transform.Extents.Cx = childCx;
        transform.Extents.Cy = childCy;
    }

    private static void ApplyGroupedChildOffset(Transform2D transform, XLPicture pic, XLPictureGroup group, XLWorkbook wb)
    {
        var sheetEmuX = DrawingUnits.PixelsToEmu(pic.Left, wb.DpiX);
        var sheetEmuY = DrawingUnits.PixelsToEmu(pic.Top, wb.DpiY);

        // Invert the composed affine (sheet = offset + child·scale) to get the child a:off.
        var childOffX = IsDegenerateScale(group.ScaleX) ? sheetEmuX : (long)Math.Round((sheetEmuX - group.OffsetX) / group.ScaleX);
        var childOffY = IsDegenerateScale(group.ScaleY) ? sheetEmuY : (long)Math.Round((sheetEmuY - group.OffsetY) / group.ScaleY);

        transform.Offset ??= new Offset();
        transform.Offset.X = childOffX;
        transform.Offset.Y = childOffY;
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
            FindGroupedPictureById(worksheetDrawing, id)?.Remove();

            // Drop the image part only if nothing else references it any more.
            if (drawingsPart.HasPartWithId(relId))
                DeleteImagePartIfUnreferenced(drawingsPart, worksheetDrawing, relId);
        }
    }

    private static void DeleteImagePartIfUnreferenced(DrawingsPart drawingsPart,
        Xdr.WorksheetDrawing worksheetDrawing, string relId)
    {
        var stillReferenced = worksheetDrawing.Descendants<Blip>()
            .Any(b => b.Embed?.Value == relId);
        if (!stillReferenced)
            drawingsPart.DeletePart(relId);
    }

    /// <summary>Renumbers every drawing id from 1, in document order.</summary>
    /// <remarks>
    /// The branches of an <c>mc:AlternateContent</c> are one drawing in two forms, and Excel gives
    /// them one id: a 3D model's frame and its fallback picture are both id 2. They keep one id
    /// between them, so the picture XLibur loads from the fallback still names the frame by its id.
    /// </remarks>
    private static void RebaseNonVisualDrawingPropertiesIds(WorksheetPart worksheetPart)
    {
        var worksheetDrawing = worksheetPart.DrawingsPart!.WorksheetDrawing;

        uint id = 1;
        var branchIds = new Dictionary<(AlternateContent Wrapper, uint OldId), uint>();
        foreach (var nvdpr in worksheetDrawing!.Descendants<Xdr.NonVisualDrawingProperties>())
        {
            if (nvdpr.Ancestors<AlternateContent>().FirstOrDefault() is { } wrapper &&
                nvdpr.Id?.Value is { } oldId)
            {
                if (!branchIds.TryGetValue((wrapper, oldId), out var newId))
                    branchIds[(wrapper, oldId)] = newId = id++;

                nvdpr.Id = newId;
                continue;
            }

            nvdpr.Id = id++;
        }
    }
}
