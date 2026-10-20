using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel.Drawings;
using XLibur.Excel.IO.DrawingML;
using XLibur.Extensions;
using static XLibur.Excel.XLWorkbook;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// Pictures inside a group shape: building the groups requested since the load, and inserting,
/// updating and removing grouped pictures in place, so the group's other shapes are kept.
/// </summary>
internal static class GroupedPictureWriter
{
    // A group scale at or extremely near zero is degenerate: dividing by it would
    // overflow the EMU long, so callers fall back to the untransformed sheet value.
    private const double DegenerateScaleEpsilon = 1e-9;

    private static bool IsDegenerateScale(double scale) => Math.Abs(scale) < DegenerateScaleEpsilon;

    /// <summary>
    /// Build the groups requested via <c>XLPictures.Group(...)</c>. For each pending group a new
    /// group shape is created in an absolute anchor matching its bounding box, and every member's
    /// existing top-level <c>xdr:pic</c> is moved into it (its child <c>off</c>/<c>ext</c> set to the
    /// member's absolute sheet position/size, since the group uses an identity child coordinate
    /// space). The members' now-empty top-level anchors are removed.
    /// </summary>
    internal static void CreateGroups(DrawingsPart drawingsPart, ICollection<XLPendingGroup> pendingGroups)
    {
        var worksheetDrawing = drawingsPart.WorksheetDrawing;
        if (worksheetDrawing is null)
            return;

        foreach (var pending in pendingGroups)
        {
            var groupId = DrawingIds.NextDrawingId(worksheetDrawing);

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
            if (DrawingIds.FindPicture(drawingsPart.WorksheetDrawing!, member.Id, member.RelId) is not
                { Anchor: var anchor, Picture: var picElement })
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
    internal static void InsertGroupedPicture(WorksheetPart worksheetPart, XLPicture pic, SaveContext context)
    {
        var drawingsPart = worksheetPart.DrawingsPart;
        var group = pic.GroupInfo;
        if (drawingsPart?.WorksheetDrawing is null || group?.GroupId is null)
            return;

        var worksheetDrawing = drawingsPart.WorksheetDrawing;

        var groupElement = FindGroupShapeById(worksheetDrawing, group.GroupId.Value);
        if (groupElement is null)
            return;

        var newId = DrawingIds.NextDrawingId(worksheetDrawing);

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
    /// Update a picture that lives inside a group shape in place: write its name, re-feed its
    /// (possibly replaced) image data and, if the picture was resized, write the new size back into
    /// its child-space extent. The surrounding group — its other pictures, connectors and shapes — is left
    /// untouched. The picture is matched to its <c>xdr:pic</c> element by drawing id, because the
    /// save DOM is an independent re-parse of the package and object references from load don't
    /// survive.
    /// </summary>
    internal static void UpdateGroupedPicture(WorksheetPart worksheetPart, XLPicture pic)
    {
        var drawingsPart = worksheetPart.DrawingsPart;
        var group = pic.GroupInfo;
        if (drawingsPart?.WorksheetDrawing is null || group is null)
            return;

        var worksheetDrawing = drawingsPart.WorksheetDrawing;

        var picElement = FindGroupedPictureById(worksheetDrawing, pic.Id);
        if (picElement is null)
            return;

        // The name goes to drawings with the picture's id in the anchor that holds the group, as a
        // plain picture's goes to drawings in its own anchor.
        var anchor = picElement.Ancestors().FirstOrDefault(a => a.Parent is Xdr.WorksheetDrawing) ?? picElement;
        PictureContentWriter.WritePictureName(anchor, picElement, pic.Name);
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
    internal static void RemoveGroupedPictures(DrawingsPart drawingsPart,
        ICollection<(int Id, string? RelId)> removed)
    {
        var worksheetDrawing = drawingsPart.WorksheetDrawing;
        if (worksheetDrawing is null)
            return;

        var released = new List<string>();
        foreach (var (id, relId) in removed)
        {
            if (relId is not null)
                released.Add(relId);

            if (FindGroupedPictureById(worksheetDrawing, id) is { } picture)
            {
                released.AddRange(PictureContentWriter.RelationshipIds(picture));
                picture.Remove();
            }
        }

        PictureContentWriter.DeleteUnreferencedRelationships(drawingsPart, worksheetDrawing, released);
    }
}
