using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel.IO.DrawingML;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// The drawing ids of a sheet's drawing: finding a picture by its id, and taking the next free id.
/// </summary>
/// <remarks>
/// One instance serves one save of one sheet's drawing. It reads the drawing once, on the first
/// call that needs it, rather than once per picture, which made a save of n pictures cost n² (issue
/// #797): an index of the pictures by drawing id and image link, and the highest drawing id. The
/// sheet may have no drawing when the save starts, so the drawing is taken from the sheet's part at
/// that first call, by which time the save has made it.
/// </remarks>
internal sealed class DrawingIds(WorksheetPart worksheetPart)
{
    private Xdr.WorksheetDrawing? _drawing;
    private Dictionary<(uint Id, string RelId), List<(OpenXmlCompositeElement Anchor, Xdr.Picture Picture)>>? _pictures;
    private Dictionary<uint, List<Xdr.Picture>>? _groupedPictures;
    private Dictionary<uint, (OpenXmlCompositeElement Anchor, Xdr.Shape Shape)>? _shapes;
    private uint? _maxId;

    private Xdr.WorksheetDrawing Drawing => _drawing ??= worksheetPart.DrawingsPart!.WorksheetDrawing!;

    /// <summary>
    /// The picture with drawing id <paramref name="id"/> and image link <paramref name="relId"/>, and
    /// the top-level anchor that holds it, or <c>null</c>. The link alone is not enough: pictures that
    /// share an image part, as a copied picture does in Excel, all have the same link (issue #781).
    /// </summary>
    /// <remarks>
    /// <para>
    /// The picture is found where the loader reads it (<see cref="DrawingAnchorContent.IsAnchorPicture"/>):
    /// the anchor's child, or a branch of the <c>mc:AlternateContent</c> Excel puts it in when the
    /// drawing also has a newer form. A picture in a group shape is never found, because grouped
    /// pictures are written in place (<see cref="GroupedPictureWriter.UpdateGroupedPicture"/>) and
    /// the group's anchor holds its other shapes too.
    /// </para>
    /// <para>
    /// The index holds the pictures the drawing had at the first call, in document order, and a
    /// picture is found only while it is still where the index saw it, with the same id and link: an
    /// anchor the save has since removed or replaced, or a picture it has moved into a group, is not
    /// found. A picture the save writes is not added. Each picture is written once and a drawing id
    /// names one drawing, so no later lookup asks for it.
    /// </para>
    /// </remarks>
    internal (OpenXmlCompositeElement Anchor, Xdr.Picture Picture)? FindPicture(int id, string? relId)
    {
        if (string.IsNullOrEmpty(relId))
            return null;

        var key = ((uint)id, relId);
        _pictures ??= IndexPictures(Drawing);
        if (!_pictures.TryGetValue(key, out var candidates))
            return null;

        foreach (var (anchor, picture) in candidates)
        {
            if (anchor.Parent == Drawing && KeyOf(picture, anchor) == key)
                return (anchor, picture);
        }

        return null;
    }

    /// <summary>
    /// The first picture in document order, inside a group shape, with drawing id <paramref name="id"/>,
    /// or <c>null</c>. The save DOM is an independent re-parse, so a grouped picture is matched by id
    /// rather than object reference.
    /// </summary>
    /// <remarks>
    /// The grouped pictures are indexed by id at the first call, rather than read for each grouped
    /// picture the save updates or removes (#804). A picture is found only while it is still in the
    /// drawing, in a group, under the same id: one the save has since removed is not. The save
    /// moves pictures into new groups after it removes grouped pictures, so building the new groups
    /// drops the index (<see cref="ForgetGroupedPictures"/>), and the next call indexes them too.
    /// </remarks>
    internal Xdr.Picture? FindGroupedPicture(int id)
    {
        _groupedPictures ??= IndexGroupedPictures(Drawing);
        if (!_groupedPictures.TryGetValue((uint)id, out var candidates))
            return null;

        foreach (var picture in candidates)
        {
            if (picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value == (uint)id &&
                IsGroupedIn(picture, Drawing))
                return picture;
        }

        return null;
    }

    /// <summary>
    /// The shape with drawing id <paramref name="id"/> that an anchor holds directly, and that anchor,
    /// or <c>null</c>. Those are the shapes the loader reads (spec 15); one in a group is not found.
    /// </summary>
    /// <remarks>
    /// The shapes are indexed at the first call. A load gives each loaded shape an id of its own, so
    /// the first in document order is the one; a shape is found only while it is still where the
    /// index saw it, under the same id.
    /// </remarks>
    internal (OpenXmlCompositeElement Anchor, Xdr.Shape Shape)? FindShape(int id)
    {
        _shapes ??= IndexShapes(Drawing);
        if (!_shapes.TryGetValue((uint)id, out var found))
            return null;

        return found.Anchor.Parent == Drawing && found.Shape.Parent == found.Anchor && ShapeIdOf(found.Shape) == (uint)id
            ? found
            : null;
    }

    /// <summary>The drawing id of a shape, or <c>null</c> when it has none.</summary>
    internal static uint? ShapeIdOf(Xdr.Shape shape) =>
        shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;

    private static Dictionary<uint, (OpenXmlCompositeElement Anchor, Xdr.Shape Shape)> IndexShapes(Xdr.WorksheetDrawing worksheetDrawing)
    {
        var shapes = new Dictionary<uint, (OpenXmlCompositeElement Anchor, Xdr.Shape Shape)>();
        foreach (var anchor in worksheetDrawing.ChildElements.OfType<OpenXmlCompositeElement>())
        {
            foreach (var shape in anchor.Elements<Xdr.Shape>())
            {
                if (ShapeIdOf(shape) is { } id)
                    shapes.TryAdd(id, (anchor, shape));
            }
        }

        return shapes;
    }

    /// <summary>Drops the index of grouped pictures, after the save has moved pictures into a group.</summary>
    internal void ForgetGroupedPictures() => _groupedPictures = null;

    private static Dictionary<uint, List<Xdr.Picture>> IndexGroupedPictures(Xdr.WorksheetDrawing worksheetDrawing)
    {
        var pictures = new Dictionary<uint, List<Xdr.Picture>>();
        foreach (var picture in worksheetDrawing.Descendants<Xdr.Picture>())
        {
            if (picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value is not { } id ||
                !picture.Ancestors<Xdr.GroupShape>().Any())
                continue;

            if (!pictures.TryGetValue(id, out var candidates))
                pictures[id] = candidates = [];

            candidates.Add(picture);
        }

        return pictures;
    }

    /// <summary>Whether <paramref name="picture"/> is in a group shape of <paramref name="worksheetDrawing"/>.</summary>
    private static bool IsGroupedIn(Xdr.Picture picture, Xdr.WorksheetDrawing worksheetDrawing)
    {
        var grouped = false;
        for (var element = picture.Parent; element is not null; element = element.Parent)
        {
            if (element == worksheetDrawing)
                return grouped;

            grouped |= element is Xdr.GroupShape;
        }

        return false;
    }

    /// <summary>
    /// Takes the next free drawing id: one more than the highest id in the drawing, or than the last
    /// id taken. A drawing id must be unique across the whole drawing (pictures, connectors, shapes,
    /// groups), so the caller puts the id it takes into the drawing.
    /// </summary>
    /// <remarks>
    /// The drawing is read at the first call. The save removes deleted pictures before it takes an
    /// id, so an id a deleted picture had is free again, as it is in the file.
    /// </remarks>
    internal uint NextDrawingId()
    {
        _maxId ??= HighestDrawingId(Drawing);
        _maxId++;
        return _maxId.Value;
    }

    private static Dictionary<(uint Id, string RelId), List<(OpenXmlCompositeElement Anchor, Xdr.Picture Picture)>>
        IndexPictures(Xdr.WorksheetDrawing worksheetDrawing)
    {
        var pictures = new Dictionary<(uint Id, string RelId), List<(OpenXmlCompositeElement Anchor, Xdr.Picture Picture)>>();
        foreach (var anchor in worksheetDrawing.ChildElements.OfType<OpenXmlCompositeElement>())
        {
            foreach (var picture in anchor.Descendants<Xdr.Picture>())
            {
                if (KeyOf(picture, anchor) is not { } key)
                    continue;

                if (!pictures.TryGetValue(key, out var candidates))
                    pictures[key] = candidates = [];

                candidates.Add((anchor, picture));
            }
        }

        return pictures;
    }

    /// <summary>
    /// The drawing id and image link <paramref name="picture"/> is found by, or <c>null</c> when it
    /// has no id or link, or is not the picture <paramref name="anchor"/> places.
    /// </summary>
    private static (uint Id, string RelId)? KeyOf(Xdr.Picture picture, OpenXmlCompositeElement anchor) =>
        picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value is { } id &&
        picture.BlipFill?.Blip?.Embed?.Value is { Length: > 0 } relId &&
        DrawingAnchorContent.IsAnchorPicture(picture, anchor)
            ? (id, relId)
            : null;

    private static uint HighestDrawingId(Xdr.WorksheetDrawing worksheetDrawing)
    {
        uint maxId = 0;
        foreach (var nvdpr in worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>())
            maxId = Math.Max(maxId, nvdpr.Id?.Value ?? 0);
        return maxId;
    }
}
