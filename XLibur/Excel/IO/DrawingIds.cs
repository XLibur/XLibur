using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel.IO.DrawingML;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// The drawing ids of a sheet's drawing: finding a picture by its id, taking the next free id, and
/// renumbering them all once the pictures are written.
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

    /// <summary>Renumbers every drawing id from 1, in document order.</summary>
    /// <remarks>
    /// The branches of an <c>mc:AlternateContent</c> are one drawing in two forms, and Excel gives
    /// them one id: a 3D model's frame and its fallback picture are both id 2. They keep one id
    /// between them, so the picture XLibur loads from the fallback still names the frame by its id.
    /// </remarks>
    internal static void RebaseNonVisualDrawingPropertiesIds(WorksheetPart worksheetPart)
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
