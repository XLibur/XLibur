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
internal static class DrawingIds
{
    /// <summary>
    /// The picture with drawing id <paramref name="id"/> and image link <paramref name="relId"/>, and
    /// the top-level anchor that holds it, or <c>null</c>. The link alone is not enough: pictures that
    /// share an image part, as a copied picture does in Excel, all have the same link (issue #781).
    /// </summary>
    /// <remarks>
    /// The picture is found where the loader reads it (<see cref="DrawingAnchorContent.IsAnchorPicture"/>):
    /// the anchor's child, or a branch of the <c>mc:AlternateContent</c> Excel puts it in when the
    /// drawing also has a newer form. A picture in a group shape is never found, because grouped
    /// pictures are written in place (<see cref="GroupedPictureWriter.UpdateGroupedPicture"/>) and
    /// the group's anchor holds its other shapes too.
    /// </remarks>
    internal static (OpenXmlCompositeElement Anchor, Xdr.Picture Picture)? FindPicture(
        Xdr.WorksheetDrawing worksheetDrawing, int id, string? relId)
    {
        if (string.IsNullOrEmpty(relId))
            return null;

        foreach (var anchor in worksheetDrawing.ChildElements.OfType<OpenXmlCompositeElement>())
        {
            foreach (var picture in anchor.Descendants<Xdr.Picture>())
            {
                if (picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value == (uint)id &&
                    picture.BlipFill?.Blip?.Embed?.Value == relId &&
                    DrawingAnchorContent.IsAnchorPicture(picture, anchor))
                    return (anchor, picture);
            }
        }

        return null;
    }

    /// <summary>
    /// The next free drawing id. A drawing id must be unique across the whole drawing (pictures,
    /// connectors, shapes, groups).
    /// </summary>
    internal static uint NextDrawingId(Xdr.WorksheetDrawing worksheetDrawing)
    {
        uint maxId = 0;
        foreach (var nvdpr in worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>())
            maxId = Math.Max(maxId, nvdpr.Id?.Value ?? 0);
        return maxId + 1;
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
