using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ContentManagers;
using XLibur.Excel.Drawings;
using XLibur.Excel.IO.DrawingML;
using static XLibur.Excel.IO.OpenXmlConst;
using static XLibur.Excel.XLWorkbook;
using Drawing = DocumentFormat.OpenXml.Spreadsheet.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// Writes a sheet's pictures and shapes into its drawing part. The drawing id repairs of the load
/// are written first (#859), then deletions and new groups, then each picture: a free-floating one
/// by <see cref="PictureAnchorWriter"/>, one in a group shape by <see cref="GroupedPictureWriter"/>.
/// The shapes follow, by <see cref="ShapeWriter"/>. Names, image links and image parts are
/// <see cref="PictureContentWriter"/>'s, and drawing ids are <see cref="DrawingIds"/>'.
/// </summary>
internal static class PictureWriter
{
    internal static void WriteDrawings(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet,
        WorksheetPart worksheetPart,
        SaveContext context)
    {
        var pictures = (XLPictures)xlWorksheet.Pictures;
        var shapes = (XLShapes)xlWorksheet.Shapes;
        var layout = GridLayout.OnFirstUse(xlWorksheet);

        // A drawing the save package does not have yet, as a sheet saved out of a template has
        // not, gets every picture and shape. Otherwise only the ones that changed since the load or
        // the last save are written, and a sheet with none is not touched at all: reading the drawing
        // attaches its DOM, and the SDK then writes the DOM back over the part's bytes.
        var drawingIsNew = worksheetPart.DrawingsPart is null;

        WriteDrawingIdRepairs(xlWorksheet, worksheetPart, context, drawingIsNew);

        // One instance serves the pictures and the shapes, so they never take the same new id. It
        // reads the drawing only when asked.
        var ids = new DrawingIds(worksheetPart);
        if (shapes.Deleted.Count > 0)
            ShapeWriter.RemoveDeleted(shapes, drawingIsNew ? null : ids, context);

        if (drawingIsNew || HasUnsavedChanges(pictures, layout))
            WritePictures(worksheetPart, xlWorksheet, context, ids, layout, drawingIsNew);

        if (shapes.Count > 0 && (drawingIsNew || shapes.HasUnsavedChanges))
        {
            var drawingsPart = DrawingPartScaffold.EnsureDrawingsPart(worksheetPart, context);
            ShapeWriter.WriteShapes(drawingsPart, shapes, context, ids, writeAll: drawingIsNew);
        }

        if ((xlWorksheet.Pictures.Count > 0 || shapes.Count > 0) && !worksheet.OfType<Drawing>().Any())
        {
            var worksheetDrawing = new Drawing { Id = worksheetPart.GetIdOfPart(worksheetPart.DrawingsPart!) };
            worksheetDrawing.AddNamespaceDeclaration("r", RelationshipsNs);
            worksheet.InsertAfter(worksheetDrawing, DrawingPartScaffold.ElementBeforeDrawing(worksheet, cm));
            cm.SetElement(XLWorksheetContents.Drawing, worksheetDrawing);
        }

        RemoveEmptyDrawingPart(worksheet, cm, xlWorksheet, worksheetPart);
    }

    /// <summary>
    /// Writes the fresh ids the load gave pictures and shapes whose id another had (#859) into the
    /// drawing, before anything reads its ids. The drawing holds them from this save on, so the
    /// repairs are dropped once it has succeeded.
    /// </summary>
    private static void WriteDrawingIdRepairs(XLWorksheet xlWorksheet, WorksheetPart worksheetPart,
        SaveContext context, bool drawingIsNew)
    {
        var repairs = xlWorksheet.DrawingIdRepairs;
        if (repairs.Count == 0)
            return;

        if (!drawingIsNew)
            repairs.Apply(worksheetPart.DrawingsPart!.WorksheetDrawing!);

        context.KeepAfterSave.Add(repairs.Clear);
    }

    /// <summary>
    /// Applies the picture deletions and new groups, then writes each picture that changed, or every
    /// picture into a drawing the save package does not have yet.
    /// </summary>
    private static void WritePictures(WorksheetPart worksheetPart, XLWorksheet xlWorksheet, SaveContext context,
        DrawingIds ids, Lazy<GridLayout> layout, bool drawingIsNew)
    {
        var pictures = (XLPictures)xlWorksheet.Pictures;
        if (worksheetPart.DrawingsPart != null)
            ProcessDeletedPicturesAndGroups(worksheetPart.DrawingsPart, pictures, ids);

        var written = new List<(XLPicture Picture, Xdr.Picture Element)>();
        WritePictureAnchors(worksheetPart, xlWorksheet, context, written, ids, layout, drawingIsNew);

        // The next save edits what this one wrote and finds each picture's anchor by its drawing id
        // and image link, so the model takes the ones the drawing now holds, once the save has
        // succeeded (XLWorkbook.KeepWhatTheSaveWrote).
        foreach (var (picture, element) in written)
        {
            context.PictureIdsForNextSave.Add((picture,
                (int)element.NonVisualPictureProperties!.NonVisualDrawingProperties!.Id!.Value,
                element.BlipFill!.Blip!.Embed!.Value!,
                element.OuterXml));
        }
    }

    /// <summary>Whether a picture was deleted, grouped, added or changed since the load or the last save.</summary>
    private static bool HasUnsavedChanges(XLPictures pictures, Lazy<GridLayout> layout)
    {
        if (pictures.Deleted.Count > 0 || pictures.DeletedFromGroups.Count > 0 || pictures.PendingGroups.Count > 0)
            return true;

        foreach (var picture in pictures)
        {
            if (picture.HasUnsavedChanges(layout))
                return true;
        }

        return false;
    }

    /// <summary>
    /// Apply pending deletions and group requests to the drawing before the picture-writing pass:
    /// remove deleted (non-grouped) pictures, remove pictures deleted from inside groups in place,
    /// and build any newly requested groups so their members appear as ordinary grouped pictures.
    /// </summary>
    private static void ProcessDeletedPicturesAndGroups(DrawingsPart drawingsPart, XLPictures xlPictures, DrawingIds ids)
    {
        if (xlPictures.Deleted.Count > 0 && drawingsPart.WorksheetDrawing is { } worksheetDrawing)
        {
            // The anchor found holds the picture itself, never a group shape, so a group and its
            // other shapes are never removed with it.
            var released = new List<string>();
            foreach (var (id, relId) in xlPictures.Deleted)
            {
                released.Add(relId);
                if (ids.FindPicture(id, relId) is { } found)
                {
                    released.AddRange(PictureContentWriter.RelationshipIds(found.Anchor));
                    found.Anchor.Remove();
                }
            }

            PictureContentWriter.DeleteUnreferencedRelationships(drawingsPart, worksheetDrawing, released);
        }

        xlPictures.Deleted.Clear();

        // Pictures deleted from inside a group are removed in place so the group and its other
        // shapes survive. Guard on Count so we don't materialize (and thus re-serialize) the
        // drawing DOM for sheets that have nothing to remove.
        if (xlPictures.DeletedFromGroups.Count > 0)
        {
            GroupedPictureWriter.RemoveGroupedPictures(drawingsPart, xlPictures.DeletedFromGroups, ids);
            xlPictures.DeletedFromGroups.Clear();
        }

        // Newly requested groups must be built before the picture loop, which then sees their
        // members as ordinary grouped pictures.
        if (xlPictures.PendingGroups.Count > 0)
        {
            GroupedPictureWriter.CreateGroups(drawingsPart, xlPictures.PendingGroups, ids);
            xlPictures.PendingGroups.Clear();
        }
    }

    /// <summary>
    /// Write every picture that changed: a new free-floating picture gets a fresh anchor, a loaded one
    /// is written into the anchor it came from, and a grouped picture is inserted into or updated
    /// within its group. With <paramref name="writeAll"/> every picture is written, changed or not.
    /// </summary>
    private static void WritePictureAnchors(WorksheetPart worksheetPart, XLWorksheet xlWorksheet, SaveContext context,
        List<(XLPicture Picture, Xdr.Picture Element)> written, DrawingIds ids, Lazy<GridLayout> layout, bool writeAll)
    {
        var groupedPictures = new List<XLPicture>();
        var images = new PictureContentWriter.ImageFeed(xlWorksheet);

        foreach (var pic in xlWorksheet.Pictures)
        {
            var xlPic = (XLPicture)pic;
            if (!writeAll && !xlPic.HasUnsavedChanges(layout))
                continue;

            if (xlPic.IsInGroup)
                groupedPictures.Add(xlPic);
            else
                written.Add((xlPic, PictureAnchorWriter.AddPictureAnchor(worksheetPart, xlPic, context, images, layout, ids)));
        }

        foreach (var groupedPicture in groupedPictures)
        {
            var element = groupedPicture.GroupInfo!.IsNew
                ? GroupedPictureWriter.InsertGroupedPicture(worksheetPart, groupedPicture, context, ids)
                : GroupedPictureWriter.UpdateGroupedPicture(worksheetPart, groupedPicture, ids);
            if (element is not null)
                written.Add((groupedPicture, element));
        }

        if (worksheetPart.DrawingsPart is { WorksheetDrawing: { } drawing } drawingsPart)
            PictureContentWriter.DeleteUnreferencedRelationships(drawingsPart, drawing, images.Released);
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
            xlWorksheet.Pictures.Count == 0 && // and also no pictures
            xlWorksheet.Shapes.Count == 0 && // or shapes.
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
}
