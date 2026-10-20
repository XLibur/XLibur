using System.Collections.Generic;
using System.IO;
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
/// Writes a sheet's pictures into its drawing part. Deletions and new groups are applied first,
/// then each picture is written: a free-floating one by <see cref="PictureAnchorWriter"/>, one in a
/// group shape by <see cref="GroupedPictureWriter"/>. Names, image links and image parts are
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
        if (worksheetPart.DrawingsPart != null)
            ProcessDeletedPicturesAndGroups(worksheetPart.DrawingsPart, (XLPictures)xlWorksheet.Pictures);

        var written = new List<(XLPicture Picture, Xdr.Picture Element)>();
        var groupedPictureCount = WritePictureAnchors(worksheetPart, xlWorksheet, context, written);

        // Rebasing renumbers every NonVisualDrawingProperties id. That would break connector
        // start/end connection references (a:stCxn/@id, a:endCxn/@id) inside a group, so only do
        // it for the historical picture-only case where the drawing contains no group shapes.
        if (xlWorksheet.Pictures.Count > 0 && groupedPictureCount == 0 &&
            !(worksheetPart.DrawingsPart?.WorksheetDrawing?.Descendants<Xdr.GroupShape>().Any() ?? false))
            DrawingIds.RebaseNonVisualDrawingPropertiesIds(worksheetPart);

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
            var released = new List<string>();
            foreach (var (id, relId) in xlPictures.Deleted)
            {
                released.Add(relId);
                if (DrawingIds.FindPicture(worksheetDrawing, id, relId) is { } found)
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
            GroupedPictureWriter.RemoveGroupedPictures(drawingsPart, xlPictures.DeletedFromGroups);
            xlPictures.DeletedFromGroups.Clear();
        }

        // Newly requested groups must be built before the picture loop, which then sees their
        // members as ordinary grouped pictures.
        if (xlPictures.PendingGroups.Count > 0)
        {
            GroupedPictureWriter.CreateGroups(drawingsPart, xlPictures.PendingGroups);
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
        var layout = GridLayout.OnFirstUse(xlWorksheet);

        foreach (var pic in xlWorksheet.Pictures)
        {
            var xlPic = (XLPicture)pic;
            if (xlPic.IsInGroup)
                groupedPictures.Add(xlPic);
            else
                written.Add((xlPic, PictureAnchorWriter.AddPictureAnchor(worksheetPart, pic, context, fedImages, layout)));
        }

        foreach (var groupedPicture in groupedPictures)
        {
            if (groupedPicture.GroupInfo!.IsNew)
                GroupedPictureWriter.InsertGroupedPicture(worksheetPart, groupedPicture, context);
            else
                GroupedPictureWriter.UpdateGroupedPicture(worksheetPart, groupedPicture);
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
}
