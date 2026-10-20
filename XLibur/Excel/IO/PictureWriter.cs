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
                written.Add((xlPic, AddPictureAnchor(worksheetPart, pic, context, fedImages, layout)));
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
    /// the save. A kept anchor is compared with <paramref name="layout"/>, the sheet's layout now.
    /// </summary>
    private static Xdr.Picture AddPictureAnchor(WorksheetPart worksheetPart, IXLPicture picture,
        SaveContext context, Dictionary<string, MemoryStream> fedImages, Lazy<GridLayout> layout)
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
        var existing = DrawingIds.FindPicture(worksheetDrawing, pic.Id, pic.RelId);

        var imagePart = PictureContentWriter.FeedImagePart(drawingsPart, pic, context, fedImages);
        var embedId = drawingsPart.GetIdOfPart(imagePart);

        if (pic.KeptAnchor is { } kept && existing is { Anchor: Xdr.TwoCellAnchor twoCellAnchor } &&
            TryWriteKeptAnchor(twoCellAnchor, existing.Value.Picture, kept, layout.Value, pic, embedId))
            return existing.Value.Picture;

        var wb = pic.Worksheet.Workbook;
        var extentsCx = DrawingUnits.PixelsToEmu(pic.Width, wb.DpiX);
        var extentsCy = DrawingUnits.PixelsToEmu(pic.Height, wb.DpiY);

        // The same xdr:pic whichever anchor form ends up holding it: the anchor decides where the
        // drawing goes, never what it is.
        var pictureElement = new Xdr.Picture(
            new Xdr.NonVisualPictureProperties(
                new Xdr.NonVisualDrawingProperties { Id = DrawingIds.NextDrawingId(worksheetDrawing), Name = pic.Name },
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

        // FindPicture finds a picture only as the anchor's child or a branch of the anchor's wrapper.
        if (DrawingAnchorContent.WrapperOf(found.Picture) is null)
        {
            worksheetDrawing.ReplaceChild(anchor, found.Anchor);
            return pictureElement;
        }

        // The picture is one branch of mc:AlternateContent, as the render of a 3D model is. The
        // wrapper and its other branches are the drawing; the regenerated anchor only says where
        // it goes.
        WriteAnchorPosition(worksheetDrawing, found.Anchor, anchor);
        PictureContentWriter.WritePictureNameAndImage(found.Anchor, found.Picture, pic.Name, embedId);
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
        GridLayout layout, XLPicture pic, string embedId)
    {
        if (!kept.IsLayoutUnchanged(layout) ||
            anchor.EditAs?.Value != kept.Placement.ToOpenXml() ||
            anchor.FromMarker is not { } from ||
            anchor.ToMarker is not { } to)
            return false;

        WriteMarkerCell(from, kept.From);
        WriteMarkerCell(to, kept.To);
        PictureContentWriter.WritePictureNameAndImage(anchor, picture, pic.Name, embedId);
        return true;
    }

    /// <summary>Writes a marker's column and row, zero-based, leaving its offsets alone.</summary>
    private static void WriteMarkerCell(Xdr.MarkerType marker, XLMarker cell)
    {
        marker.ColumnId = new Xdr.ColumnId((cell.ColumnNumber - 1).ToInvariantString());
        marker.RowId = new Xdr.RowId((cell.RowNumber - 1).ToInvariantString());
    }
}
