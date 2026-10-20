using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel.Drawings;
using XLibur.Excel.IO.DrawingML;
using static XLibur.Excel.IO.OpenXmlConst;
using static XLibur.Excel.XLWorkbook;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// Where a free-floating picture goes: a fresh anchor for a new picture, the anchor it was loaded
/// from kept and edited, or a regenerated anchor in place of the loaded one.
/// </summary>
internal static class PictureAnchorWriter
{
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
    internal static Xdr.Picture AddPictureAnchor(WorksheetPart worksheetPart, IXLPicture picture,
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
