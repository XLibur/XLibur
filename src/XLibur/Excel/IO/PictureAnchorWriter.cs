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
/// Where a free-floating picture goes: a fresh anchor for a new picture, or the anchor it was loaded
/// from, edited in place.
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
    /// Write a free-floating picture: a new one into a fresh anchor, a loaded one into the anchor it
    /// was loaded from. Returns the picture written, whose drawing id and image link the picture
    /// takes after the save. A kept anchor is compared with <paramref name="layout"/>, the sheet's
    /// layout now.
    /// </summary>
    internal static Xdr.Picture AddPictureAnchor(WorksheetPart worksheetPart, XLPicture pic,
        SaveContext context, PictureContentWriter.ImageFeed images, Lazy<GridLayout> layout, DrawingIds ids)
    {
        var drawingsPart = worksheetPart.DrawingsPart ??
                           worksheetPart.AddNewPart<DrawingsPart>(context.RelIdGenerator.GetNext(RelType.Workbook));

        drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

        var worksheetDrawing = drawingsPart.WorksheetDrawing;

        EnsureDrawingNamespaces(worksheetDrawing);

        // Found by the image link the picture was loaded with, before the picture can take a part
        // of its own below. The anchor holds the picture itself, never a group shape, so a group is
        // never overwritten.
        var existing = ids.FindPicture(pic.Id, pic.RelId);

        var imagePart = PictureContentWriter.FeedImagePart(drawingsPart, pic, context, images);
        var embedId = drawingsPart.GetIdOfPart(imagePart);

        if (existing is { } found)
        {
            PatchPicture(worksheetDrawing, found.Anchor, found.Picture, pic, layout, embedId);
            return found.Picture;
        }

        var drawingId = ids.NextDrawingId();
        var pictureElement = (pic.LoadedPictureXml is { } template ? FromTemplate(template, pic, embedId, drawingId) : null) ??
                             NewPictureElement(pic, embedId, drawingId);
        PictureStyleWriter.Write(pictureElement, pic);
        worksheetDrawing.Append(DrawingAnchorFactory.Create(pic.Placement, GeometryOf(pic), pictureElement));
        return pictureElement;
    }

    /// <summary>
    /// The <c>xdr:pic</c> of a copy of a loaded picture: the <c>xdr:pic</c> the source was loaded from,
    /// so the copy keeps what XLibur does not model, given the copy's id, name, image and size. Or
    /// <c>null</c> when the source's XML cannot be read back.
    /// </summary>
    /// <remarks>
    /// Anything else in it that names a relationship, as a hyperlink or an SVG form of the image
    /// does, names one of the source's drawing part, so it is dropped rather than copied dangling.
    /// </remarks>
    private static Xdr.Picture? FromTemplate(string template, XLPicture pic, string embedId, uint drawingId)
    {
        Xdr.Picture picture;
        try
        {
            picture = new Xdr.Picture(template);
        }
        catch (System.Xml.XmlException)
        {
            return null;
        }

        if (picture.NonVisualPictureProperties?.NonVisualDrawingProperties is not { } properties ||
            picture.BlipFill?.Blip is not { } blip)
            return null;

        properties.Id = drawingId;
        properties.Name = pic.Name;
        blip.Link = null;
        blip.Embed = embedId;
        DrawingTemplate.DropRelationshipReferences(picture, keep: blip);

        if (picture.ShapeProperties?.Transform2D is { } transform)
        {
            var wb = pic.Worksheet.Workbook;
            var offset = transform.Offset ?? transform.PrependChild(new Offset());
            offset.X = 0;
            offset.Y = 0;
            var extents = transform.Extents ?? transform.InsertAfter(new Extents(), offset);
            extents.Cx = DrawingUnits.PixelsToEmu(pic.Width, wb.DpiX);
            extents.Cy = DrawingUnits.PixelsToEmu(pic.Height, wb.DpiY);
        }

        return picture;
    }

    /// <summary>
    /// The <c>xdr:pic</c> of a new picture. The same element whichever anchor form ends up holding
    /// it: the anchor decides where the drawing goes, never what it is.
    /// </summary>
    private static Xdr.Picture NewPictureElement(XLPicture pic, string embedId, uint drawingId)
    {
        var wb = pic.Worksheet.Workbook;
        return new Xdr.Picture(
            new Xdr.NonVisualPictureProperties(
                new Xdr.NonVisualDrawingProperties { Id = drawingId, Name = pic.Name },
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
                    new Extents
                    {
                        Cx = DrawingUnits.PixelsToEmu(pic.Width, wb.DpiX),
                        Cy = DrawingUnits.PixelsToEmu(pic.Height, wb.DpiY)
                    }
                ),
                new PresetGeometry { Preset = ShapeTypeValues.Rectangle }
            )
        );
    }

    /// <summary>
    /// Where the model puts the picture. A marker left null here is not an oversight:
    /// <see cref="DrawingAnchorFactory"/> falls back to A1, and a picture nobody moved arrives with
    /// both of them unset. See its remarks.
    /// </summary>
    private static DrawingAnchorGeometry GeometryOf(XLPicture pic) => new()
    {
        Worksheet = pic.Worksheet,
        LeftPx = pic.Left,
        TopPx = pic.Top,
        WidthPx = pic.Width,
        HeightPx = pic.Height,
        FromMarker = pic.Markers[XLMarkerPosition.TopLeft],
        ToMarker = pic.Markers[XLMarkerPosition.BottomRight]
    };

    /// <summary>
    /// Write what changed of a loaded picture into the anchor and <c>xdr:pic</c> it was loaded from,
    /// and nothing else, so whatever Excel put there that XLibur does not model is kept: a rotation,
    /// a border, a shadow, a crop, the other branches of an <c>mc:AlternateContent</c>.
    /// </summary>
    /// <remarks>
    /// The position is written only when the placement, a marker, the size or the position changed,
    /// or when the rows and columns under a kept anchor were resized. A kept anchor whose layout
    /// held has only its markers' cells written. A change of placement changes the anchor's form,
    /// and the drawing moves into an anchor of the new form with all it holds.
    /// </remarks>
    private static void PatchPicture(Xdr.WorksheetDrawing worksheetDrawing, OpenXmlCompositeElement anchor,
        Xdr.Picture picture, XLPicture pic, Lazy<GridLayout> layout, string embedId)
    {
        var keptLayoutChanged = pic.KeptAnchor is { } keptAnchor && !keptAnchor.IsLayoutUnchanged(layout.Value);
        if (pic.IsGeometryChanged || keptLayoutChanged)
        {
            if (!(pic.KeptAnchor is { } kept && anchor is Xdr.TwoCellAnchor twoCellAnchor &&
                  TryWriteKeptAnchor(twoCellAnchor, kept, layout.Value)))
            {
                anchor = WriteAnchorPosition(worksheetDrawing, anchor,
                    DrawingAnchorFactory.Create(pic.Placement, GeometryOf(pic), new Xdr.Picture()));
            }

            if (pic.IsSizeChanged)
                WritePictureSize(picture, pic);
        }

        PictureContentWriter.WritePictureNameAndImage(anchor, picture, pic.Name, embedId);
        PictureStyleWriter.Write(picture, pic);
    }

    /// <summary>
    /// Writes the picture's size into its <c>a:xfrm</c>, leaving the offset, the rotation and the
    /// flips as they were. Excel places an anchored picture by its anchor, but keeps the two sizes
    /// equal.
    /// </summary>
    private static void WritePictureSize(Xdr.Picture picture, XLPicture pic)
    {
        if (picture.ShapeProperties?.Transform2D is not { } transform)
            return;

        var wb = pic.Worksheet.Workbook;
        var extents = transform.Extents ?? transform.InsertAfter(new Extents(), transform.Offset);
        extents.Cx = DrawingUnits.PixelsToEmu(pic.Width, wb.DpiX);
        extents.Cy = DrawingUnits.PixelsToEmu(pic.Height, wb.DpiY);
    }

    /// <summary>
    /// Gives <paramref name="existing"/> the position of the regenerated <paramref name="generated"/>
    /// and keeps everything else it holds: the drawing, in whatever wrapper Excel put it, and the
    /// client data.
    /// </summary>
    /// <remarks>
    /// The anchor is edited in place when it keeps its form. When the placement changed it, as a
    /// <c>MoveTo(cell)</c> makes a one-cell anchor out of a two-cell one, the generated anchor takes
    /// the existing anchor's place and its contents. Returns the anchor that holds the drawing now.
    /// </remarks>
    private static OpenXmlCompositeElement WriteAnchorPosition(Xdr.WorksheetDrawing worksheetDrawing,
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

            return existing;
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
        return generated;
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
    /// a row or column edit may have moved. The cell offsets, <c>a:xfrm</c> and anything Excel added
    /// that XLibur does not model stay as the file had them. Returns <c>false</c> when the anchor is not the one that was loaded, or
    /// when the rows and columns under it were resized so the kept markers would no longer place the
    /// picture where Excel showed it; the caller then regenerates it.
    /// </remarks>
    private static bool TryWriteKeptAnchor(Xdr.TwoCellAnchor anchor, XLKeptAnchor kept, GridLayout layout)
    {
        if (!kept.IsLayoutUnchanged(layout) ||
            anchor.EditAs?.Value != kept.Placement.ToOpenXml() ||
            anchor.FromMarker is not { } from ||
            anchor.ToMarker is not { } to)
            return false;

        WriteMarkerCell(from, kept.From);
        WriteMarkerCell(to, kept.To);
        return true;
    }

    /// <summary>Writes a marker's column and row, zero-based, leaving its offsets alone.</summary>
    private static void WriteMarkerCell(Xdr.MarkerType marker, XLMarker cell)
    {
        marker.ColumnId = new Xdr.ColumnId((cell.ColumnNumber - 1).ToInvariantString());
        marker.RowId = new Xdr.RowId((cell.RowNumber - 1).ToInvariantString());
    }
}
