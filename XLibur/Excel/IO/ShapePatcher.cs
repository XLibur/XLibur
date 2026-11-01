using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using XLibur.Excel.Drawings;
using XLibur.Excel.IO.DrawingML;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// Writes what changed of a loaded shape into the anchor and <c>xdr:sp</c> it was read from, and
/// nothing else (spec 15). XLibur never regenerates XML it read, so what it does not model — a
/// rotation, a gradient, a shadow, 3-D, a hyperlink, adjust handles, the anchor's <c>editAs</c> —
/// stays as the file has it.
/// </summary>
/// <remarks>
/// Each write goes to the narrowest element that holds the value: the anchor's markers or extent
/// for a move, <c>a:xfrm/a:ext</c> for a resize, <c>cNvPr/@name</c> for a rename,
/// <see cref="ShapeFormatWriter"/> for the fill and line, <see cref="ShapeTextWriter"/> for the text.
/// </remarks>
internal static class ShapePatcher
{
    internal static void Patch(OpenXmlCompositeElement anchor, Xdr.Shape element, XLShape shape)
    {
        if (shape.IsGeometryChanged)
        {
            WriteAnchorPosition(anchor, shape);
            WriteTransform(element, shape);
        }

        if (shape.IsNameChanged && element.NonVisualShapeProperties?.NonVisualDrawingProperties is { } properties)
            properties.Name = shape.Name;

        PatchContent(element, shape);
    }

    /// <summary>
    /// Writes the fill, line and text edits not saved yet into <paramref name="element"/>, the
    /// shape's <c>xdr:sp</c> as the file holds it, or a copy of it.
    /// </summary>
    internal static void PatchContent(Xdr.Shape element, XLShape shape)
    {
        if (shape.Assigned != XLShapeProperties.None)
            ShapeFormatWriter.Write(element.ShapeProperties ??= new Xdr.ShapeProperties(), shape);

        if (shape.Body.HasChanges)
            ShapeTextWriter.Write(element.TextBody ?? AddTextBody(element), shape.Body, alignNewParagraphs: false);
    }

    /// <summary>
    /// Puts the model's position into the anchor: into its own markers, position and extent when the
    /// placement kept its form, keeping every attribute and child it has besides, such as
    /// <c>editAs</c>. A new placement is a new form of anchor, and the shape moves into it with the
    /// client data, as a picture does.
    /// </summary>
    private static void WriteAnchorPosition(OpenXmlCompositeElement anchor, XLShape shape)
    {
        var generated = DrawingAnchorFactory.Create(shape.Placement, ShapeWriter.GeometryOf(shape), new Xdr.Shape());
        var position = generated.ChildElements.Where(IsAnchorPosition).ToList();
        foreach (var part in position)
            part.Remove();

        if (generated.GetType() == anchor.GetType())
            ReplaceChangedParts(anchor, position, shape);
        else
            MoveIntoNewAnchor(anchor, generated, position);
    }

    /// <summary>
    /// Puts each part of <paramref name="position"/> whose model value changed in the place of the
    /// anchor's own. The rest stays as the file has it, rather than coming back from pixels half a
    /// pixel off.
    /// </summary>
    private static void ReplaceChangedParts(OpenXmlCompositeElement anchor, List<OpenXmlElement> position, XLShape shape)
    {
        foreach (var part in position)
        {
            var changed = part switch
            {
                Xdr.FromMarker or Xdr.Position => shape.IsTopLeftChanged,
                Xdr.ToMarker => shape.IsBottomRightChanged,
                _ => shape.IsSizeChanged,
            };
            if (!changed)
                continue;

            if (anchor.ChildElements.FirstOrDefault(e => e.GetType() == part.GetType()) is { } old)
                anchor.ReplaceChild(part, old);
            else
                anchor.InsertAt(part, position.IndexOf(part));
        }
    }

    /// <summary>Moves everything the anchor holds but its position into <paramref name="generated"/>, which takes its place.</summary>
    private static void MoveIntoNewAnchor(OpenXmlCompositeElement anchor, OpenXmlCompositeElement generated,
        List<OpenXmlElement> position)
    {
        generated.RemoveAllChildren();
        foreach (var part in position)
            generated.AppendChild(part);

        foreach (var child in anchor.ChildElements.Where(e => !IsAnchorPosition(e)).ToList())
        {
            child.Remove();
            generated.AppendChild(child);
        }

        anchor.Parent!.ReplaceChild(generated, anchor);
    }

    /// <summary>Writes a new size into <c>a:xfrm/a:ext</c>, and a free-floating shape's position into <c>a:off</c>.</summary>
    private static void WriteTransform(Xdr.Shape element, XLShape shape)
    {
        if (element.ShapeProperties?.Transform2D is not { } transform)
            return;

        var wb = shape.Worksheet.Workbook;
        if (shape.IsSizeChanged)
        {
            var extents = transform.Extents ?? transform.InsertAfter(new A.Extents(), transform.Offset);
            extents.Cx = DrawingUnits.PixelsToEmu(shape.Width, wb.DpiX);
            extents.Cy = DrawingUnits.PixelsToEmu(shape.Height, wb.DpiY);
        }

        if (shape.Placement == XLPicturePlacement.FreeFloating)
        {
            var offset = transform.Offset ?? transform.PrependChild(new A.Offset());
            offset.X = DrawingUnits.PixelsToEmu(shape.Left, wb.DpiX);
            offset.Y = DrawingUnits.PixelsToEmu(shape.Top, wb.DpiY);
        }
    }

    /// <summary>A body for a shape the file gave none, in its schema place after the style.</summary>
    private static Xdr.TextBody AddTextBody(Xdr.Shape element)
    {
        var body = new Xdr.TextBody(new A.BodyProperties(), new A.ListStyle());
        var before = (OpenXmlElement?)element.ShapeStyle ?? element.ShapeProperties;
        if (before is not null)
            element.InsertAfter(body, before);
        else
            element.Append(body);

        return body;
    }

    private static bool IsAnchorPosition(OpenXmlElement element) =>
        element is Xdr.FromMarker or Xdr.ToMarker or Xdr.Position or Xdr.Extent;
}
