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
            WriteGeometry(anchor, element, shape);

        if (shape.IsNameChanged && element.NonVisualShapeProperties?.NonVisualDrawingProperties is { } properties)
            properties.Name = shape.Name;

        if (shape.Assigned != XLShapeProperties.None)
            ShapeFormatWriter.Write(element.ShapeProperties ??= new Xdr.ShapeProperties(), shape);

        if (shape.Body.HasChanges)
        {
            var body = element.TextBody ?? AddTextBody(element);
            ShapeTextWriter.Write(body, shape.Body, alignNewParagraphs: false);
        }
    }

    /// <summary>
    /// Puts the model's position into the anchor: into its own markers, position and extent when the
    /// placement kept its form, keeping every attribute and child it has besides, such as
    /// <c>editAs</c>. A new placement is a new form of anchor, and the shape moves into it with the
    /// client data, as a picture does.
    /// </summary>
    private static void WriteGeometry(OpenXmlCompositeElement anchor, Xdr.Shape element, XLShape shape)
    {
        var generated = DrawingAnchorFactory.Create(shape.Placement, ShapeWriter.GeometryOf(shape), new Xdr.Shape());
        var position = generated.ChildElements.Where(IsAnchorPosition).ToList();
        foreach (var part in position)
            part.Remove();

        if (generated.GetType() == anchor.GetType())
        {
            // Only the parts whose model value changed: the rest stays as the file has it, rather
            // than coming back from pixels a half pixel off.
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
        else
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
