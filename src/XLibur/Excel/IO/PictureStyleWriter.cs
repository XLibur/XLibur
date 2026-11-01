using System.Linq;
using XLibur.Excel.Drawings;
using XLibur.Excel.IO.DrawingML;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// Writes the styling properties assigned to a picture into its <c>xdr:pic</c>, and nothing else:
/// a property nobody assigned is left as the file has it (spec 17).
/// </summary>
/// <remarks>
/// This decides whether to write each property; how is the shared layer's,
/// <see cref="ShapePropertiesWriter"/> for <c>xdr:spPr</c> and <see cref="BlipEffectsWriter"/> for
/// <c>xdr:blipFill</c>. The same writes serve a new picture, a loaded one and a grouped one.
/// </remarks>
internal static class PictureStyleWriter
{
    internal static void Write(Xdr.Picture picture, XLPicture pic)
    {
        var assigned = pic.AssignedStyle;
        if (assigned == XLPictureStyleProperties.None)
            return;

        var style = pic.Style;
        if ((assigned & XLPictureStyleProperties.ShapeProperties) != 0)
            WriteShapeProperties(picture.ShapeProperties ??= new Xdr.ShapeProperties(), style, assigned);

        if ((assigned & XLPictureStyleProperties.BlipFill) != 0 && picture.BlipFill is { } blipFill)
            WriteBlipFill(blipFill, style, assigned);
    }

    private static void WriteShapeProperties(Xdr.ShapeProperties shapeProperties, XLPictureStyle style,
        XLPictureStyleProperties assigned)
    {
        if (Has(assigned, XLPictureStyleProperties.Rotation | XLPictureStyleProperties.FlipHorizontal |
                          XLPictureStyleProperties.FlipVertical))
        {
            var transform = shapeProperties.Transform2D ?? shapeProperties.PrependChild(new A.Transform2D());
            if (Has(assigned, XLPictureStyleProperties.Rotation))
                ShapePropertiesWriter.SetRotation(transform, style.Rotation);
            if (Has(assigned, XLPictureStyleProperties.FlipHorizontal))
                ShapePropertiesWriter.SetFlip(transform, horizontal: true, style.FlipHorizontal);
            if (Has(assigned, XLPictureStyleProperties.FlipVertical))
                ShapePropertiesWriter.SetFlip(transform, horizontal: false, style.FlipVertical);
        }

        if (Has(assigned, XLPictureStyleProperties.NoFill) && style.NoFill)
            ShapePropertiesWriter.SetNoFill(shapeProperties);
        else if (Has(assigned, XLPictureStyleProperties.FillColor))
            ShapePropertiesWriter.SetFill(shapeProperties, style.FillColor);
        else if (Has(assigned, XLPictureStyleProperties.NoFill))
            ShapePropertiesWriter.RemoveNoFill(shapeProperties);

        WriteBorder(shapeProperties, style, assigned);

        if (Has(assigned, XLPictureStyleProperties.Shadow))
            ShapePropertiesWriter.SetOuterShadow(shapeProperties, style.Shadow);
    }

    private static void WriteBorder(Xdr.ShapeProperties shapeProperties, XLPictureStyle style,
        XLPictureStyleProperties assigned)
    {
        const XLPictureStyleProperties border = XLPictureStyleProperties.BorderColor |
                                                XLPictureStyleProperties.BorderWidth |
                                                XLPictureStyleProperties.BorderDash |
                                                XLPictureStyleProperties.NoBorder;
        if (!Has(assigned, border))
            return;

        // An outline is added only for a value to put in it; one that only loses values is edited
        // if it is there.
        var wantsOutline = (Has(assigned, XLPictureStyleProperties.BorderColor) && style.BorderColor is not null) ||
                           (Has(assigned, XLPictureStyleProperties.BorderWidth) && style.BorderWidthPt is not null) ||
                           (Has(assigned, XLPictureStyleProperties.BorderDash) && style.BorderDash is not null) ||
                           (Has(assigned, XLPictureStyleProperties.NoBorder) && style.NoBorder);
        var outline = wantsOutline
            ? ShapePropertiesWriter.EnsureOutline(shapeProperties)
            : shapeProperties.Elements<A.Outline>().FirstOrDefault();
        if (outline is null)
            return;

        if (Has(assigned, XLPictureStyleProperties.NoBorder) && style.NoBorder)
            ShapePropertiesWriter.SetOutlineNoFill(outline);
        else if (Has(assigned, XLPictureStyleProperties.BorderColor))
            ShapePropertiesWriter.SetOutlineColor(outline, style.BorderColor);
        else if (Has(assigned, XLPictureStyleProperties.NoBorder))
            ShapePropertiesWriter.RemoveOutlineNoFill(outline);

        if (Has(assigned, XLPictureStyleProperties.BorderWidth))
            ShapePropertiesWriter.SetOutlineWidth(outline, style.BorderWidthPt);

        if (Has(assigned, XLPictureStyleProperties.BorderDash))
            ShapePropertiesWriter.SetOutlineDash(outline, style.BorderDash);
    }

    private static void WriteBlipFill(Xdr.BlipFill blipFill, XLPictureStyle style, XLPictureStyleProperties assigned)
    {
        if (blipFill.Blip is { } blip)
        {
            if (Has(assigned, XLPictureStyleProperties.Transparency))
                BlipEffectsWriter.SetTransparency(blip, style.Transparency);

            if (Has(assigned, XLPictureStyleProperties.Recolor))
                BlipEffectsWriter.SetRecolor(blip, style.Recolor);

            var brightness = Has(assigned, XLPictureStyleProperties.Brightness);
            var contrast = Has(assigned, XLPictureStyleProperties.Contrast);
            if (brightness || contrast)
                BlipEffectsWriter.SetLuminance(blip, brightness, style.Brightness, contrast, style.Contrast);
        }

        if (Has(assigned, XLPictureStyleProperties.Crop))
            BlipEffectsWriter.SetCrop(blipFill, style.Crop);
    }

    private static bool Has(XLPictureStyleProperties assigned, XLPictureStyleProperties any) => (assigned & any) != 0;
}
