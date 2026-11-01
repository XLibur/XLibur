using System.Linq;
using XLibur.Excel.Drawings;
using XLibur.Excel.IO.DrawingML;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// Reads the styling of a loaded <c>xdr:pic</c> into the picture's model (spec 17). Only the simple
/// forms the model holds are read; anything else reads as no value and stays in the file as it is.
/// </summary>
internal static class PictureStyleReader
{
    internal static void Read(Xdr.Picture picture, XLPictureStyle style)
    {
        if (picture.ShapeProperties is { } shapeProperties)
            ReadShapeProperties(shapeProperties, style);

        if (picture.BlipFill is { } blipFill)
        {
            if (blipFill.Blip is { } blip)
            {
                style.Transparency = BlipEffectsWriter.ReadTransparency(blip);
                style.Recolor = BlipEffectsWriter.ReadRecolor(blip);
                (style.Brightness, style.Contrast) = BlipEffectsWriter.ReadLuminance(blip);
            }

            style.Crop = BlipEffectsWriter.ReadCrop(blipFill);
        }
    }

    private static void ReadShapeProperties(Xdr.ShapeProperties shapeProperties, XLPictureStyle style)
    {
        if (shapeProperties.Transform2D is { } transform)
        {
            style.Rotation = transform.Rotation?.Value is { } rotation and not 0
                ? DrawingUnits.AngleToDegrees(rotation)
                : null;
            style.FlipHorizontal = transform.HorizontalFlip?.Value;
            style.FlipVertical = transform.VerticalFlip?.Value;
        }

        style.NoFill = shapeProperties.Elements<A.NoFill>().Any();
        style.FillColor = DrawingColorReader.ReadSolidFillColor(shapeProperties);

        if (shapeProperties.Elements<A.Outline>().FirstOrDefault() is { } outline)
        {
            style.NoBorder = outline.Elements<A.NoFill>().Any();
            style.BorderColor = DrawingColorReader.ReadSolidFillColor(outline);
            style.BorderWidthPt = outline.Width?.Value is { } width ? DrawingUnits.EmuToPoints(width) : null;
            style.BorderDash = outline.Elements<A.PresetDash>().FirstOrDefault()?.Val is { } dash
                ? ShapePropertiesWriter.FromPresetDash(dash.Value)
                : null;
        }

        style.Shadow = ReadShadow(shapeProperties);
    }

    /// <summary>The shadow, when the effects are one outer shadow and nothing else.</summary>
    private static XLPictureShadow? ReadShadow(Xdr.ShapeProperties shapeProperties)
    {
        if (shapeProperties.Elements<A.EffectList>().FirstOrDefault() is not { } effects ||
            effects.ChildElements.ToList() is not [A.OuterShadow shadow])
            return null;

        var color = shadow.ChildElements.FirstOrDefault();
        return new XLPictureShadow(
            DrawingColorReader.ReadColor(color) ?? XLColor.Black,
            DrawingColorReader.ReadTransparency(color),
            DrawingUnits.EmuToPoints(shadow.BlurRadius?.Value ?? 0),
            DrawingUnits.EmuToPoints(shadow.Distance?.Value ?? 0),
            DrawingUnits.AngleToDegrees(shadow.Direction?.Value ?? 0));
    }
}
