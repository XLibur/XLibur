using System;

namespace XLibur.Excel.Drawings;

/// <summary>The styling properties of a picture, one flag each, for which ones were assigned.</summary>
[Flags]
internal enum XLPictureStyleProperties
{
    None = 0,
    BorderColor = 1 << 0,
    BorderWidth = 1 << 1,
    BorderDash = 1 << 2,
    NoBorder = 1 << 3,
    FillColor = 1 << 4,
    NoFill = 1 << 5,
    Rotation = 1 << 6,
    FlipHorizontal = 1 << 7,
    FlipVertical = 1 << 8,
    Transparency = 1 << 9,
    Recolor = 1 << 10,
    Brightness = 1 << 11,
    Contrast = 1 << 12,
    Crop = 1 << 13,
    Shadow = 1 << 14,

    /// <summary>What goes into the picture's <c>xdr:spPr</c>.</summary>
    ShapeProperties = BorderColor | BorderWidth | BorderDash | NoBorder | FillColor | NoFill |
                      Rotation | FlipHorizontal | FlipVertical | Shadow,

    /// <summary>What goes into the picture's <c>xdr:blipFill</c>.</summary>
    BlipFill = Transparency | Recolor | Brightness | Contrast | Crop,
}

/// <summary>
/// The styling of a picture: what the file has, as the loader read it, or what was assigned since.
/// </summary>
/// <remarks>
/// A plain bag of values. Whether a value is written is decided by the picture's assigned flags,
/// not by this, so a value read from the file is never written back unless it is assigned.
/// </remarks>
internal sealed class XLPictureStyle
{
    internal XLColor? BorderColor { get; set; }
    internal double? BorderWidthPt { get; set; }
    internal XLLineDashStyle? BorderDash { get; set; }
    internal bool NoBorder { get; set; }
    internal XLColor? FillColor { get; set; }
    internal bool NoFill { get; set; }
    internal double? Rotation { get; set; }
    internal bool? FlipHorizontal { get; set; }
    internal bool? FlipVertical { get; set; }
    internal double? Transparency { get; set; }
    internal XLPictureRecolor? Recolor { get; set; }
    internal double? Brightness { get; set; }
    internal double? Contrast { get; set; }
    internal XLCrop? Crop { get; set; }
    internal XLPictureShadow? Shadow { get; set; }

    /// <summary>The properties that hold a value, or say there is none: what a copy of the picture assigns.</summary>
    internal XLPictureStyleProperties Held()
    {
        return If(BorderColor is not null, XLPictureStyleProperties.BorderColor)
               | If(BorderWidthPt is not null, XLPictureStyleProperties.BorderWidth)
               | If(BorderDash is not null, XLPictureStyleProperties.BorderDash)
               | If(NoBorder, XLPictureStyleProperties.NoBorder)
               | If(FillColor is not null, XLPictureStyleProperties.FillColor)
               | If(NoFill, XLPictureStyleProperties.NoFill)
               | If(Rotation is not null, XLPictureStyleProperties.Rotation)
               | If(FlipHorizontal is not null, XLPictureStyleProperties.FlipHorizontal)
               | If(FlipVertical is not null, XLPictureStyleProperties.FlipVertical)
               | If(Transparency is not null, XLPictureStyleProperties.Transparency)
               | If(Recolor is not null and not XLPictureRecolor.Custom, XLPictureStyleProperties.Recolor)
               | If(Brightness is not null, XLPictureStyleProperties.Brightness)
               | If(Contrast is not null, XLPictureStyleProperties.Contrast)
               | If(Crop is not null, XLPictureStyleProperties.Crop)
               | If(Shadow is not null, XLPictureStyleProperties.Shadow);

        static XLPictureStyleProperties If(bool holds, XLPictureStyleProperties property) =>
            holds ? property : XLPictureStyleProperties.None;
    }

    internal XLPictureStyle Clone() => (XLPictureStyle)MemberwiseClone();
}
