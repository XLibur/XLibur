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
        var held = XLPictureStyleProperties.None;
        if (BorderColor is not null) held |= XLPictureStyleProperties.BorderColor;
        if (BorderWidthPt is not null) held |= XLPictureStyleProperties.BorderWidth;
        if (BorderDash is not null) held |= XLPictureStyleProperties.BorderDash;
        if (NoBorder) held |= XLPictureStyleProperties.NoBorder;
        if (FillColor is not null) held |= XLPictureStyleProperties.FillColor;
        if (NoFill) held |= XLPictureStyleProperties.NoFill;
        if (Rotation is not null) held |= XLPictureStyleProperties.Rotation;
        if (FlipHorizontal is not null) held |= XLPictureStyleProperties.FlipHorizontal;
        if (FlipVertical is not null) held |= XLPictureStyleProperties.FlipVertical;
        if (Transparency is not null) held |= XLPictureStyleProperties.Transparency;
        if (Recolor is not null and not XLPictureRecolor.Custom) held |= XLPictureStyleProperties.Recolor;
        if (Brightness is not null) held |= XLPictureStyleProperties.Brightness;
        if (Contrast is not null) held |= XLPictureStyleProperties.Contrast;
        if (Crop is not null) held |= XLPictureStyleProperties.Crop;
        if (Shadow is not null) held |= XLPictureStyleProperties.Shadow;
        return held;
    }

    internal XLPictureStyle Clone() => (XLPictureStyle)MemberwiseClone();
}
