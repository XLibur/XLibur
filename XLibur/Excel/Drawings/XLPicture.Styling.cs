using System;

namespace XLibur.Excel.Drawings;

/// <summary>The styling of a picture: border, fill, orientation and image effects (spec 17).</summary>
/// <remarks>
/// Every setter records its property in <see cref="AssignedStyle"/>, and a save writes only those,
/// into the picture's own <c>xdr:spPr</c> and <c>xdr:blipFill</c>. A property nobody assigned is
/// left as the file has it, whatever it holds.
/// </remarks>
internal sealed partial class XLPicture
{
    private const double MaxBorderWidthPt = 1584;

    /// <summary>The longest shadow blur and distance, in points, the EMU attributes can hold with room to spare.</summary>
    private const double MaxShadowLengthPt = 1584;

    /// <summary>How far a crop edge can extend the image outwards, in percent.</summary>
    private const double MinCropPercent = -1000;

    /// <summary>What the file has, as the loader read it, or what was assigned since.</summary>
    internal XLPictureStyle Style { get; private set; } = new();

    /// <summary>The styling properties assigned since the picture was loaded or last saved.</summary>
    internal XLPictureStyleProperties AssignedStyle { get; private set; }

    /// <summary>
    /// The <c>xdr:pic</c> the picture was loaded from, which a copy of the picture starts from, so
    /// the copy keeps what XLibur does not model. <c>null</c> for a picture XLibur made.
    /// </summary>
    internal string? LoadedPictureXml { get; set; }

    public XLColor? BorderColor
    {
        get => Style.BorderColor;
        set
        {
            Style.BorderColor = value;
            Assign(XLPictureStyleProperties.BorderColor);
            if (value is not null && Style.NoBorder)
            {
                Style.NoBorder = false;
                Assign(XLPictureStyleProperties.NoBorder);
            }
        }
    }

    public double? BorderWidthPt
    {
        get => Style.BorderWidthPt;
        set
        {
            if (value is { } width && (double.IsNaN(width) || width < 0 || width > MaxBorderWidthPt))
                throw new ArgumentOutOfRangeException(nameof(value), value,
                    $"A border is 0 to {MaxBorderWidthPt} points wide.");

            Style.BorderWidthPt = value;
            Assign(XLPictureStyleProperties.BorderWidth);
        }
    }

    public XLLineDashStyle? BorderDash
    {
        get => Style.BorderDash;
        set
        {
            if (value is { } dash && !Enum.IsDefined(dash))
                throw new ArgumentOutOfRangeException(nameof(value), value, "Unknown dash style.");

            Style.BorderDash = value;
            Assign(XLPictureStyleProperties.BorderDash);
        }
    }

    public bool NoBorder
    {
        get => Style.NoBorder;
        set
        {
            Style.NoBorder = value;
            Assign(XLPictureStyleProperties.NoBorder);
            if (value && Style.BorderColor is not null)
            {
                Style.BorderColor = null;
                Assign(XLPictureStyleProperties.BorderColor);
            }
        }
    }

    public XLColor? FillColor
    {
        get => Style.FillColor;
        set
        {
            Style.FillColor = value;
            Assign(XLPictureStyleProperties.FillColor);
            if (value is not null && Style.NoFill)
            {
                Style.NoFill = false;
                Assign(XLPictureStyleProperties.NoFill);
            }
        }
    }

    public bool NoFill
    {
        get => Style.NoFill;
        set
        {
            Style.NoFill = value;
            Assign(XLPictureStyleProperties.NoFill);
            if (value && Style.FillColor is not null)
            {
                Style.FillColor = null;
                Assign(XLPictureStyleProperties.FillColor);
            }
        }
    }

    public double? Rotation
    {
        get => Style.Rotation;
        set
        {
            if (value is { } degrees && !double.IsFinite(degrees))
                throw new ArgumentOutOfRangeException(nameof(value), value, "A rotation is a finite angle.");

            Style.Rotation = value is { } angle ? NormalizeAngle(angle) : null;
            Assign(XLPictureStyleProperties.Rotation);
        }
    }

    public bool? FlipHorizontal
    {
        get => Style.FlipHorizontal;
        set
        {
            Style.FlipHorizontal = value;
            Assign(XLPictureStyleProperties.FlipHorizontal);
        }
    }

    public bool? FlipVertical
    {
        get => Style.FlipVertical;
        set
        {
            Style.FlipVertical = value;
            Assign(XLPictureStyleProperties.FlipVertical);
        }
    }

    public double? Transparency
    {
        get => Style.Transparency;
        set
        {
            ThrowIfOutside(value, 0, 100, "A transparency is 0 to 100 percent.");
            Style.Transparency = value;
            Assign(XLPictureStyleProperties.Transparency);
        }
    }

    public XLPictureRecolor? Recolor
    {
        get => Style.Recolor;
        set
        {
            if (value == XLPictureRecolor.Custom)
                throw new ArgumentException(
                    "A custom recolour can be read and kept, but not assigned. Assign a preset instead.", nameof(value));

            if (value is { } recolor && !Enum.IsDefined(recolor))
                throw new ArgumentOutOfRangeException(nameof(value), value, "Unknown recolour.");

            // Washout is a brightness and a contrast, so it is assigned through them, and leaving it
            // takes them back to the image's own.
            if (value == XLPictureRecolor.Washout)
                SetLuminance(WashoutBrightness, WashoutContrast);
            else if (Style.Recolor == XLPictureRecolor.Washout)
                SetLuminance(null, null);

            Style.Recolor = value;
            Assign(XLPictureStyleProperties.Recolor);
        }
    }

    public double? Brightness
    {
        get => Style.Brightness;
        set
        {
            ThrowIfOutside(value, -100, 100, "A brightness is -100 to 100 percent.");
            Style.Brightness = value;
            Assign(XLPictureStyleProperties.Brightness);
            ReadRecolorFromLuminance();
        }
    }

    public double? Contrast
    {
        get => Style.Contrast;
        set
        {
            ThrowIfOutside(value, -100, 100, "A contrast is -100 to 100 percent.");
            Style.Contrast = value;
            Assign(XLPictureStyleProperties.Contrast);
            ReadRecolorFromLuminance();
        }
    }

    public XLCrop? Crop
    {
        get => Style.Crop;
        set
        {
            if (value is { } crop)
            {
                if (!double.IsFinite(crop.Left) || !double.IsFinite(crop.Top) ||
                    !double.IsFinite(crop.Right) || !double.IsFinite(crop.Bottom) ||
                    crop.Left >= 100 || crop.Top >= 100 || crop.Right >= 100 || crop.Bottom >= 100 ||
                    crop.Left < MinCropPercent || crop.Top < MinCropPercent ||
                    crop.Right < MinCropPercent || crop.Bottom < MinCropPercent ||
                    crop.Left + crop.Right >= 100 || crop.Top + crop.Bottom >= 100)
                    throw new ArgumentOutOfRangeException(nameof(value), value,
                        "A crop leaves part of the image: each edge, and each pair of opposite edges, under 100 percent, " +
                        $"and no edge below {MinCropPercent} percent.");
            }

            Style.Crop = value;
            Assign(XLPictureStyleProperties.Crop);
        }
    }

    public XLPictureShadow? Shadow
    {
        get => Style.Shadow;
        set
        {
            if (value is { } shadow)
            {
                ArgumentNullException.ThrowIfNull(shadow.Color, nameof(value));
                if (!double.IsFinite(shadow.TransparencyPct) || shadow.TransparencyPct < 0 || shadow.TransparencyPct > 100)
                    throw new ArgumentOutOfRangeException(nameof(value), value, "A shadow's transparency is 0 to 100 percent.");
                if (!double.IsFinite(shadow.BlurPt) || shadow.BlurPt < 0 || shadow.BlurPt > MaxShadowLengthPt ||
                    !double.IsFinite(shadow.DistancePt) || shadow.DistancePt < 0 || shadow.DistancePt > MaxShadowLengthPt ||
                    !double.IsFinite(shadow.AngleDeg))
                    throw new ArgumentOutOfRangeException(nameof(value), value,
                        $"A shadow's blur and distance are 0 to {MaxShadowLengthPt} points, and its angle is finite.");

                value = shadow with { AngleDeg = NormalizeAngle(shadow.AngleDeg) };
            }

            Style.Shadow = value;
            Assign(XLPictureStyleProperties.Shadow);
        }
    }

    internal const double WashoutBrightness = 70;
    internal const double WashoutContrast = -70;

    /// <summary>
    /// Give a copy of a picture its styling: the source's values, all assigned, so the copy is
    /// written with them, and the source's loaded <c>xdr:pic</c>, for what XLibur does not model.
    /// </summary>
    internal void TakeStyleOf(XLPicture source)
    {
        Style = source.Style.Clone();
        AssignedStyle = Style.Held();
        LoadedPictureXml = source.LoadedPictureXml;
    }

    private void Assign(XLPictureStyleProperties property) => AssignedStyle |= property;

    private void SetLuminance(double? brightness, double? contrast)
    {
        Style.Brightness = brightness;
        Style.Contrast = contrast;
        AssignedStyle |= XLPictureStyleProperties.Brightness | XLPictureStyleProperties.Contrast;
    }

    /// <summary>
    /// Keep <see cref="Recolor"/> reading what the file will: washout is a brightness and a
    /// contrast, so a picture with no other recolour is washed out exactly when it has those two,
    /// as the loader reads it.
    /// </summary>
    private void ReadRecolorFromLuminance()
    {
        var isWashout = Style.Brightness == WashoutBrightness && Style.Contrast == WashoutContrast;
        if (Style.Recolor == XLPictureRecolor.Washout && !isWashout)
        {
            Style.Recolor = XLPictureRecolor.None;
            AssignedStyle |= XLPictureStyleProperties.Recolor;
        }
        else if (Style.Recolor is null or XLPictureRecolor.None && isWashout)
        {
            // No element to write: the brightness and contrast are the recolour.
            Style.Recolor = XLPictureRecolor.Washout;
        }
    }

    private static void ThrowIfOutside(double? value, double min, double max, string message)
    {
        if (value is { } v && (double.IsNaN(v) || v < min || v > max))
            throw new ArgumentOutOfRangeException(nameof(value), value, message);
    }

    /// <summary>An angle brought into 0 up to 360 degrees.</summary>
    private static double NormalizeAngle(double degrees)
    {
        var angle = degrees % 360;
        if (angle < 0)
            angle += 360;

        // A tiny negative angle rounds up to a whole turn.
        return angle >= 360 ? 0 : angle;
    }
}
