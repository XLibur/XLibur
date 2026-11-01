namespace XLibur.Excel.Drawings;

/// <summary>
/// The dash pattern of a DrawingML line, one member per <c>a:prstDash</c> value.
/// </summary>
/// <remarks>
/// <see cref="XLDashStyle"/> is the VML set that comment shapes use. DrawingML has its own, and a
/// mapping between the two would lose values both ways.
/// </remarks>
public enum XLLineDashStyle
{
    /// <summary><c>solid</c></summary>
    Solid,

    /// <summary><c>dot</c></summary>
    Dot,

    /// <summary><c>dash</c></summary>
    Dash,

    /// <summary><c>lgDash</c></summary>
    LongDash,

    /// <summary><c>dashDot</c></summary>
    DashDot,

    /// <summary><c>lgDashDot</c></summary>
    LongDashDot,

    /// <summary><c>lgDashDotDot</c></summary>
    LongDashDotDot,

    /// <summary><c>sysDash</c></summary>
    SystemDash,

    /// <summary><c>sysDot</c></summary>
    SystemDot,

    /// <summary><c>sysDashDot</c></summary>
    SystemDashDot,

    /// <summary><c>sysDashDotDot</c></summary>
    SystemDashDotDot,
}

/// <summary>
/// A recolour preset from Excel's Picture Format ▸ Color gallery.
/// </summary>
public enum XLPictureRecolor
{
    /// <summary>The image's own colours.</summary>
    None,

    /// <summary>Shades of grey: <c>a:grayscl</c>.</summary>
    Grayscale,

    /// <summary>Brown tones: an <c>a:duotone</c> from black to a light sepia.</summary>
    Sepia,

    /// <summary>Faded, for a picture behind text: <c>a:lum bright="70000" contrast="-70000"</c>.</summary>
    Washout,

    /// <summary>Black and white only: <c>a:biLevel thresh="50000"</c>.</summary>
    BlackAndWhite,

    /// <summary>
    /// A recolour the file has that is none of the presets, such as a duotone of other colours.
    /// It is read back and kept, but cannot be assigned.
    /// </summary>
    Custom,
}

/// <summary>
/// How much of each edge of the image is cut away, in percent of the image's width or height.
/// </summary>
/// <param name="Left">Percent of the width cut from the left edge.</param>
/// <param name="Top">Percent of the height cut from the top edge.</param>
/// <param name="Right">Percent of the width cut from the right edge.</param>
/// <param name="Bottom">Percent of the height cut from the bottom edge.</param>
/// <remarks>
/// Written as <c>a:srcRect</c>, in thousandths of a percent. A negative value adds space around
/// the image instead.
/// </remarks>
public readonly record struct XLCrop(double Left, double Top, double Right, double Bottom);

/// <summary>
/// A shadow outside the picture, as Excel's shadow presets are: <c>a:outerShdw</c>.
/// </summary>
/// <param name="Color">The colour of the shadow.</param>
/// <param name="TransparencyPct">How transparent the shadow is, 0 (opaque) to 100.</param>
/// <param name="BlurPt">The blur radius, in points.</param>
/// <param name="DistancePt">How far the shadow falls from the picture, in points.</param>
/// <param name="AngleDeg">The direction the shadow falls in, in degrees clockwise from the right.</param>
public readonly record struct XLPictureShadow(
    XLColor Color,
    double TransparencyPct,
    double BlurPt,
    double DistancePt,
    double AngleDeg);
