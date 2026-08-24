using System;

namespace XLibur.Excel.IO.DrawingML;

/// <summary>
/// Converts the screen-pixel lengths the drawing model holds into the English Metric Units
/// DrawingML writes. There are 914400 EMU to the inch, which is why the conversion needs the
/// resolution the pixel count was measured at.
/// </summary>
/// <remarks>
/// <para>
/// Not to be confused with <see cref="Coordinates.Emu"/>, which is a length <em>value</em> — a
/// number with a preferred unit — and converts between inches, points, picas and the rest. It has
/// no notion of pixels or of a resolution, and it rounds away from zero rather than to even, so the
/// two are not interchangeable and neither can be expressed in terms of the other.
/// </para>
/// <para>
/// See http://polymathprogrammer.com/2009/10/22/english-metric-units-and-open-xml/,
/// http://archive.oreilly.com/pub/post/what_is_an_emu.html and
/// https://en.wikipedia.org/wiki/Office_Open_XML_file_formats#DrawingML.
/// </para>
/// </remarks>
internal static class DrawingUnits
{
    private const long EmuPerInch = 914400L;

    /// <summary>
    /// EMU per point, the unit DrawingML states line widths in. There are 72 points to the inch.
    /// </summary>
    internal const double EmuPerPoint = 12700;

    /// <summary>Converts a pixel length, measured at the given resolution, to EMU.</summary>
    /// <remarks>
    /// The rounding is <see cref="Convert.ToInt64(double)"/>'s — to even — and not a cast, which
    /// would truncate. Every extent and offset XLibur has ever written came out of this exact
    /// expression, so the two are not interchangeable.
    /// </remarks>
    internal static long PixelsToEmu(int pixels, double resolution) =>
        Convert.ToInt64(EmuPerInch * pixels / resolution);

    /// <summary>Converts a length in points to EMU, for the attributes that are stated that way.</summary>
    /// <remarks>
    /// <see cref="Math.Round(double)"/> then a cast, which is what every line width XLibur has
    /// written came out of. The result is <see cref="int"/> because that is what <c>a:ln/@w</c> is.
    /// </remarks>
    internal static int PointsToEmu(double points) => (int)Math.Round(points * EmuPerPoint);
}
