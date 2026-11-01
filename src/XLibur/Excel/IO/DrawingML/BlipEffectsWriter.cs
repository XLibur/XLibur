using System;
using System.Linq;
using DocumentFormat.OpenXml;
using XLibur.Excel.Drawings;
using A = DocumentFormat.OpenXml.Drawing;

namespace XLibur.Excel.IO.DrawingML;

/// <summary>
/// Writes into, and reads, a picture's <c>blipFill</c>: the effects on its <c>a:blip</c> and the
/// crop in <c>a:srcRect</c>. The <c>blipFill</c> counterpart of <see cref="ShapePropertiesWriter"/>.
/// </summary>
/// <remarks>
/// <para>
/// The children of <c>a:blip</c> are a list of effects applied in order, then <c>a:extLst</c>. The
/// order matters to the image, so an effect that is there is edited where it is, and a new one goes
/// where Excel writes it: the transparency (<c>a:alphaModFix</c>), then the recolour, then the
/// brightness and contrast (<c>a:lum</c>), all before <c>a:extLst</c>. Effects XLibur does not model
/// keep their places.
/// </para>
/// <para>
/// The recolour presets are fixed effect patterns, taken from files Excel wrote: grayscale is
/// <c>a:grayscl</c>, black and white is <c>a:biLevel thresh="50000"</c>, and washout is no recolour
/// element at all but <c>a:lum bright="70000" contrast="-70000"</c>. Sepia is the duotone below,
/// which the current Excel object model cannot apply, so it is the pattern Office's gallery writes
/// as documented rather than one captured from a file.
/// </para>
/// </remarks>
internal static class BlipEffectsWriter
{
    private const int BlackAndWhiteThreshold = 50000;
    private const string SepiaColor = "D9C3A5";
    private const int SepiaTint = 50000;
    private const int SepiaSaturation = 180000;

    // ── Transparency ────────────────────────────────────────────────────

    /// <summary>
    /// Sets how transparent the image is, 0 to 100 percent, or removes the transparency when it is
    /// <c>null</c> or 0.
    /// </summary>
    internal static void SetTransparency(A.Blip blip, double? transparency)
    {
        var existing = blip.Elements<A.AlphaModulationFixed>().FirstOrDefault();
        if (transparency is not { } value || value == 0)
        {
            existing?.Remove();
            return;
        }

        if (existing is null)
        {
            existing = new A.AlphaModulationFixed();
            InsertBefore(blip, existing, e => IsRecolor(e) || e is A.LuminanceEffect or A.BlipExtensionList);
        }

        existing.Amount = DrawingUnits.PercentToThousandths(100 - value);
    }

    /// <summary>How transparent the image is, from <c>a:alphaModFix</c>, or <c>null</c> when it has none.</summary>
    internal static double? ReadTransparency(A.Blip blip) =>
        blip.Elements<A.AlphaModulationFixed>().FirstOrDefault() is { } alpha
            ? 100 - (alpha.Amount?.Value ?? 100000) / 1000.0
            : null;

    // ── Recolour ────────────────────────────────────────────────────────

    /// <summary>
    /// Replaces the recolour with a preset, or removes it for <see cref="XLPictureRecolor.None"/>,
    /// <see cref="XLPictureRecolor.Washout"/> and <c>null</c>. Washout is a brightness and a contrast,
    /// which <see cref="SetLuminance"/> writes.
    /// </summary>
    internal static void SetRecolor(A.Blip blip, XLPictureRecolor? recolor)
    {
        var existing = blip.ChildElements.Where(IsRecolor).ToList();
        var place = existing.Count > 0 ? existing[^1].NextSibling() : null;
        foreach (var element in existing)
            element.Remove();

        OpenXmlElement? preset = recolor switch
        {
            XLPictureRecolor.Grayscale => new A.Grayscale(),
            XLPictureRecolor.BlackAndWhite => new A.BiLevel { Threshold = BlackAndWhiteThreshold },
            XLPictureRecolor.Sepia => new A.Duotone(
                new A.PresetColor { Val = A.PresetColorValues.Black },
                new A.RgbColorModelHex(
                    new A.Tint { Val = SepiaTint },
                    new A.SaturationModulation { Val = SepiaSaturation }) { Val = SepiaColor }),
            XLPictureRecolor.None or XLPictureRecolor.Washout or null => null,
            _ => throw new ArgumentOutOfRangeException(nameof(recolor), recolor, "A recolour that cannot be written."),
        };

        if (preset is null)
            return;

        if (place is not null)
            blip.InsertBefore(preset, place);
        else
            InsertBefore(blip, preset, e => e is A.LuminanceEffect or A.BlipExtensionList);
    }

    /// <summary>
    /// The recolour the blip has: a preset when its effects are exactly one, <see cref="XLPictureRecolor.Custom"/>
    /// for any other recolour, and <c>null</c> for none.
    /// </summary>
    internal static XLPictureRecolor? ReadRecolor(A.Blip blip)
    {
        var recolors = blip.ChildElements.Where(IsRecolor).ToList();
        if (recolors.Count == 0)
        {
            var (bright, contrast) = ReadLuminance(blip);
            return bright == XLPicture.WashoutBrightness && contrast == XLPicture.WashoutContrast
                ? XLPictureRecolor.Washout
                : null;
        }

        if (recolors.Count > 1)
            return XLPictureRecolor.Custom;

        return recolors[0] switch
        {
            A.Grayscale { HasChildren: false } => XLPictureRecolor.Grayscale,
            A.BiLevel { Threshold.Value: BlackAndWhiteThreshold } => XLPictureRecolor.BlackAndWhite,
            A.Duotone duotone when IsSepia(duotone) => XLPictureRecolor.Sepia,
            _ => XLPictureRecolor.Custom,
        };
    }

    private static bool IsSepia(A.Duotone duotone)
    {
        var colors = duotone.ChildElements.ToList();
        return colors is [A.PresetColor { HasChildren: false } dark, A.RgbColorModelHex light] &&
               dark.Val?.Value == A.PresetColorValues.Black &&
               string.Equals(light.Val?.Value, SepiaColor, StringComparison.OrdinalIgnoreCase) &&
               light.ChildElements.ToList() is [A.Tint { Val.Value: SepiaTint }, A.SaturationModulation { Val.Value: SepiaSaturation }];
    }

    /// <summary>The effects that recolour the image, of which the recolour presets are made.</summary>
    private static bool IsRecolor(OpenXmlElement element) => element is A.Grayscale or A.BiLevel or A.Duotone;

    // ── Brightness and contrast ─────────────────────────────────────────

    /// <summary>
    /// Writes the brightness and contrast that <paramref name="writeBrightness"/> and
    /// <paramref name="writeContrast"/> say to, each removed when <c>null</c> or 0, into the blip's
    /// <c>a:lum</c>. An <c>a:lum</c> left with nothing is removed.
    /// </summary>
    internal static void SetLuminance(A.Blip blip, bool writeBrightness, double? brightness, bool writeContrast,
        double? contrast)
    {
        var lum = blip.Elements<A.LuminanceEffect>().FirstOrDefault();
        if (lum is null)
        {
            if (!(writeBrightness && brightness is { } b && b != 0) && !(writeContrast && contrast is { } c && c != 0))
                return;

            lum = new A.LuminanceEffect();
            InsertBefore(blip, lum, e => e is A.BlipExtensionList);
        }

        if (writeBrightness)
            lum.Brightness = brightness is { } bright && bright != 0 ? DrawingUnits.PercentToThousandths(bright) : null;

        if (writeContrast)
            lum.Contrast = contrast is { } con && con != 0 ? DrawingUnits.PercentToThousandths(con) : null;

        if (!lum.HasAttributes)
            lum.Remove();
    }

    /// <summary>The brightness and contrast, from <c>a:lum</c>; <c>null</c> for each it does not set.</summary>
    internal static (double? Brightness, double? Contrast) ReadLuminance(A.Blip blip)
    {
        var lum = blip.Elements<A.LuminanceEffect>().FirstOrDefault();
        return (lum?.Brightness?.Value / 1000.0, lum?.Contrast?.Value / 1000.0);
    }

    // ── Crop ────────────────────────────────────────────────────────────

    /// <summary>
    /// Sets the crop on <c>a:srcRect</c>, or removes it when <paramref name="crop"/> is <c>null</c>.
    /// An edge that is not cut is not written, as Excel writes it. The fill mode after it, a stretch
    /// or a tile, is untouched.
    /// </summary>
    internal static void SetCrop(OpenXmlCompositeElement blipFill, XLCrop? crop)
    {
        var sourceRectangle = blipFill.Elements<A.SourceRectangle>().FirstOrDefault();
        if (crop is not { } value)
        {
            if (sourceRectangle is not null)
            {
                sourceRectangle.Left = null;
                sourceRectangle.Top = null;
                sourceRectangle.Right = null;
                sourceRectangle.Bottom = null;
            }

            return;
        }

        if (sourceRectangle is null)
        {
            // CT_BlipFillProperties is a sequence: a:blip, a:srcRect, then the fill mode.
            sourceRectangle = new A.SourceRectangle();
            if (blipFill.Elements<A.Blip>().FirstOrDefault() is { } blip)
                blipFill.InsertAfter(sourceRectangle, blip);
            else
                blipFill.InsertAt(sourceRectangle, 0);
        }

        sourceRectangle.Left = Edge(value.Left);
        sourceRectangle.Top = Edge(value.Top);
        sourceRectangle.Right = Edge(value.Right);
        sourceRectangle.Bottom = Edge(value.Bottom);

        static Int32Value? Edge(double percent) =>
            DrawingUnits.PercentToThousandths(percent) is var edge && edge != 0 ? edge : null;
    }

    /// <summary>The crop, from <c>a:srcRect</c>, or <c>null</c> when nothing is cut.</summary>
    internal static XLCrop? ReadCrop(OpenXmlCompositeElement blipFill)
    {
        if (blipFill.Elements<A.SourceRectangle>().FirstOrDefault() is not { } rectangle)
            return null;

        var crop = new XLCrop(
            (rectangle.Left?.Value ?? 0) / 1000.0,
            (rectangle.Top?.Value ?? 0) / 1000.0,
            (rectangle.Right?.Value ?? 0) / 1000.0,
            (rectangle.Bottom?.Value ?? 0) / 1000.0);
        return crop == default ? null : crop;
    }

    /// <summary>Inserts <paramref name="element"/> before the first child that must follow it, or last.</summary>
    private static void InsertBefore(OpenXmlCompositeElement parent, OpenXmlElement element, Func<OpenXmlElement, bool> follows)
    {
        var anchor = parent.ChildElements.FirstOrDefault(follows);
        if (anchor != null)
            parent.InsertBefore(element, anchor);
        else
            parent.Append(element);
    }
}
