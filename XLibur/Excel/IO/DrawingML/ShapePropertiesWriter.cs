using System.Linq;
using DocumentFormat.OpenXml;
using XLibur.Extensions;
using A = DocumentFormat.OpenXml.Drawing;

namespace XLibur.Excel.IO.DrawingML;

/// <summary>
/// Writes into <c>a:CT_ShapeProperties</c> — the fill, the outline and the colours they carry.
/// </summary>
/// <remarks>
/// <para>
/// A chart series' <c>c:spPr</c> and a shape's <c>xdr:sp/spPr</c> are the same schema type, so the
/// rules for writing into one are the rules for writing into the other. They are subtle enough that
/// a second implementation would be a second set of bugs, which is why this takes an
/// <see cref="OpenXmlCompositeElement"/> rather than any one host's element type.
/// </para>
/// <para>Three facts about the schema shape everything here:</para>
/// <list type="number">
/// <item>
/// The fills are a <b>choice</b> group. Setting a colour over an <c>a:gradFill</c> means removing
/// the whole group first, not appending a second fill beside it.
/// </item>
/// <item>
/// The type is a <b>sequence</b> — <c>a:xfrm</c>, geometry, fill, <c>a:ln</c>, <c>a:effectLst</c>,
/// 3-D, <c>a:extLst</c> — and the SDK does not order the children of an element built by hand. Every
/// insertion here goes in front of the first child that must follow it.
/// </item>
/// <item>
/// Existing elements are <b>mutated in place, never rebuilt</b>, so children XLibur does not model —
/// an <c>a:round</c> join, a <c>cap</c>, arrowheads — survive an edit that never mentions them.
/// </item>
/// </list>
/// <para>
/// Every operation takes plain values. Whether to write at all is the caller's decision and stays
/// with the caller; what the schema then requires is this layer's. A <c>null</c> colour means "no
/// explicit colour", which removes the element rather than writing black — the model convention
/// throughout XLibur.
/// </para>
/// </remarks>
internal static class ShapePropertiesWriter
{
    /// <summary>
    /// Sets the fill to a solid <paramref name="color"/>, or removes the fill entirely when the
    /// colour is <c>null</c> or has no value.
    /// </summary>
    /// <remarks>
    /// Whatever fill was there goes first either way, because the fills are one choice group: a
    /// solid fill appended beside a surviving gradient would be schema-invalid and would render as
    /// the gradient.
    /// </remarks>
    internal static void SetFill(OpenXmlCompositeElement shapeProperties, XLColor? color)
    {
        foreach (var existing in shapeProperties.ChildElements.Where(IsFillElement).ToList())
            existing.Remove();

        if (color == null || !color.HasValue)
            return;

        // The fill precedes the outline and every effect in CT_ShapeProperties.
        var anchor = shapeProperties.ChildElements.FirstOrDefault(IsAfterFill);
        var fill = new A.SolidFill(BuildColor(color));
        if (anchor != null)
            shapeProperties.InsertBefore(fill, anchor);
        else
            shapeProperties.Append(fill);
    }

    /// <summary>
    /// Returns the <c>a:ln</c>, adding an empty one in its schema position if there is none.
    /// </summary>
    /// <remarks>
    /// Deciding whether an outline is wanted at all belongs to the caller: an element added here for
    /// properties that all turn out to be absent would read back as "this shape has an explicit
    /// outline". Callers that may have nothing to write check first and do not call.
    /// </remarks>
    internal static A.Outline EnsureOutline(OpenXmlCompositeElement shapeProperties)
    {
        var existing = shapeProperties.Elements<A.Outline>().FirstOrDefault();
        if (existing != null)
            return existing;

        var outline = new A.Outline();
        var anchor = shapeProperties.ChildElements.FirstOrDefault(IsAfterOutline);
        if (anchor != null)
            shapeProperties.InsertBefore(outline, anchor);
        else
            shapeProperties.Append(outline);

        return outline;
    }

    /// <summary>
    /// Sets the outline width in points, or removes the width when <paramref name="widthPoints"/> is
    /// <c>null</c>, leaving everything else about the outline alone.
    /// </summary>
    internal static void SetOutlineWidth(A.Outline outline, double? widthPoints)
    {
        if (widthPoints == null)
            outline.Width = null;
        else
            outline.Width = DrawingUnits.PointsToEmu(widthPoints.Value);
    }

    /// <summary>
    /// Sets the outline's fill to a solid <paramref name="color"/>, or removes it when the colour is
    /// <c>null</c> or has no value. The outline's width, cap, join and ends are untouched.
    /// </summary>
    /// <remarks>
    /// The fill goes first in <c>a:ln</c>, ahead of the dash, the join and the line ends — which is
    /// why this inserts at the front rather than in front of a computed successor the way
    /// <see cref="SetFill"/> has to.
    /// </remarks>
    internal static void SetOutlineColor(A.Outline outline, XLColor? color)
    {
        foreach (var existing in outline.ChildElements.Where(IsFillElement).ToList())
            existing.Remove();

        if (color is { HasValue: true })
            outline.InsertAt(new A.SolidFill(BuildColor(color)), 0);
    }

    // ── The CT_ShapeProperties sequence ─────────────────────────────────

    /// <summary>The members of the fill choice group, exactly one of which may be present.</summary>
    private static bool IsFillElement(OpenXmlElement element) =>
        element is A.NoFill or A.SolidFill or A.GradientFill or A.BlipFill or A.PatternFill or A.GroupFill;

    /// <summary>The children that must follow the fill.</summary>
    private static bool IsAfterFill(OpenXmlElement element) =>
        element is A.Outline || IsEffectOrLater(element);

    /// <summary>The children that must follow the outline.</summary>
    private static bool IsAfterOutline(OpenXmlElement element) => IsEffectOrLater(element);

    private static bool IsEffectOrLater(OpenXmlElement element) =>
        element is A.EffectList or A.EffectDag or A.Scene3DType or A.Shape3DType or A.ExtensionList;

    // ── Colours ─────────────────────────────────────────────────────────

    /// <summary>
    /// The colour element for a fill: a theme colour becomes <c>a:schemeClr</c> so that it keeps
    /// following the theme, and anything else becomes a literal <c>a:srgbClr</c>.
    /// </summary>
    private static OpenXmlElement BuildColor(XLColor color)
    {
        if (color.ColorType == XLColorType.Theme)
            return new A.SchemeColor { Val = MapSchemeColor(color.ThemeColor) };

        return new A.RgbColorModelHex { Val = color.Color.ToHex().Substring(2) };
    }

    private static A.SchemeColorValues MapSchemeColor(XLThemeColor themeColor) => themeColor switch
    {
        XLThemeColor.Background1 => A.SchemeColorValues.Background1,
        XLThemeColor.Text1 => A.SchemeColorValues.Text1,
        XLThemeColor.Background2 => A.SchemeColorValues.Background2,
        XLThemeColor.Text2 => A.SchemeColorValues.Text2,
        XLThemeColor.Accent1 => A.SchemeColorValues.Accent1,
        XLThemeColor.Accent2 => A.SchemeColorValues.Accent2,
        XLThemeColor.Accent3 => A.SchemeColorValues.Accent3,
        XLThemeColor.Accent4 => A.SchemeColorValues.Accent4,
        XLThemeColor.Accent5 => A.SchemeColorValues.Accent5,
        XLThemeColor.Accent6 => A.SchemeColorValues.Accent6,
        XLThemeColor.Hyperlink => A.SchemeColorValues.Hyperlink,
        XLThemeColor.FollowedHyperlink => A.SchemeColorValues.FollowedHyperlink,
        _ => throw new System.ArgumentOutOfRangeException(
            nameof(themeColor), themeColor, "Unknown theme colour.")
    };
}
