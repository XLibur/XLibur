using System;
using System.Linq;
using DocumentFormat.OpenXml;
using XLibur.Excel.Drawings;
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
    internal static void SetFill(OpenXmlCompositeElement shapeProperties, XLColor? color) =>
        ReplaceFill(shapeProperties, color is { HasValue: true } ? new A.SolidFill(BuildColor(color)) : null);

    /// <summary>
    /// Sets the fill to an explicit <c>a:noFill</c>, which says the shape has no fill rather than
    /// leaving it to the shape's style.
    /// </summary>
    internal static void SetNoFill(OpenXmlCompositeElement shapeProperties) =>
        ReplaceFill(shapeProperties, new A.NoFill());

    /// <summary>Removes an explicit <c>a:noFill</c>, leaving any other fill alone.</summary>
    internal static void RemoveNoFill(OpenXmlCompositeElement shapeProperties)
    {
        foreach (var noFill in shapeProperties.Elements<A.NoFill>().ToList())
            noFill.Remove();
    }

    /// <summary>
    /// Puts <paramref name="fill"/> in the place of whatever fill there is, or only removes that fill
    /// when <paramref name="fill"/> is <c>null</c>.
    /// </summary>
    private static void ReplaceFill(OpenXmlCompositeElement shapeProperties, OpenXmlElement? fill)
    {
        foreach (var existing in shapeProperties.ChildElements.Where(IsFillElement).ToList())
            existing.Remove();

        if (fill is null)
            return;

        // The fill precedes the outline and every effect in CT_ShapeProperties.
        var anchor = shapeProperties.ChildElements.FirstOrDefault(IsAfterFill);
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

    /// <summary>
    /// Sets the outline's fill to an explicit <c>a:noFill</c>: the shape has no outline, whatever its
    /// style says. The outline's width, dash, join and ends are untouched.
    /// </summary>
    internal static void SetOutlineNoFill(A.Outline outline)
    {
        foreach (var existing in outline.ChildElements.Where(IsFillElement).ToList())
            existing.Remove();

        outline.InsertAt(new A.NoFill(), 0);
    }

    /// <summary>Removes an explicit <c>a:noFill</c> from the outline, leaving any other fill alone.</summary>
    internal static void RemoveOutlineNoFill(A.Outline outline)
    {
        foreach (var noFill in outline.Elements<A.NoFill>().ToList())
            noFill.Remove();
    }

    /// <summary>
    /// Sets the outline's dash to the preset <paramref name="dash"/>, or removes the dash when it is
    /// <c>null</c>. A custom dash goes too: the dashes are one choice group.
    /// </summary>
    /// <remarks>
    /// <c>CT_LineProperties</c> is a sequence of the fill, the dash, the join, the two line ends and
    /// <c>a:extLst</c>, so the dash goes straight after the fill.
    /// </remarks>
    internal static void SetOutlineDash(A.Outline outline, XLLineDashStyle? dash)
    {
        foreach (var existing in outline.ChildElements.Where(e => e is A.PresetDash or A.CustomDash).ToList())
            existing.Remove();

        if (dash is not { } value)
            return;

        var presetDash = new A.PresetDash { Val = ToPresetDash(value) };
        var fill = outline.ChildElements.LastOrDefault(IsFillElement);
        if (fill != null)
            outline.InsertAfter(presetDash, fill);
        else
            outline.InsertAt(presetDash, 0);
    }

    /// <summary>The <c>a:prstDash</c> value of a dash style.</summary>
    internal static A.PresetLineDashValues ToPresetDash(XLLineDashStyle dash) => dash switch
    {
        XLLineDashStyle.Solid => A.PresetLineDashValues.Solid,
        XLLineDashStyle.Dot => A.PresetLineDashValues.Dot,
        XLLineDashStyle.Dash => A.PresetLineDashValues.Dash,
        XLLineDashStyle.LongDash => A.PresetLineDashValues.LargeDash,
        XLLineDashStyle.DashDot => A.PresetLineDashValues.DashDot,
        XLLineDashStyle.LongDashDot => A.PresetLineDashValues.LargeDashDot,
        XLLineDashStyle.LongDashDotDot => A.PresetLineDashValues.LargeDashDotDot,
        XLLineDashStyle.SystemDash => A.PresetLineDashValues.SystemDash,
        XLLineDashStyle.SystemDot => A.PresetLineDashValues.SystemDot,
        XLLineDashStyle.SystemDashDot => A.PresetLineDashValues.SystemDashDot,
        XLLineDashStyle.SystemDashDotDot => A.PresetLineDashValues.SystemDashDotDot,
        _ => throw new ArgumentOutOfRangeException(nameof(dash), dash, "Unknown dash style."),
    };

    /// <summary>The dash style of an <c>a:prstDash</c> value, or <c>null</c> for one XLibur does not know.</summary>
    internal static XLLineDashStyle? FromPresetDash(A.PresetLineDashValues dash)
    {
        foreach (var style in Enum.GetValues<XLLineDashStyle>())
        {
            if (ToPresetDash(style) == dash)
                return style;
        }

        return null;
    }

    // ── a:xfrm ──────────────────────────────────────────────────────────

    /// <summary>
    /// Sets the rotation, in degrees clockwise, on <c>a:xfrm</c>, or removes it when
    /// <paramref name="degrees"/> is <c>null</c>. The offset and extents are untouched.
    /// </summary>
    internal static void SetRotation(A.Transform2D transform, double? degrees) =>
        transform.Rotation = degrees is { } angle ? DrawingUnits.DegreesToAngle(angle) : null;

    /// <summary>
    /// Sets one flip on <c>a:xfrm</c> when <paramref name="flip"/> is <c>true</c>, and removes it
    /// otherwise: not flipped is the attribute's default, and Excel writes it only when set.
    /// </summary>
    internal static void SetFlip(A.Transform2D transform, bool horizontal, bool? flip)
    {
        BooleanValue? value = flip == true ? true : null;
        if (horizontal)
            transform.HorizontalFlip = value;
        else
            transform.VerticalFlip = value;
    }

    // ── a:effectLst ─────────────────────────────────────────────────────

    /// <summary>
    /// Sets the outer shadow, or removes it when <paramref name="shadow"/> is <c>null</c>. An outer
    /// shadow the list already holds is edited in place, keeping what XLibur does not model, such as
    /// its alignment and scale; the other effects are untouched. A list left empty is removed.
    /// </summary>
    /// <remarks>
    /// A shape with an <c>a:effectDag</c> instead of a list is left alone: the two are one choice
    /// group, and a DAG cannot be edited as a list.
    /// </remarks>
    internal static void SetOuterShadow(OpenXmlCompositeElement shapeProperties, XLPictureShadow? shadow)
    {
        if (shapeProperties.Elements<A.EffectDag>().Any())
            return;

        var effects = shapeProperties.Elements<A.EffectList>().FirstOrDefault();
        if (shadow is not { } value)
        {
            if (effects is null)
                return;

            foreach (var existing in effects.Elements<A.OuterShadow>().ToList())
                existing.Remove();

            if (!effects.HasChildren)
                effects.Remove();

            return;
        }

        if (effects is null)
        {
            effects = new A.EffectList();
            var anchor = shapeProperties.ChildElements.FirstOrDefault(IsAfterEffects);
            if (anchor != null)
                shapeProperties.InsertBefore(effects, anchor);
            else
                shapeProperties.Append(effects);
        }

        var outerShadow = effects.Elements<A.OuterShadow>().FirstOrDefault();
        if (outerShadow is null)
        {
            // Excel writes rotWithShape="0" on the shadows its gallery makes.
            outerShadow = new A.OuterShadow { RotateWithShape = false };

            // CT_EffectList is a sequence: blur, fillOverlay, glow, innerShdw, outerShdw, prstShdw,
            // reflection, softEdge.
            var after = effects.ChildElements.FirstOrDefault(e => e is A.PresetShadow or A.Reflection or A.SoftEdge);
            if (after != null)
                effects.InsertBefore(outerShadow, after);
            else
                effects.Append(outerShadow);
        }

        outerShadow.BlurRadius = DrawingUnits.PointsToEmu(value.BlurPt);
        outerShadow.Distance = DrawingUnits.PointsToEmu(value.DistancePt);
        outerShadow.Direction = DrawingUnits.DegreesToAngle(value.AngleDeg);

        foreach (var color in outerShadow.ChildElements.ToList())
            color.Remove();

        var shadowColor = BuildColor(value.Color);
        if (value.TransparencyPct > 0)
            shadowColor.AppendChild(new A.Alpha { Val = DrawingUnits.PercentToThousandths(100 - value.TransparencyPct) });

        outerShadow.Append(shadowColor);
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
        element is A.EffectList or A.EffectDag || IsAfterEffects(element);

    /// <summary>The children that must follow the effects.</summary>
    private static bool IsAfterEffects(OpenXmlElement element) =>
        element is A.Scene3DType or A.Shape3DType or A.ExtensionList;

    // ── Colours ─────────────────────────────────────────────────────────

    /// <summary>
    /// The colour element for a fill: a theme colour becomes <c>a:schemeClr</c> so that it keeps
    /// following the theme, and anything else becomes a literal <c>a:srgbClr</c>.
    /// </summary>
    internal static OpenXmlElement BuildColor(XLColor color)
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
