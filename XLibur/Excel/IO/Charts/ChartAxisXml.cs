using System;
using System.Linq;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace XLibur.Excel.IO.Charts;

/// <summary>
/// A chart axis — <c>c:catAx</c> or <c>c:valAx</c> — and its <c>c:scaling</c>: how one is read into
/// <see cref="XLChartAxis"/>, and how an assigned model is written back.
/// </summary>
/// <remarks>
/// <para>
/// <see cref="Apply"/> covers both a chart being created and a chart loaded from a file. The writer
/// hands it an axis carrying only the children <c>CT_CatAx</c> and <c>CT_ValAx</c> require —
/// <c>c:axId</c>, <c>c:scaling</c>, <c>c:delete</c>, <c>c:axPos</c>, <c>c:crossAx</c> — and the
/// patcher hands it a fully populated one. Every property is written by schema order rather than by
/// append order, so the two arrive at the same XML without either caller knowing which case it is.
/// </para>
/// <para>
/// Everything the model does not carry — tick marks, label positions, line and text formatting — is
/// left exactly as it was.
/// </para>
/// </remarks>
internal static class ChartAxisXml
{
    /// <summary>
    /// Seeds the model from a <c>c:catAx</c> or <c>c:valAx</c> element.
    /// </summary>
    internal static void Read(OpenXmlCompositeElement? axis, XLChartAxis model)
    {
        if (axis == null)
            return;

        var scaling = axis.Elements<C.Scaling>().FirstOrDefault();
        var logBase = scaling?.Elements<C.LogBase>().FirstOrDefault()?.Val?.Value;
        var orientation = scaling?.Elements<C.Orientation>().FirstOrDefault()?.Val;

        model.SeedLoaded(
            title: ReadTitle(axis),
            numberFormat: axis.Elements<C.NumberingFormat>().FirstOrDefault()?.FormatCode?.Value,
            min: scaling?.Elements<C.MinAxisValue>().FirstOrDefault()?.Val?.Value,
            max: scaling?.Elements<C.MaxAxisValue>().FirstOrDefault()?.Val?.Value,
            majorUnit: axis.Elements<C.MajorUnit>().FirstOrDefault()?.Val?.Value,
            minorUnit: axis.Elements<C.MinorUnit>().FirstOrDefault()?.Val?.Value,
            // c:delete says the axis is hidden, and defaults to false when absent.
            visible: !(axis.Elements<C.Delete>().FirstOrDefault()?.Val?.Value ?? false),
            majorGridlines: axis.Elements<C.MajorGridlines>().Any(),
            orientation: orientation != null && orientation.Value == C.OrientationValues.MaxMin
                ? XLAxisOrientation.MaxMin
                : XLAxisOrientation.MinMax,
            logScale: logBase != null,
            logBase: logBase ?? 10);
    }

    /// <summary>
    /// Writes the assigned axis properties into <paramref name="axis"/>, adding, editing or removing
    /// each child as the model requires. An axis nobody edited is not modified at all.
    /// </summary>
    internal static void Apply(OpenXmlCompositeElement? axis, XLChartAxis model)
    {
        var assigned = model.AssignedFormat;
        if (axis == null || assigned == XLChartAxisFormat.None)
            return;

        if (IsAssigned(assigned, XLChartAxisFormat.Visible))
            ReplaceAxisChild<C.Delete>(axis, new C.Delete { Val = !model.Visible });

        if (IsAssigned(assigned, XLChartAxisFormat.MajorGridlines))
            ReplaceAxisChild<C.MajorGridlines>(axis, model.MajorGridlines ? new C.MajorGridlines() : null);

        if (IsAssigned(assigned, XLChartAxisFormat.Title))
            ReplaceAxisChild<C.Title>(axis, model.Title != null ? TitleElement(model.Title) : null);

        if (IsAssigned(assigned, XLChartAxisFormat.NumberFormat))
        {
            ReplaceAxisChild<C.NumberingFormat>(axis, model.NumberFormat != null
                ? new C.NumberingFormat { FormatCode = model.NumberFormat, SourceLinked = false }
                : null);
        }

        if (IsAssigned(assigned, XLChartAxisFormat.Min | XLChartAxisFormat.Max
                                 | XLChartAxisFormat.Orientation | XLChartAxisFormat.LogScale
                                 | XLChartAxisFormat.LogBase))
        {
            ApplyScaling(axis, model, assigned);
        }

        // CT_CatAx has no unit elements, so a bubble chart's horizontal axis — which the model calls
        // a value axis, because it plots numbers — still cannot carry them.
        if (model.IsValueAxis && axis is C.ValueAxis)
            ApplyUnits(axis, model, assigned);
    }

    private static void ApplyScaling(
        OpenXmlCompositeElement axis, XLChartAxis model, XLChartAxisFormat assigned)
    {
        var scaling = axis.Elements<C.Scaling>().FirstOrDefault();
        if (scaling == null)
        {
            scaling = new C.Scaling();
            ChartElementOrder.InsertOrdered(axis, scaling, ChartElementOrder.AxisChildOrder);
        }

        if (IsAssigned(assigned, XLChartAxisFormat.LogScale | XLChartAxisFormat.LogBase))
        {
            // c:logBase belongs to a value axis; Excel rejects it on a category axis.
            ReplaceScalingChild<C.LogBase>(scaling, model.LogScale && model.IsValueAxis
                ? new C.LogBase { Val = model.LogBase }
                : null);
        }

        if (IsAssigned(assigned, XLChartAxisFormat.Orientation))
        {
            ReplaceScalingChild<C.Orientation>(scaling, new C.Orientation
            {
                Val = model.Orientation == XLAxisOrientation.MaxMin
                    ? C.OrientationValues.MaxMin
                    : C.OrientationValues.MinMax
            });
        }

        if (IsAssigned(assigned, XLChartAxisFormat.Max))
        {
            ReplaceScalingChild<C.MaxAxisValue>(scaling,
                model.Max is { } max ? new C.MaxAxisValue { Val = max } : null);
        }

        if (IsAssigned(assigned, XLChartAxisFormat.Min))
        {
            ReplaceScalingChild<C.MinAxisValue>(scaling,
                model.Min is { } min ? new C.MinAxisValue { Val = min } : null);
        }
    }

    private static void ApplyUnits(
        OpenXmlCompositeElement axis, XLChartAxis model, XLChartAxisFormat assigned)
    {
        if (IsAssigned(assigned, XLChartAxisFormat.MajorUnit))
        {
            ReplaceAxisChild<C.MajorUnit>(axis,
                model.MajorUnit is { } majorUnit ? new C.MajorUnit { Val = majorUnit } : null);
        }

        if (IsAssigned(assigned, XLChartAxisFormat.MinorUnit))
        {
            ReplaceAxisChild<C.MinorUnit>(axis,
                model.MinorUnit is { } minorUnit ? new C.MinorUnit { Val = minorUnit } : null);
        }
    }

    private static bool IsAssigned(XLChartAxisFormat assigned, XLChartAxisFormat properties) =>
        (assigned & properties) != 0;

    private static void ReplaceAxisChild<T>(OpenXmlCompositeElement axis, OpenXmlElement? replacement)
        where T : OpenXmlElement =>
        ReplaceChild<T>(axis, replacement, ChartElementOrder.AxisChildOrder);

    private static void ReplaceScalingChild<T>(C.Scaling scaling, OpenXmlElement? replacement)
        where T : OpenXmlElement =>
        ReplaceChild<T>(scaling, replacement, ChartElementOrder.ScalingChildOrder);

    /// <summary>
    /// Removes every <typeparamref name="T"/> child of <paramref name="parent"/>, then inserts
    /// <paramref name="replacement"/> at its schema position, or leaves the child absent when it is
    /// <c>null</c>.
    /// </summary>
    private static void ReplaceChild<T>(OpenXmlCompositeElement parent, OpenXmlElement? replacement, Type[] order)
        where T : OpenXmlElement
    {
        foreach (var existing in parent.Elements<T>().ToList())
            existing.Remove();
        if (replacement != null)
            ChartElementOrder.InsertOrdered(parent, replacement, order);
    }

    /// <summary>
    /// The <c>c:title</c> of an axis: the same rich text block a chart title carries, under a
    /// different parent.
    /// </summary>
    private static C.Title TitleElement(string title) =>
        new(ChartTitleXml.LiteralText(title), new C.Overlay { Val = false });

    private static string? ReadTitle(OpenXmlCompositeElement axis)
    {
        var title = axis.Elements<C.Title>().FirstOrDefault();
        if (title == null)
            return null;

        var text = string.Concat(title.Descendants<A.Text>().Select(t => t.Text));
        return string.IsNullOrEmpty(text) ? null : text;
    }
}
