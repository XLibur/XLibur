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
#pragma warning disable S3776 // One independent, flat block per assigned axis property
    internal static void Apply(OpenXmlCompositeElement? axis, XLChartAxis model)
    {
        var assigned = model.AssignedFormat;
        if (axis == null || assigned == XLChartAxisFormat.None)
            return;

        if ((assigned & XLChartAxisFormat.Visible) != 0)
        {
            foreach (var existing in axis.Elements<C.Delete>().ToList())
                existing.Remove();
            ChartElementOrder.InsertOrdered(axis, new C.Delete { Val = !model.Visible },
                ChartElementOrder.AxisChildOrder);
        }

        if ((assigned & XLChartAxisFormat.MajorGridlines) != 0)
        {
            foreach (var existing in axis.Elements<C.MajorGridlines>().ToList())
                existing.Remove();
            if (model.MajorGridlines)
                ChartElementOrder.InsertOrdered(axis, new C.MajorGridlines(),
                    ChartElementOrder.AxisChildOrder);
        }

        if ((assigned & XLChartAxisFormat.Title) != 0)
        {
            foreach (var existing in axis.Elements<C.Title>().ToList())
                existing.Remove();
            if (model.Title != null)
                ChartElementOrder.InsertOrdered(axis, TitleElement(model.Title),
                    ChartElementOrder.AxisChildOrder);
        }

        if ((assigned & XLChartAxisFormat.NumberFormat) != 0)
        {
            foreach (var existing in axis.Elements<C.NumberingFormat>().ToList())
                existing.Remove();
            if (model.NumberFormat != null)
            {
                ChartElementOrder.InsertOrdered(axis,
                    new C.NumberingFormat { FormatCode = model.NumberFormat, SourceLinked = false },
                    ChartElementOrder.AxisChildOrder);
            }
        }

        if ((assigned & (XLChartAxisFormat.Min | XLChartAxisFormat.Max
                         | XLChartAxisFormat.Orientation | XLChartAxisFormat.LogScale
                         | XLChartAxisFormat.LogBase)) != 0)
        {
            ApplyScaling(axis, model, assigned);
        }

        // CT_CatAx has no unit elements, so a bubble chart's horizontal axis — which the model calls
        // a value axis, because it plots numbers — still cannot carry them.
        if (model.IsValueAxis && axis is C.ValueAxis)
            ApplyUnits(axis, model, assigned);
    }
#pragma warning restore S3776

#pragma warning disable S3776 // One independent, flat block per assigned scaling property
    private static void ApplyScaling(
        OpenXmlCompositeElement axis, XLChartAxis model, XLChartAxisFormat assigned)
    {
        var scaling = axis.Elements<C.Scaling>().FirstOrDefault();
        if (scaling == null)
        {
            scaling = new C.Scaling();
            ChartElementOrder.InsertOrdered(axis, scaling, ChartElementOrder.AxisChildOrder);
        }

        if ((assigned & (XLChartAxisFormat.LogScale | XLChartAxisFormat.LogBase)) != 0)
        {
            foreach (var existing in scaling.Elements<C.LogBase>().ToList())
                existing.Remove();
            // c:logBase belongs to a value axis; Excel rejects it on a category axis.
            if (model.LogScale && model.IsValueAxis)
                ChartElementOrder.InsertOrdered(scaling, new C.LogBase { Val = model.LogBase },
                    ChartElementOrder.ScalingChildOrder);
        }

        if ((assigned & XLChartAxisFormat.Orientation) != 0)
        {
            foreach (var existing in scaling.Elements<C.Orientation>().ToList())
                existing.Remove();
            ChartElementOrder.InsertOrdered(scaling, new C.Orientation
            {
                Val = model.Orientation == XLAxisOrientation.MaxMin
                    ? C.OrientationValues.MaxMin
                    : C.OrientationValues.MinMax
            }, ChartElementOrder.ScalingChildOrder);
        }

        if ((assigned & XLChartAxisFormat.Max) != 0)
        {
            foreach (var existing in scaling.Elements<C.MaxAxisValue>().ToList())
                existing.Remove();
            if (model.Max != null)
                ChartElementOrder.InsertOrdered(scaling, new C.MaxAxisValue { Val = model.Max.Value },
                    ChartElementOrder.ScalingChildOrder);
        }

        if ((assigned & XLChartAxisFormat.Min) != 0)
        {
            foreach (var existing in scaling.Elements<C.MinAxisValue>().ToList())
                existing.Remove();
            if (model.Min != null)
                ChartElementOrder.InsertOrdered(scaling, new C.MinAxisValue { Val = model.Min.Value },
                    ChartElementOrder.ScalingChildOrder);
        }
    }
#pragma warning restore S3776

    private static void ApplyUnits(
        OpenXmlCompositeElement axis, XLChartAxis model, XLChartAxisFormat assigned)
    {
        if ((assigned & XLChartAxisFormat.MajorUnit) != 0)
        {
            foreach (var existing in axis.Elements<C.MajorUnit>().ToList())
                existing.Remove();
            if (model.MajorUnit != null)
                ChartElementOrder.InsertOrdered(axis, new C.MajorUnit { Val = model.MajorUnit.Value },
                    ChartElementOrder.AxisChildOrder);
        }

        if ((assigned & XLChartAxisFormat.MinorUnit) != 0)
        {
            foreach (var existing in axis.Elements<C.MinorUnit>().ToList())
                existing.Remove();
            if (model.MinorUnit != null)
                ChartElementOrder.InsertOrdered(axis, new C.MinorUnit { Val = model.MinorUnit.Value },
                    ChartElementOrder.AxisChildOrder);
        }
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
