using System;
using System.Linq;
using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace XLibur.Excel.IO.Charts;

/// <summary>
/// The <c>c:dLbls</c> element, at both the levels a chart carries one: on a series (<c>c:ser</c>) and
/// on a chart group (<c>c:barChart</c> and friends).
/// </summary>
/// <remarks>
/// <para>
/// <see cref="Apply"/> covers a chart being created and a chart loaded from a file, and both levels.
/// The parent used to decide which of four functions ran — two builders' worth of call sites and two
/// patch entry points — when all it actually decides is where the element is inserted.
/// </para>
/// <para>
/// Individual point overrides (<c>c:dLbl</c>), label text and shape properties, and the separator are
/// never touched.
/// </para>
/// </remarks>
internal static class ChartDataLabelsXml
{
    /// <summary>
    /// Whether a chart group and its series accept a <c>c:dLbls</c> child. Only the surface types do
    /// not: neither <c>CT_SurfaceChart</c> nor <c>CT_SurfaceSer</c> has one.
    /// </summary>
    internal static bool Supports(XLChartGroupKind kind) =>
        kind is not (XLChartGroupKind.Surface or XLChartGroupKind.Surface3D);

    /// <summary>
    /// Seeds the model from a <c>c:dLbls</c> element, without marking the values as assigned by the
    /// caller.
    /// </summary>
    internal static void Read(C.DataLabels? dataLabels, XLChartDataLabels labels)
    {
        if (dataLabels == null)
            return;

        labels.SeedLoaded(
            showValue: dataLabels.Elements<C.ShowValue>().FirstOrDefault()?.Val?.Value ?? false,
            showCategoryName: dataLabels.Elements<C.ShowCategoryName>().FirstOrDefault()?.Val?.Value ?? false,
            showSeriesName: dataLabels.Elements<C.ShowSeriesName>().FirstOrDefault()?.Val?.Value ?? false,
            showPercentage: dataLabels.Elements<C.ShowPercent>().FirstOrDefault()?.Val?.Value ?? false,
            numberFormat: dataLabels.Elements<C.NumberingFormat>().FirstOrDefault()?.FormatCode?.Value,
            position: ReadPosition(dataLabels));
    }

    /// <summary>
    /// Writes the assigned data-label properties into <paramref name="parent"/>, which is either a
    /// series element (<c>c:ser</c>) or a chart-group element. Labels nobody assigned are not written.
    /// </summary>
    /// <param name="parent">The element the <c>c:dLbls</c> belongs to.</param>
    /// <param name="labels">The labels to write.</param>
    /// <param name="chartType">
    /// The chart type the labels belong to, which decides whether an explicit position can be written
    /// at all.
    /// </param>
    internal static void Apply(
        OpenXmlCompositeElement parent, XLChartDataLabels labels, XLChartType chartType)
    {
        if (labels.AssignedFormat == XLDataLabelsFormat.None)
            return;

        var element = parent.Elements<C.DataLabels>().FirstOrDefault();
        if (element == null)
        {
            Insert(parent, Body(labels, chartType));
            return;
        }

        Patch(element, labels, chartType);
    }

    /// <summary>
    /// Puts a new <c>c:dLbls</c> where the schema wants it. On a series that is by child order; on a
    /// chart group it is directly after the last <c>c:ser</c>, which is where every group type that
    /// has one keeps it.
    /// </summary>
    private static void Insert(OpenXmlCompositeElement parent, C.DataLabels element)
    {
        if (parent.LocalName == "ser")
        {
            ChartElementOrder.InsertOrdered(parent, element, ChartElementOrder.SeriesChildOrder);
            return;
        }

        var lastSeries = parent.ChildElements.LastOrDefault(e => e.LocalName == "ser");
        if (lastSeries != null)
            parent.InsertAfter(element, lastSeries);
        else
            parent.Append(element);
    }

    /// <summary>
    /// The <c>c:dLbls</c> a model with nothing to edit into produces. Excel writes all of the show
    /// flags whenever it writes <c>c:dLbls</c> at all, and reads a missing one as "inherit", which
    /// makes the result depend on the chart style; writing them out keeps what the caller asked for
    /// unambiguous.
    /// </summary>
    private static C.DataLabels Body(XLChartDataLabels labels, XLChartType chartType)
    {
        var dataLabels = new C.DataLabels();

        if (labels.NumberFormat != null)
        {
            dataLabels.Append(new C.NumberingFormat
            {
                FormatCode = labels.NumberFormat,
                SourceLinked = false
            });
        }

        var position = MapPosition(labels.EffectivePosition(chartType));
        if (position != null)
            dataLabels.Append(new C.DataLabelPosition { Val = position });

        dataLabels.Append(new C.ShowLegendKey { Val = false });
        dataLabels.Append(new C.ShowValue { Val = labels.ShowValue });
        dataLabels.Append(new C.ShowCategoryName { Val = labels.ShowCategoryName });
        dataLabels.Append(new C.ShowSeriesName { Val = labels.ShowSeriesName });
        dataLabels.Append(new C.ShowPercent { Val = labels.ShowPercentage });
        dataLabels.Append(new C.ShowBubbleSize { Val = false });

        return dataLabels;
    }

#pragma warning disable S3776 // One independent, flat block per assigned data-label property
    private static void Patch(
        C.DataLabels dataLabels, XLChartDataLabels labels, XLChartType chartType)
    {
        var assigned = labels.AssignedFormat;

        // A c:delete of the whole label set would override every flag below it.
        foreach (var deleted in dataLabels.Elements<C.Delete>().ToList())
            deleted.Remove();

        if ((assigned & XLDataLabelsFormat.NumberFormat) != 0)
        {
            foreach (var existing in dataLabels.Elements<C.NumberingFormat>().ToList())
                existing.Remove();

            if (labels.NumberFormat != null)
            {
                ChartElementOrder.InsertOrdered(dataLabels,
                    new C.NumberingFormat { FormatCode = labels.NumberFormat, SourceLinked = false },
                    ChartElementOrder.DataLabelsChildOrder);
            }
        }

        if ((assigned & XLDataLabelsFormat.Position) != 0)
        {
            foreach (var existing in dataLabels.Elements<C.DataLabelPosition>().ToList())
                existing.Remove();

            var position = MapPosition(labels.EffectivePosition(chartType));
            if (position != null)
                ChartElementOrder.InsertOrdered(dataLabels, new C.DataLabelPosition { Val = position },
                    ChartElementOrder.DataLabelsChildOrder);
        }

        if ((assigned & XLDataLabelsFormat.ShowValue) != 0)
            SetFlag<C.ShowValue>(dataLabels, labels.ShowValue);
        if ((assigned & XLDataLabelsFormat.ShowCategoryName) != 0)
            SetFlag<C.ShowCategoryName>(dataLabels, labels.ShowCategoryName);
        if ((assigned & XLDataLabelsFormat.ShowSeriesName) != 0)
            SetFlag<C.ShowSeriesName>(dataLabels, labels.ShowSeriesName);
        if ((assigned & XLDataLabelsFormat.ShowPercentage) != 0)
            SetFlag<C.ShowPercent>(dataLabels, labels.ShowPercentage);
    }
#pragma warning restore S3776

    private static void SetFlag<TFlag>(C.DataLabels dataLabels, bool value)
        where TFlag : OpenXmlLeafElement, new()
    {
        var existing = dataLabels.Elements<TFlag>().FirstOrDefault();
        if (existing != null)
        {
            existing.SetAttribute(new OpenXmlAttribute("val", string.Empty, value ? "1" : "0"));
            return;
        }

        var flag = new TFlag();
        flag.SetAttribute(new OpenXmlAttribute("val", string.Empty, value ? "1" : "0"));
        ChartElementOrder.InsertOrdered(dataLabels, flag, ChartElementOrder.DataLabelsChildOrder);
    }

    private static XLDataLabelPosition ReadPosition(C.DataLabels dataLabels)
    {
        var position = dataLabels.Elements<C.DataLabelPosition>().FirstOrDefault()?.Val;
        if (position == null)
            return XLDataLabelPosition.Auto;

        var value = position.Value;
        if (value == C.DataLabelPositionValues.Center) return XLDataLabelPosition.Center;
        if (value == C.DataLabelPositionValues.InsideEnd) return XLDataLabelPosition.InsideEnd;
        if (value == C.DataLabelPositionValues.InsideBase) return XLDataLabelPosition.InsideBase;
        if (value == C.DataLabelPositionValues.OutsideEnd) return XLDataLabelPosition.OutsideEnd;
        if (value == C.DataLabelPositionValues.BestFit) return XLDataLabelPosition.BestFit;
        if (value == C.DataLabelPositionValues.Left) return XLDataLabelPosition.Left;
        if (value == C.DataLabelPositionValues.Right) return XLDataLabelPosition.Right;
        if (value == C.DataLabelPositionValues.Top) return XLDataLabelPosition.Above;
        if (value == C.DataLabelPositionValues.Bottom) return XLDataLabelPosition.Below;
        return XLDataLabelPosition.Auto;
    }

    private static C.DataLabelPositionValues? MapPosition(XLDataLabelPosition position) => position switch
    {
        XLDataLabelPosition.Auto => null,
        XLDataLabelPosition.Center => C.DataLabelPositionValues.Center,
        XLDataLabelPosition.InsideEnd => C.DataLabelPositionValues.InsideEnd,
        XLDataLabelPosition.InsideBase => C.DataLabelPositionValues.InsideBase,
        XLDataLabelPosition.OutsideEnd => C.DataLabelPositionValues.OutsideEnd,
        XLDataLabelPosition.BestFit => C.DataLabelPositionValues.BestFit,
        XLDataLabelPosition.Left => C.DataLabelPositionValues.Left,
        XLDataLabelPosition.Right => C.DataLabelPositionValues.Right,
        XLDataLabelPosition.Above => C.DataLabelPositionValues.Top,
        XLDataLabelPosition.Below => C.DataLabelPositionValues.Bottom,
        _ => throw new ArgumentOutOfRangeException(nameof(position), position,
            "Unknown data label position.")
    };
}
