using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace XLibur.Excel.IO;

/// <summary>
/// Applies model changes to a chart that already exists in the package.
/// </summary>
/// <remarks>
/// <para>
/// XLibur never regenerates the XML of a chart it read from a file: <see cref="ChartWriter"/> only
/// emits charts created through <see cref="IXLCharts.Add(XLChartType)"/>, so everything an existing
/// chart part contains — trendlines, error bars, gradient fills, per-point formatting, the chart
/// style and colour parts next to it — survives a load/save round trip untouched.
/// </para>
/// <para>
/// The price of that guarantee is that edits to an existing chart have to be patched into the DOM
/// the reader loaded. This class does exactly that, and only for the properties the caller actually
/// assigned (see <see cref="XLChartSeries.AssignedFormat"/>): a chart nobody edited is not modified
/// at all.
/// </para>
/// </remarks>
internal static class ChartPatcher
{
    /// <summary>
    /// Writes the pending series formatting of a loaded chart back into its chart part.
    /// </summary>
    internal static void PatchChart(WorksheetPart worksheetPart, XLChart xlChart)
    {
        if (!HasPendingChanges(xlChart))
            return;

        // Extended (ChartEx) charts model their series in the cx namespace and do not support the
        // series formatting properties, so there is nothing to patch.
        var chartPart = ResolveChartPart(worksheetPart, xlChart);
        var chart = chartPart?.ChartSpace?.Elements<C.Chart>().FirstOrDefault();
        if (chart == null)
            return;

        ChartFormatting.PatchLegend(chart, xlChart.LegendInternal);

        var plotArea = chart.PlotArea;
        if (plotArea == null)
            return;

        var groups = ChartPlotAreaScanner.Scan(plotArea);
        var primaryKind = ChartPlotAreaScanner.ChoosePrimaryKind(groups);
        if (primaryKind == null)
            return;

        // The reader walked the same groups in the same order, so the n-th series element of the
        // primary kind is the n-th series of the model collection.
        var primaryElements = new List<(OpenXmlCompositeElement Element, XLChartGroupKind Kind)>();
        var secondaryElements = new List<(OpenXmlCompositeElement Element, XLChartGroupKind Kind)>();
        foreach (var group in groups)
        {
            var target = group.Kind == primaryKind.Value ? primaryElements : secondaryElements;
            foreach (var seriesElement in group.SeriesElements)
                target.Add((seriesElement, group.Kind));
        }

        PatchSeries(xlChart.SeriesInternal, primaryElements, xlChart.ChartType);
        PatchSeries(xlChart.SecondarySeriesInternal, secondaryElements,
            xlChart.SecondaryChartType ?? xlChart.ChartType);

        PatchGroupDataLabels(groups, primaryKind.Value, xlChart);
        PatchAxes(plotArea, groups, primaryKind.Value, xlChart);
    }

    /// <summary>
    /// Writes the pending axis properties back. The axes are found by the identifiers the chart groups
    /// point at, the same way the reader found them.
    /// </summary>
    private static void PatchAxes(
        C.PlotArea plotArea, List<XLChartGroup> groups, XLChartGroupKind primaryKind, XLChart xlChart)
    {
        var primaryGroup = ChartPlotAreaScanner.PrimaryGroup(plotArea, groups, primaryKind);
        var primaryValueAxisId = primaryGroup.ValueAxisId;

        ChartFormatting.PatchAxis(
            ChartPlotAreaScanner.FindAxis(plotArea, primaryGroup.CategoryAxisId),
            xlChart.CategoryAxisInternal);
        ChartFormatting.PatchAxis(
            ChartPlotAreaScanner.FindAxis(plotArea, primaryValueAxisId),
            xlChart.ValueAxisInternal);

        var secondaryGroup = groups.Find(g =>
            g.ValueAxisId != null && primaryValueAxisId != null && g.ValueAxisId != primaryValueAxisId);
        if (secondaryGroup != null)
        {
            ChartFormatting.PatchAxis(
                ChartPlotAreaScanner.FindAxis(plotArea, secondaryGroup.ValueAxisId),
                xlChart.SecondaryValueAxisInternal);
        }
    }

    /// <summary>
    /// Writes the chart-wide data labels into every group that accepts them. They apply to the whole
    /// chart, and <see cref="ChartWriter"/> puts them on each group of a new combo chart, so a loaded
    /// combo chart is patched the same way rather than only on its primary group.
    /// </summary>
    private static void PatchGroupDataLabels(
        List<XLChartGroup> groups, XLChartGroupKind primaryKind, XLChart xlChart)
    {
        if (xlChart.DataLabelsInternal.AssignedFormat == XLDataLabelsFormat.None)
            return;

        foreach (var group in groups)
        {
            if (!ChartFormatting.SupportsDataLabels(group.Kind))
                continue;

            // A position Excel does not offer for a group's own type degrades to automatic, so each
            // group is patched against its type — the same mapping PatchSeries uses.
            var chartType = group.Kind == primaryKind
                ? xlChart.ChartType
                : xlChart.SecondaryChartType ?? xlChart.ChartType;

            ChartFormatting.PatchGroupDataLabels(group.Element, xlChart.DataLabelsInternal, chartType);
        }
    }

    private static bool HasPendingChanges(XLChart xlChart) =>
        xlChart.DataLabelsInternal.AssignedFormat != XLDataLabelsFormat.None
        || xlChart.LegendInternal.AssignedFormat != XLChartLegendFormat.None
        || xlChart.CategoryAxisInternal.AssignedFormat != XLChartAxisFormat.None
        || xlChart.ValueAxisInternal.AssignedFormat != XLChartAxisFormat.None
        || xlChart.SecondaryValueAxisInternal.AssignedFormat != XLChartAxisFormat.None
        || HasPendingChanges(xlChart.SeriesInternal)
        || HasPendingChanges(xlChart.SecondarySeriesInternal);

    private static bool HasPendingChanges(XLChartSeriesCollection series)
    {
        foreach (var s in series.Items)
        {
            if (s.AssignedFormat != XLChartSeriesFormat.None
                || s.DataLabelsInternal.AssignedFormat != XLDataLabelsFormat.None)
                return true;
        }

        return false;
    }

    private static void PatchSeries(
        XLChartSeriesCollection series,
        List<(OpenXmlCompositeElement Element, XLChartGroupKind Kind)> seriesElements,
        XLChartType chartType)
    {
        var count = Math.Min(series.Count, seriesElements.Count);
        for (var i = 0; i < count; i++)
        {
            var (element, kind) = seriesElements[i];
            ChartFormatting.PatchSeriesFormat(element, series.Items[i], kind, chartType);
        }
    }

    private static ChartPart? ResolveChartPart(WorksheetPart worksheetPart, XLChart xlChart)
    {
        if (xlChart.RelId == null)
            return null;

        var drawingsPart = worksheetPart.DrawingsPart;
        if (drawingsPart == null)
            return null;

        // GetPartById throws for an unknown id, and the relationship can be missing when the chart
        // was created in a workbook that has not been saved through this package yet.
        if (!drawingsPart.Parts.Any(p => p.RelationshipId == xlChart.RelId))
            return null;

        return drawingsPart.GetPartById(xlChart.RelId) as ChartPart;
    }
}
