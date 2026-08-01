using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace XLibur.Excel.IO;

/// <summary>
/// The kind of chart group element (<c>c:barChart</c>, <c>c:lineChart</c>, …) found in a plot area.
/// Several <see cref="XLChartType"/> values map onto one kind.
/// </summary>
internal enum XLChartGroupKind
{
    Bar,
    Bar3D,
    Pie,
    Pie3D,

    /// <summary>A pie-of-pie or bar-of-pie chart (<c>c:ofPieChart</c>).</summary>
    OfPie,
    Doughnut,
    Area,
    Area3D,
    Line,
    Line3D,
    Radar,
    Bubble,
    Scatter,
    Stock,
    Surface,
    Surface3D
}

/// <summary>
/// One chart group of a plot area: the group element itself, its <c>c:ser</c> children in document
/// order, and the value axis it is plotted against.
/// </summary>
internal sealed class XLChartGroup
{
    internal XLChartGroup(
        XLChartGroupKind kind,
        OpenXmlCompositeElement element,
        List<OpenXmlCompositeElement> seriesElements,
        uint? categoryAxisId,
        uint? valueAxisId)
    {
        Kind = kind;
        Element = element;
        SeriesElements = seriesElements;
        CategoryAxisId = categoryAxisId;
        ValueAxisId = valueAxisId;
    }

    internal XLChartGroupKind Kind { get; }

    internal OpenXmlCompositeElement Element { get; }

    internal List<OpenXmlCompositeElement> SeriesElements { get; }

    /// <summary>
    /// The identifier of the horizontal axis this group is plotted against — the first <c>c:axId</c>
    /// of the group — or <c>null</c> for the axis-less pie and doughnut groups.
    /// </summary>
    internal uint? CategoryAxisId { get; }

    /// <summary>
    /// The identifier of the value axis this group is plotted against — the second <c>c:axId</c> of
    /// the group — or <c>null</c> for the axis-less pie and doughnut groups.
    /// </summary>
    internal uint? ValueAxisId { get; }

    /// <summary>
    /// Whether the group carries X/Y values (<c>c:xVal</c>/<c>c:yVal</c>) instead of categories and
    /// values (<c>c:cat</c>/<c>c:val</c>).
    /// </summary>
    internal bool IsXyBased => Kind is XLChartGroupKind.Scatter or XLChartGroupKind.Bubble;

    /// <summary>Whether this is one of the 3D group types, which XLibur reads but never writes.</summary>
    internal bool Is3D => Kind is XLChartGroupKind.Bar3D or XLChartGroupKind.Pie3D
        or XLChartGroupKind.Area3D or XLChartGroupKind.Line3D or XLChartGroupKind.Surface3D;
}

/// <summary>
/// Walks a plot area and reports its chart groups. Shared by <see cref="ChartReader"/> — which turns
/// the groups into the model — and <see cref="ChartPatcher"/> — which walks them again on save to
/// find the <c>c:ser</c> element belonging to a given model series. Because both use the same scan,
/// series positions line up on load and on save.
/// </summary>
internal static class ChartPlotAreaScanner
{
    /// <summary>
    /// The order in which a plot area's groups are considered for the role of primary chart type.
    /// The first kind present wins, which keeps a bar-plus-line combo reading as a bar chart with a
    /// line secondary regardless of the order the two groups appear in the file.
    /// </summary>
    private static readonly XLChartGroupKind[] PrimaryPrecedence =
    [
        XLChartGroupKind.Bar,
        XLChartGroupKind.Bar3D,
        XLChartGroupKind.Pie,
        XLChartGroupKind.Pie3D,
        XLChartGroupKind.OfPie,
        XLChartGroupKind.Doughnut,
        XLChartGroupKind.Area,
        XLChartGroupKind.Area3D,
        XLChartGroupKind.Line,
        XLChartGroupKind.Line3D,
        XLChartGroupKind.Radar,
        XLChartGroupKind.Bubble,
        XLChartGroupKind.Scatter,
        XLChartGroupKind.Stock,
        XLChartGroupKind.Surface,
        XLChartGroupKind.Surface3D
    ];

    /// <summary>
    /// Collects the chart groups of a plot area in document order.
    /// </summary>
    internal static List<XLChartGroup> Scan(C.PlotArea plotArea)
    {
        var groups = new List<XLChartGroup>();

        foreach (var child in plotArea.ChildElements)
        {
            switch (child)
            {
                case C.BarChart bar:
                    groups.Add(Build(XLChartGroupKind.Bar, bar, bar.Elements<C.BarChartSeries>()));
                    break;
                case C.Bar3DChart bar3D:
                    groups.Add(Build(XLChartGroupKind.Bar3D, bar3D, bar3D.Elements<C.BarChartSeries>()));
                    break;
                case C.PieChart pie:
                    groups.Add(Build(XLChartGroupKind.Pie, pie, pie.Elements<C.PieChartSeries>()));
                    break;
                case C.Pie3DChart pie3D:
                    groups.Add(Build(XLChartGroupKind.Pie3D, pie3D, pie3D.Elements<C.PieChartSeries>()));
                    break;
                case C.OfPieChart ofPie:
                    groups.Add(Build(XLChartGroupKind.OfPie, ofPie, ofPie.Elements<C.PieChartSeries>()));
                    break;
                case C.DoughnutChart doughnut:
                    groups.Add(Build(XLChartGroupKind.Doughnut, doughnut, doughnut.Elements<C.PieChartSeries>()));
                    break;
                case C.AreaChart area:
                    groups.Add(Build(XLChartGroupKind.Area, area, area.Elements<C.AreaChartSeries>()));
                    break;
                case C.Area3DChart area3D:
                    groups.Add(Build(XLChartGroupKind.Area3D, area3D, area3D.Elements<C.AreaChartSeries>()));
                    break;
                case C.LineChart line:
                    groups.Add(Build(XLChartGroupKind.Line, line, line.Elements<C.LineChartSeries>()));
                    break;
                case C.Line3DChart line3D:
                    groups.Add(Build(XLChartGroupKind.Line3D, line3D, line3D.Elements<C.LineChartSeries>()));
                    break;
                case C.RadarChart radar:
                    groups.Add(Build(XLChartGroupKind.Radar, radar, radar.Elements<C.RadarChartSeries>()));
                    break;
                case C.BubbleChart bubble:
                    groups.Add(Build(XLChartGroupKind.Bubble, bubble, bubble.Elements<C.BubbleChartSeries>()));
                    break;
                case C.ScatterChart scatter:
                    groups.Add(Build(XLChartGroupKind.Scatter, scatter, scatter.Elements<C.ScatterChartSeries>()));
                    break;
                case C.StockChart stock:
                    groups.Add(Build(XLChartGroupKind.Stock, stock, stock.Elements<C.LineChartSeries>()));
                    break;
                case C.SurfaceChart surface:
                    groups.Add(Build(XLChartGroupKind.Surface, surface, surface.Elements<C.SurfaceChartSeries>()));
                    break;
                case C.Surface3DChart surface3D:
                    groups.Add(Build(XLChartGroupKind.Surface3D, surface3D,
                        surface3D.Elements<C.SurfaceChartSeries>()));
                    break;
            }
        }

        return groups;
    }

    /// <summary>
    /// Picks the group kind that represents the chart as a whole. Returns <c>null</c> for a plot area
    /// with no recognised group.
    /// </summary>
    internal static XLChartGroupKind? ChoosePrimaryKind(List<XLChartGroup> groups)
    {
        foreach (var kind in PrimaryPrecedence)
        {
            if (groups.Any(g => g.Kind == kind))
                return kind;
        }

        return null;
    }

    /// <summary>
    /// The group that carries the chart's primary axis pair. Groups plotted against a different value
    /// axis are on a secondary axis.
    /// </summary>
    /// <remarks>
    /// A plot area may hold several groups of the primary kind — two <c>c:barChart</c> elements, say,
    /// one per axis pair — and nothing in the schema says the primary one comes first. So rather than
    /// trust document order, a group whose value axis crosses the category axis at its maximum is
    /// passed over: that is how a secondary value axis ends up drawn on the right, and Excel writes it
    /// on every secondary axis it creates. When every candidate looks secondary — or none does —
    /// document order decides after all.
    /// </remarks>
    internal static XLChartGroup PrimaryGroup(
        C.PlotArea plotArea, List<XLChartGroup> groups, XLChartGroupKind primaryKind)
    {
        XLChartGroup? firstOfKind = null;

        foreach (var group in groups)
        {
            if (group.Kind != primaryKind)
                continue;

            firstOfKind ??= group;
            if (!CrossesAtMaximum(FindAxis(plotArea, group.ValueAxisId)))
                return group;
        }

        // ChoosePrimaryKind only ever returns a kind that is present, so there is always a candidate.
        return firstOfKind!;
    }

    /// <summary>
    /// Whether an axis element carries <c>c:crosses val="max"</c>, the mark of a secondary value axis.
    /// </summary>
    private static bool CrossesAtMaximum(OpenXmlCompositeElement? axis) =>
        axis?.Elements<C.Crosses>().FirstOrDefault()?.Val?.Value == C.CrossesValues.Maximum;

    /// <summary>
    /// Finds the axis element of a plot area carrying the given <c>c:axId</c>, whichever axis type it
    /// happens to be.
    /// </summary>
    internal static OpenXmlCompositeElement? FindAxis(C.PlotArea plotArea, uint? axisId)
    {
        if (axisId == null)
            return null;

        foreach (var child in plotArea.ChildElements)
        {
            if (child is not (C.CategoryAxis or C.ValueAxis or C.DateAxis or C.SeriesAxis))
                continue;

            var composite = (OpenXmlCompositeElement)child;
            if (composite.Elements<C.AxisId>().FirstOrDefault()?.Val?.Value == axisId)
                return composite;
        }

        return null;
    }

    private static XLChartGroup Build<TSeries>(
        XLChartGroupKind kind, OpenXmlCompositeElement element, IEnumerable<TSeries> seriesElements)
        where TSeries : OpenXmlCompositeElement
    {
        var axisIds = element.Elements<C.AxisId>().ToList();
        var categoryAxisId = axisIds.Count >= 1 ? axisIds[0].Val?.Value : null;
        var valueAxisId = axisIds.Count >= 2 ? axisIds[1].Val?.Value : null;
        return new XLChartGroup(kind, element, seriesElements.Cast<OpenXmlCompositeElement>().ToList(),
            categoryAxisId, valueAxisId);
    }
}
