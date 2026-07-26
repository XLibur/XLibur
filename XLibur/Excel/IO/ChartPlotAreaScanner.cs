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
    Doughnut,
    Area,
    Line,
    Radar,
    Bubble,
    Scatter,
    Stock,
    Surface
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
        uint? valueAxisId)
    {
        Kind = kind;
        Element = element;
        SeriesElements = seriesElements;
        ValueAxisId = valueAxisId;
    }

    internal XLChartGroupKind Kind { get; }

    internal OpenXmlCompositeElement Element { get; }

    internal List<OpenXmlCompositeElement> SeriesElements { get; }

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
        XLChartGroupKind.Doughnut,
        XLChartGroupKind.Area,
        XLChartGroupKind.Line,
        XLChartGroupKind.Radar,
        XLChartGroupKind.Bubble,
        XLChartGroupKind.Scatter,
        XLChartGroupKind.Stock,
        XLChartGroupKind.Surface
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
                case C.DoughnutChart doughnut:
                    groups.Add(Build(XLChartGroupKind.Doughnut, doughnut, doughnut.Elements<C.PieChartSeries>()));
                    break;
                case C.AreaChart area:
                    groups.Add(Build(XLChartGroupKind.Area, area, area.Elements<C.AreaChartSeries>()));
                    break;
                case C.LineChart line:
                    groups.Add(Build(XLChartGroupKind.Line, line, line.Elements<C.LineChartSeries>()));
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
    /// The value axis of the first group of the primary kind. Groups plotted against a different
    /// value axis are on a secondary axis.
    /// </summary>
    internal static uint? PrimaryValueAxisId(List<XLChartGroup> groups, XLChartGroupKind primaryKind) =>
        groups.First(g => g.Kind == primaryKind).ValueAxisId;

    private static XLChartGroup Build<TSeries>(
        XLChartGroupKind kind, OpenXmlCompositeElement element, IEnumerable<TSeries> seriesElements)
        where TSeries : OpenXmlCompositeElement
    {
        var axisIds = element.Elements<C.AxisId>().ToList();
        var valueAxisId = axisIds.Count >= 2 ? axisIds[1].Val?.Value : null;
        return new XLChartGroup(kind, element, seriesElements.Cast<OpenXmlCompositeElement>().ToList(), valueAxisId);
    }
}
