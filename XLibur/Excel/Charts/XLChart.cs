using System;
using System.Collections.Generic;
using System.Linq;

namespace XLibur.Excel;

internal enum XLChartTypeCategory
{
    Bar3D
}

internal enum XLBarOrientation
{
    Vertical,
    Horizontal
}

internal enum XLBarGrouping
{
    Clustered,
    Percent,
    Stacked,
    Standard
}

/// <summary>
/// Internal implementation of <see cref="IXLChart"/>. Holds chart metadata (type, title, series,
/// positions) and save/load coordination state (<see cref="IsNew"/>, <see cref="RelId"/>).
/// </summary>
internal sealed class XLChart : XLDrawing<IXLChart>, IXLChart
{
    /// <summary>
    /// Creates a new chart belonging to the specified worksheet.
    /// Automatically assigns ZOrder and ShapeId.
    /// </summary>
    public XLChart(XLWorksheet worksheet)
    {
        Container = this;
        Worksheet = worksheet;
        int zOrder;
        if (worksheet.Charts.Count != 0)
            zOrder = worksheet.Charts.Max(c => c.ZOrder) + 1;
        else
            zOrder = 1;
        ZOrder = zOrder;
        ShapeId = worksheet.Workbook.ShapeIdManager.GetNext();
        RightAngleAxes = true;
        Series = new XLChartSeriesCollection();
        SecondarySeries = new XLChartSeriesCollection();
        SecondPosition = new XLDrawingPosition();
    }

    public string? Title { get; set; }

    public IXLChart SetTitle(string? title)
    {
        Title = title;
        return this;
    }

    public IXLChartSeriesCollection Series { get; }

    public IXLWorksheet Worksheet { get; }

    public IXLDrawingPosition SecondPosition { get; }

    public XLChartType? SecondaryChartType { get; set; }

    public IXLChartSeriesCollection SecondarySeries { get; }

    /// <summary>
    /// The relationship ID linking this chart to its ChartPart within the DrawingsPart.
    /// Set during load; <c>null</c> for newly created charts until save.
    /// </summary>
    internal string? RelId { get; set; }

    /// <summary>
    /// Whether this chart was created programmatically (<c>true</c>) or loaded from a file (<c>false</c>).
    /// ChartWriter only writes charts where this is <c>true</c>.
    /// </summary>
    internal bool IsNew { get; set; } = true;

    public bool RightAngleAxes { get; set; }

    public IXLChart SetRightAngleAxes()
    {
        RightAngleAxes = true;
        return this;
    }

    public IXLChart SetRightAngleAxes(bool rightAngleAxes)
    {
        RightAngleAxes = rightAngleAxes;
        return this;
    }

    public XLChartType ChartType { get; set; }

    public IXLChart SetChartType(XLChartType chartType)
    {
        ChartType = chartType;
        return this;
    }

    /// <summary>
    /// Gets the broad category of the chart type. Currently only supports Bar3D;
    /// throws <see cref="NotImplementedException"/> for other categories.
    /// </summary>
    public XLChartTypeCategory ChartTypeCategory => Bar3DCharts.Contains(ChartType)
        ? XLChartTypeCategory.Bar3D
        : throw new NotImplementedException();

    private readonly HashSet<XLChartType> Bar3DCharts =
    [
        XLChartType.BarClustered3D,
        XLChartType.BarStacked100Percent3D,
        XLChartType.BarStacked3D,
        XLChartType.Column3D,
        XLChartType.ColumnClustered3D,
        XLChartType.ColumnStacked100Percent3D,
        XLChartType.ColumnStacked3D
    ];

    /// <summary>
    /// Gets whether this chart type renders bars horizontally or vertically,
    /// based on the <see cref="ChartType"/>.
    /// </summary>
    public XLBarOrientation BarOrientation => HorizontalCharts.Contains(ChartType)
        ? XLBarOrientation.Horizontal
        : XLBarOrientation.Vertical;

    private readonly HashSet<XLChartType> HorizontalCharts =
    [
        XLChartType.BarClustered,
        XLChartType.BarClustered3D,
        XLChartType.BarStacked,
        XLChartType.BarStacked100Percent,
        XLChartType.BarStacked100Percent3D,
        XLChartType.BarStacked3D,
        XLChartType.ConeHorizontalClustered,
        XLChartType.ConeHorizontalStacked,
        XLChartType.ConeHorizontalStacked100Percent,
        XLChartType.CylinderHorizontalClustered,
        XLChartType.CylinderHorizontalStacked,
        XLChartType.CylinderHorizontalStacked100Percent,
        XLChartType.PyramidHorizontalClustered,
        XLChartType.PyramidHorizontalStacked,
        XLChartType.PyramidHorizontalStacked100Percent
    ];

    /// <summary>
    /// Gets the bar grouping style (Clustered, Stacked, Percent, or Standard)
    /// derived from the <see cref="ChartType"/>.
    /// </summary>
    public XLBarGrouping BarGrouping
    {
        get
        {
            if (ClusteredCharts.Contains(ChartType))
                return XLBarGrouping.Clustered;
            if (PercentCharts.Contains(ChartType))
                return XLBarGrouping.Percent;
            return StackedCharts.Contains(ChartType) ? XLBarGrouping.Stacked : XLBarGrouping.Standard;
        }
    }

    public readonly HashSet<XLChartType> ClusteredCharts =
    [
        XLChartType.BarClustered,
        XLChartType.BarClustered3D,
        XLChartType.ColumnClustered,
        XLChartType.ColumnClustered3D,
        XLChartType.ConeClustered,
        XLChartType.ConeHorizontalClustered,
        XLChartType.CylinderClustered,
        XLChartType.CylinderHorizontalClustered,
        XLChartType.PyramidClustered,
        XLChartType.PyramidHorizontalClustered
    ];

    public readonly HashSet<XLChartType> PercentCharts =
    [
        XLChartType.AreaStacked100Percent,
        XLChartType.AreaStacked100Percent3D,
        XLChartType.BarStacked100Percent,
        XLChartType.BarStacked100Percent3D,
        XLChartType.ColumnStacked100Percent,
        XLChartType.ColumnStacked100Percent3D,
        XLChartType.ConeHorizontalStacked100Percent,
        XLChartType.ConeStacked100Percent,
        XLChartType.CylinderHorizontalStacked100Percent,
        XLChartType.CylinderStacked100Percent,
        XLChartType.LineStacked100Percent,
        XLChartType.LineWithMarkersStacked100Percent,
        XLChartType.PyramidHorizontalStacked100Percent,
        XLChartType.PyramidStacked100Percent
    ];

    public readonly HashSet<XLChartType> StackedCharts =
    [
        XLChartType.AreaStacked,
        XLChartType.AreaStacked3D,
        XLChartType.BarStacked,
        XLChartType.BarStacked3D,
        XLChartType.ColumnStacked,
        XLChartType.ColumnStacked3D,
        XLChartType.ConeHorizontalStacked,
        XLChartType.ConeStacked,
        XLChartType.CylinderHorizontalStacked,
        XLChartType.CylinderStacked,
        XLChartType.LineStacked,
        XLChartType.LineWithMarkersStacked,
        XLChartType.PyramidHorizontalStacked,
        XLChartType.PyramidStacked
    ];
}
