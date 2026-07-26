namespace XLibur.Excel;

public enum XLChartType
{
    Area,
    Area3D,
    AreaStacked,
    AreaStacked100Percent,
    AreaStacked100Percent3D,
    AreaStacked3D,
    BarClustered,
    BarClustered3D,
    BarStacked,
    BarStacked100Percent,
    BarStacked100Percent3D,
    BarStacked3D,
    Bubble,
    Bubble3D,
    Column3D,
    ColumnClustered,
    ColumnClustered3D,
    ColumnStacked,
    ColumnStacked100Percent,
    ColumnStacked100Percent3D,
    ColumnStacked3D,
    Cone,
    ConeClustered,
    ConeHorizontalClustered,
    ConeHorizontalStacked,
    ConeHorizontalStacked100Percent,
    ConeStacked,
    ConeStacked100Percent,
    Cylinder,
    CylinderClustered,
    CylinderHorizontalClustered,
    CylinderHorizontalStacked,
    CylinderHorizontalStacked100Percent,
    CylinderStacked,
    CylinderStacked100Percent,
    Doughnut,
    DoughnutExploded,
    Line,
    Line3D,
    LineStacked,
    LineStacked100Percent,
    LineWithMarkers,
    LineWithMarkersStacked,
    LineWithMarkersStacked100Percent,
    Pie,
    Pie3D,
    PieExploded,
    PieExploded3D,
    PieToBar,
    PieToPie,
    Pyramid,
    PyramidClustered,
    PyramidHorizontalClustered,
    PyramidHorizontalStacked,
    PyramidHorizontalStacked100Percent,
    PyramidStacked,
    PyramidStacked100Percent,
    Radar,
    RadarFilled,
    RadarWithMarkers,
    StockHighLowClose,
    StockOpenHighLowClose,
    StockVolumeHighLowClose,
    StockVolumeOpenHighLowClose,
    Surface,
    SurfaceContour,
    SurfaceContourWireframe,
    SurfaceWireframe,
    XYScatterMarkers,
    XYScatterSmoothLinesNoMarkers,
    XYScatterSmoothLinesWithMarkers,
    XYScatterStraightLinesNoMarkers,
    XYScatterStraightLinesWithMarkers,

    // Extended chart types (Office 2016+, uses ExtendedChartPart / cx namespace)
    BoxWhisker,
    Funnel,
    Sunburst,
    Treemap,
    Waterfall
}
/// <summary>
/// Represents an Excel chart embedded in a worksheet.
/// </summary>
public interface IXLChart : IXLDrawing<IXLChart>
{
    /// <summary>
    /// Gets or sets whether the chart axes are displayed at right angles, independent of chart rotation.
    /// </summary>
    bool RightAngleAxes { get; set; }

    /// <summary>
    /// Sets <see cref="RightAngleAxes"/> to <c>true</c>.
    /// </summary>
    /// <returns>The current chart instance for fluent chaining.</returns>
    IXLChart SetRightAngleAxes();

    /// <summary>
    /// Sets <see cref="RightAngleAxes"/> to the specified value.
    /// </summary>
    /// <param name="rightAngleAxes">Whether to display axes at right angles.</param>
    /// <returns>The current chart instance for fluent chaining.</returns>
    IXLChart SetRightAngleAxes(bool rightAngleAxes);

    /// <summary>
    /// Gets or sets the chart type (e.g. ColumnClustered, BarStacked, Line).
    /// </summary>
    XLChartType ChartType { get; set; }

    /// <summary>
    /// Sets <see cref="ChartType"/> to the specified value.
    /// </summary>
    /// <param name="chartType">The chart type to apply.</param>
    /// <returns>The current chart instance for fluent chaining.</returns>
    IXLChart SetChartType(XLChartType chartType);

    /// <summary>
    /// Gets or sets the chart title text. Set to <c>null</c> to remove the title.
    /// </summary>
    string? Title { get; set; }

    /// <summary>
    /// Sets <see cref="Title"/> to the specified value.
    /// </summary>
    /// <param name="title">The title text, or <c>null</c> to remove the title.</param>
    /// <returns>The current chart instance for fluent chaining.</returns>
    IXLChart SetTitle(string? title);

    /// <summary>
    /// Gets the collection of data series plotted by this chart.
    /// </summary>
    IXLChartSeriesCollection Series { get; }

    /// <summary>
    /// Gets the worksheet that contains this chart.
    /// </summary>
    IXLWorksheet Worksheet { get; }

    /// <summary>
    /// Gets the bottom-right anchor position of the chart's two-cell anchor.
    /// Use together with <see cref="IXLDrawing{T}.Position"/> (top-left) to define the chart's size and location.
    /// Only used while <see cref="Anchor"/> is <see cref="XLDrawingAnchor.MoveAndSizeWithCells"/>.
    /// </summary>
    IXLDrawingPosition SecondPosition { get; }

    /// <summary>
    /// Gets or sets how the chart is tied to the grid.
    /// </summary>
    /// <remarks>
    /// <list type="bullet">
    /// <item>
    /// <see cref="XLDrawingAnchor.MoveAndSizeWithCells"/> — the default — spans the rectangle between
    /// <see cref="IXLDrawing{T}.Position"/> and <see cref="SecondPosition"/>, so inserting rows or
    /// columns moves and resizes the chart.
    /// </item>
    /// <item>
    /// <see cref="XLDrawingAnchor.MoveWithCells"/> keeps <see cref="Width"/> and <see cref="Height"/>
    /// and hangs the top-left corner off <see cref="IXLDrawing{T}.Position"/>.
    /// </item>
    /// <item>
    /// <see cref="XLDrawingAnchor.Absolute"/> pins the chart to <see cref="Left"/>/<see cref="Top"/>
    /// with <see cref="Width"/> and <see cref="Height"/>, ignoring the grid entirely.
    /// </item>
    /// </list>
    /// </remarks>
    XLDrawingAnchor Anchor { get; set; }

    /// <summary>
    /// Gets or sets the chart width in pixels. Used when <see cref="Anchor"/> is
    /// <see cref="XLDrawingAnchor.MoveWithCells"/> or <see cref="XLDrawingAnchor.Absolute"/>; a
    /// two-cell anchored chart takes its width from <see cref="SecondPosition"/> instead.
    /// </summary>
    int Width { get; set; }

    /// <summary>
    /// Gets or sets the chart height in pixels. Used when <see cref="Anchor"/> is
    /// <see cref="XLDrawingAnchor.MoveWithCells"/> or <see cref="XLDrawingAnchor.Absolute"/>.
    /// </summary>
    int Height { get; set; }

    /// <summary>
    /// Gets or sets the distance in pixels from the left edge of the sheet. Used only when
    /// <see cref="Anchor"/> is <see cref="XLDrawingAnchor.Absolute"/>.
    /// </summary>
    int Left { get; set; }

    /// <summary>
    /// Gets or sets the distance in pixels from the top edge of the sheet. Used only when
    /// <see cref="Anchor"/> is <see cref="XLDrawingAnchor.Absolute"/>.
    /// </summary>
    int Top { get; set; }

    /// <summary>
    /// Gets or sets the secondary chart type for combo charts.
    /// When set, the chart renders both <see cref="ChartType"/> and this type
    /// in the same plot area, each with its own series. Set to <c>null</c> for single-type charts.
    /// </summary>
    /// <remarks>
    /// The model holds two chart types, which is all XLibur writes. A chart read from a file may mix
    /// more: Excel allows a plot area to combine any number of chart groups. In that case
    /// <see cref="ChartType"/> and this property name the first two, every series outside the primary
    /// group lands in <see cref="SecondarySeries"/>, and the series belonging to a third or later
    /// group are reported under this type rather than their own. Saving is unaffected — a loaded chart
    /// keeps its original XML — but the model understates such a chart.
    /// </remarks>
    XLChartType? SecondaryChartType { get; set; }

    /// <summary>
    /// Gets the collection of data series for the secondary chart type in a combo chart.
    /// Only used when <see cref="SecondaryChartType"/> is set.
    /// </summary>
    IXLChartSeriesCollection SecondarySeries { get; }

    /// <summary>
    /// Gets the data labels that apply to every series of the chart. A series can override them
    /// through <see cref="IXLChartSeries.DataLabels"/>.
    /// </summary>
    IXLDataLabels DataLabels { get; }

    /// <summary>
    /// Gets the chart's legend.
    /// </summary>
    IXLChartLegend Legend { get; }

    /// <summary>
    /// Gets the horizontal axis — the category axis, or the X value axis of a scatter or bubble chart.
    /// Pie and doughnut charts have no axes and ignore it.
    /// </summary>
    IXLChartAxis CategoryAxis { get; }

    /// <summary>
    /// Gets the vertical value axis. Pie and doughnut charts have no axes and ignore it.
    /// </summary>
    IXLChartAxis ValueAxis { get; }

    /// <summary>
    /// Gets the value axis on the right, which exists only while at least one series has
    /// <see cref="IXLChartSeries.UseSecondaryAxis"/> set.
    /// </summary>
    IXLChartAxis SecondaryValueAxis { get; }
}
