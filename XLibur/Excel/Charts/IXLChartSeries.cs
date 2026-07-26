namespace XLibur.Excel;

/// <summary>
/// Represents a single data series within a chart.
/// </summary>
/// <remarks>
/// The formatting properties (<see cref="FillColor"/>, <see cref="LineColor"/>,
/// <see cref="LineWidthPt"/>, <see cref="MarkerStyle"/>, <see cref="MarkerSize"/>,
/// <see cref="MarkerFillColor"/> and <see cref="Smooth"/>) apply to standard chart types only.
/// The extended (Office 2016+) types — Sunburst, Treemap, Waterfall, Funnel and Box &amp; Whisker —
/// ignore them.
/// </remarks>
public interface IXLChartSeries
{
    /// <summary>
    /// Gets or sets the display name of the series (shown in the chart legend).
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// Gets or sets the cell reference for category (axis) labels, e.g. <c>"Sheet1!$A$2:$A$5"</c>.
    /// Set to <c>null</c> if the series has no explicit category data.
    /// </summary>
    string? CategoryReferences { get; set; }

    /// <summary>
    /// Gets or sets the cell reference for the series values, e.g. <c>"Sheet1!$B$2:$B$5"</c>.
    /// </summary>
    string ValueReferences { get; set; }

    /// <summary>
    /// Gets the zero-based index that uniquely identifies this series within the chart.
    /// </summary>
    uint Index { get; }

    /// <summary>
    /// Gets the zero-based plot order of this series within the chart.
    /// </summary>
    uint Order { get; }

    /// <summary>
    /// Gets or sets the solid fill colour of the series interior — the bar, column, area or pie
    /// body. Line and scatter series have no interior; use <see cref="MarkerFillColor"/> for their
    /// markers. <c>null</c> — the default — omits the fill, so Excel applies the automatic colour
    /// from the workbook theme.
    /// </summary>
    /// <remarks>
    /// A theme colour (<see cref="XLColor.FromTheme(XLThemeColor)"/>) is written as a DrawingML
    /// scheme colour; its <see cref="XLColor.ThemeTint"/> is not applied.
    /// </remarks>
    XLColor? FillColor { get; set; }

    /// <summary>
    /// Gets or sets the colour of the series outline — the line itself for line, scatter and radar
    /// series, the border for bar, column, area and pie series. <c>null</c> — the default — omits
    /// the outline colour, so Excel applies the automatic colour.
    /// </summary>
    XLColor? LineColor { get; set; }

    /// <summary>
    /// Gets or sets the width of the series outline in points. <c>null</c> — the default — omits
    /// the width, so Excel applies its own.
    /// </summary>
    /// <exception cref="System.ArgumentOutOfRangeException">
    /// The value is negative or greater than 1584 (the maximum line width Excel accepts).
    /// </exception>
    double? LineWidthPt { get; set; }

    /// <summary>
    /// Gets or sets the marker symbol drawn at each data point.
    /// Defaults to <see cref="XLMarkerStyle.Auto"/>, which leaves the choice to Excel.
    /// Only line, scatter and radar series draw markers.
    /// </summary>
    XLMarkerStyle MarkerStyle { get; set; }

    /// <summary>
    /// Gets or sets the marker size in points. <c>null</c> — the default — lets Excel choose.
    /// </summary>
    /// <exception cref="System.ArgumentOutOfRangeException">
    /// The value is outside the range Excel accepts (2 to 72 points).
    /// </exception>
    double? MarkerSize { get; set; }

    /// <summary>
    /// Gets or sets the fill colour of the marker interior.
    /// <c>null</c> — the default — lets Excel choose.
    /// </summary>
    XLColor? MarkerFillColor { get; set; }

    /// <summary>
    /// Gets or sets whether the line connecting the data points is smoothed.
    /// Applies to line and scatter series. When <c>false</c> — the default — no smoothing element
    /// is written and the chart type's own default applies (the <c>XYScatterSmoothLines*</c> types
    /// are smoothed, everything else is not).
    /// </summary>
    bool Smooth { get; set; }

    /// <summary>
    /// Gets or sets whether the series is plotted against a secondary value axis, drawn on the
    /// right-hand side of the plot area.
    /// </summary>
    /// <remarks>
    /// Honoured for chart types with one category and one value axis (bar, column, line, area,
    /// radar and stock). Chart types without a value axis (pie, doughnut), the two-value-axis types
    /// (scatter, bubble) and surface — which adds a series axis — ignore it.
    /// </remarks>
    /// <exception cref="System.NotSupportedException">
    /// The chart was loaded from a file. Moving a series of an existing chart onto a secondary
    /// axis would require regrouping the chart, which XLibur does not do; recreate the chart with
    /// <see cref="IXLCharts.Add(XLChartType)"/> instead.
    /// </exception>
    bool UseSecondaryAxis { get; set; }

    /// <summary>
    /// Gets the data labels of this series. Anything left untouched here falls back to the
    /// chart-wide <see cref="IXLChart.DataLabels"/>.
    /// </summary>
    IXLDataLabels DataLabels { get; }
}
