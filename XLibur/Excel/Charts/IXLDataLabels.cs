namespace XLibur.Excel;

/// <summary>
/// Where a data label sits relative to its data point.
/// </summary>
/// <remarks>
/// Excel only offers a subset of these for any given chart type, and rejects a file that uses one it
/// does not offer. <see cref="IXLDataLabels.Position"/> validates the combination when it is set.
/// </remarks>
public enum XLDataLabelPosition
{
    /// <summary>
    /// No explicit position: the label element is omitted and Excel places the label itself.
    /// This is the default, and the only value the chart types that do not offer a position accept.
    /// </summary>
    Auto,

    /// <summary>Centred on the data point.</summary>
    Center,

    /// <summary>Inside the end of the bar, column or slice.</summary>
    InsideEnd,

    /// <summary>Inside the base of the bar or column.</summary>
    InsideBase,

    /// <summary>Outside the end of the bar, column or slice.</summary>
    OutsideEnd,

    /// <summary>Wherever it fits best — pie charts only.</summary>
    BestFit,

    /// <summary>To the left of the marker.</summary>
    Left,

    /// <summary>To the right of the marker.</summary>
    Right,

    /// <summary>Above the marker.</summary>
    Above,

    /// <summary>Below the marker.</summary>
    Below
}

/// <summary>
/// The labels drawn next to the data points of a chart or one of its series.
/// </summary>
/// <remarks>
/// <para>
/// Every property is off or automatic by default, and nothing is written to the file until one of
/// them is set — so an untouched chart keeps whatever Excel's own defaults are.
/// </para>
/// <para>
/// Series labels win over the chart-level <see cref="IXLChart.DataLabels"/>: set the chart's to give
/// every series the same labels, then override the odd series that needs something different.
/// </para>
/// <para>
/// Surface charts and the extended (Office 2016+) types — Sunburst, Treemap, Waterfall, Funnel and
/// Box &amp; Whisker — have no data label support and ignore these properties.
/// </para>
/// </remarks>
public interface IXLDataLabels
{
    /// <summary>
    /// Gets or sets whether the point's value is shown. Defaults to <c>false</c>.
    /// </summary>
    bool ShowValue { get; set; }

    /// <summary>
    /// Gets or sets whether the point's category name is shown. Defaults to <c>false</c>.
    /// </summary>
    bool ShowCategoryName { get; set; }

    /// <summary>
    /// Gets or sets whether the series name is shown. Defaults to <c>false</c>.
    /// </summary>
    bool ShowSeriesName { get; set; }

    /// <summary>
    /// Gets or sets whether the point's share of the total is shown as a percentage.
    /// Pie and doughnut charts only. Defaults to <c>false</c>.
    /// </summary>
    bool ShowPercentage { get; set; }

    /// <summary>
    /// Gets or sets the number format the labels are drawn with, e.g. <c>"#,##0"</c> or <c>"0.0%"</c>.
    /// <c>null</c> — the default — takes the format from the source cells.
    /// </summary>
    string? NumberFormat { get; set; }

    /// <summary>
    /// Gets or sets where the labels sit. Defaults to <see cref="XLDataLabelPosition.Auto"/>.
    /// </summary>
    /// <remarks>
    /// What Excel accepts depends on the chart type:
    /// <list type="bullet">
    /// <item>clustered bar and column: <c>Center</c>, <c>InsideEnd</c>, <c>InsideBase</c>, <c>OutsideEnd</c></item>
    /// <item>stacked bar and column: <c>Center</c>, <c>InsideEnd</c>, <c>InsideBase</c></item>
    /// <item>line, scatter and radar: <c>Center</c>, <c>Left</c>, <c>Right</c>, <c>Above</c>, <c>Below</c></item>
    /// <item>pie: <c>BestFit</c>, <c>Center</c>, <c>InsideEnd</c>, <c>OutsideEnd</c></item>
    /// <item>everything else — area, doughnut, bubble, stock and every 3D type — <c>Auto</c> only</item>
    /// </list>
    /// </remarks>
    /// <exception cref="System.ArgumentException">
    /// The position is one the chart's current <see cref="IXLChart.ChartType"/> does not offer.
    /// </exception>
    XLDataLabelPosition Position { get; set; }
}
