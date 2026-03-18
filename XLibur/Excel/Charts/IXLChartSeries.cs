namespace XLibur.Excel;

/// <summary>
/// Represents a single data series within a chart.
/// </summary>
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
}
