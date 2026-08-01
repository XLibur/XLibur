using System.Collections.Generic;

namespace XLibur.Excel;

/// <summary>
/// A collection of data series belonging to a chart.
/// </summary>
public interface IXLChartSeriesCollection : IEnumerable<IXLChartSeries>
{
    /// <summary>
    /// Gets the number of series in the collection.
    /// </summary>
    int Count { get; }

    /// <summary>
    /// Creates a new data series, adds it to the collection, and returns it.
    /// The series <see cref="IXLChartSeries.Index"/> and <see cref="IXLChartSeries.Order"/> are assigned automatically.
    /// </summary>
    /// <param name="name">The display name of the series (shown in the chart legend).</param>
    /// <param name="valueReferences">Cell reference for the series values, e.g. <c>"Sheet1!$B$2:$B$5"</c>.</param>
    /// <param name="categoryReferences">Optional cell reference for category labels, e.g. <c>"Sheet1!$A$2:$A$5"</c>.</param>
    /// <returns>The newly created series.</returns>
    /// <exception cref="System.NotSupportedException">
    /// The chart was loaded from a file. XLibur patches the properties of an existing chart rather
    /// than regenerating its XML, so there is nowhere to write a new series; recreate the chart with
    /// <see cref="IXLCharts.Add(XLChartType)"/> instead.
    /// </exception>
    IXLChartSeries Add(string name, string valueReferences, string? categoryReferences = null);
}
