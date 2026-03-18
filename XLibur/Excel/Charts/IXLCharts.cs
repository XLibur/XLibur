using System.Collections.Generic;

namespace XLibur.Excel;

/// <summary>
/// A collection of charts embedded in a worksheet.
/// </summary>
public interface IXLCharts : IEnumerable<IXLChart>
{
    /// <summary>
    /// Gets the number of charts in the collection.
    /// </summary>
    int Count { get; }

    /// <summary>
    /// Adds an existing chart instance to the collection.
    /// </summary>
    /// <param name="chart">The chart to add.</param>
    void Add(IXLChart chart);

    /// <summary>
    /// Creates a new chart of the specified type, adds it to the collection, and returns it.
    /// </summary>
    /// <param name="chartType">The type of chart to create (e.g. ColumnClustered, BarStacked).</param>
    /// <returns>The newly created chart.</returns>
    IXLChart Add(XLChartType chartType);
}
