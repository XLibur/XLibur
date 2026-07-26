using System;
using System.Collections;
using System.Collections.Generic;

namespace XLibur.Excel;

internal sealed class XLChartSeriesCollection : IXLChartSeriesCollection
{
    private readonly List<XLChartSeries> _series = [];
    private readonly XLChart _chart;
    private readonly bool _secondary;

    /// <param name="chart">The chart the series belong to.</param>
    /// <param name="secondary">
    /// <c>true</c> for the combo chart's secondary collection, whose series follow
    /// <see cref="IXLChart.SecondaryChartType"/>.
    /// </param>
    internal XLChartSeriesCollection(XLChart chart, bool secondary = false)
    {
        _chart = chart;
        _secondary = secondary;
    }

    public int Count => _series.Count;

    /// <summary>
    /// The series of this collection, in plot order.
    /// </summary>
    internal IReadOnlyList<XLChartSeries> Items => _series;

    public IXLChartSeries Add(string name, string valueReferences, string? categoryReferences = null)
    {
        if (_chart.LoadedFromFile)
        {
            throw new NotSupportedException(
                "Cannot add a series to a chart loaded from a file. XLibur patches the properties of "
                + "an existing chart rather than regenerating its XML, so a new series has nowhere to "
                + "be written; recreate the chart with IXLCharts.Add(XLChartType) instead.");
        }

        var series = new XLChartSeries(_chart, _secondary)
        {
            Name = name,
            ValueReferences = valueReferences,
            CategoryReferences = categoryReferences,
            Index = (uint)_series.Count,
            Order = (uint)_series.Count
        };
        _series.Add(series);
        return series;
    }

    public IEnumerator<IXLChartSeries> GetEnumerator() => _series.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
