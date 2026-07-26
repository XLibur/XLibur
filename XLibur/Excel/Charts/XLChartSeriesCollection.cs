using System.Collections;
using System.Collections.Generic;

namespace XLibur.Excel;

internal sealed class XLChartSeriesCollection : IXLChartSeriesCollection
{
    private readonly List<XLChartSeries> _series = [];
    private readonly XLChart _chart;

    internal XLChartSeriesCollection(XLChart chart)
    {
        _chart = chart;
    }

    public int Count => _series.Count;

    /// <summary>
    /// The series of this collection, in plot order.
    /// </summary>
    internal IReadOnlyList<XLChartSeries> Items => _series;

    public IXLChartSeries Add(string name, string valueReferences, string? categoryReferences = null)
    {
        var series = new XLChartSeries(_chart)
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
