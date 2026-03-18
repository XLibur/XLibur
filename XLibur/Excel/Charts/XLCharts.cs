using System.Collections.Generic;

namespace XLibur.Excel;

internal sealed class XLCharts : IXLCharts
{
    private readonly List<IXLChart> _charts = [];
    private readonly XLWorksheet _worksheet;

    internal XLCharts(XLWorksheet worksheet)
    {
        _worksheet = worksheet;
    }

    public int Count => _charts.Count;

    public IEnumerator<IXLChart> GetEnumerator()
    {
        return _charts.GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public void Add(IXLChart chart)
    {
        _charts.Add(chart);
    }

    public IXLChart Add(XLChartType chartType)
    {
        var chart = new XLChart(_worksheet)
        {
            ChartType = chartType
        };
        _charts.Add(chart);
        return chart;
    }
}
