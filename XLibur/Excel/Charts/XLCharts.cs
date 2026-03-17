using System.Collections.Generic;

namespace XLibur.Excel;

internal sealed class XLCharts : IXLCharts
{
    private readonly List<IXLChart> _charts = [];
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
}
