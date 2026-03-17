using System.Collections.Generic;

namespace XLibur.Excel;

public interface IXLCharts : IEnumerable<IXLChart>
{
    int Count { get; }
    void Add(IXLChart chart);
    IXLChart Add(XLChartType chartType);
}
