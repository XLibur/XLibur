using System.Collections.Generic;

namespace XLibur.Excel;

public interface IXLChartSeriesCollection : IEnumerable<IXLChartSeries>
{
    int Count { get; }
    IXLChartSeries Add(string name, string valueReferences, string? categoryReferences = null);
}
