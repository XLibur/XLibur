#nullable disable

using System.Collections.Generic;

namespace XLibur.Excel;

public interface IXLCharts : IEnumerable<IXLChart>
{
    void Add(IXLChart chart);
}
