using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel.CalcEngine.Visitors;

namespace XLibur.Excel;

internal sealed class XLCharts : IXLCharts, IWorkbookListener
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

    IEnumerator IEnumerable.GetEnumerator()
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

    /// <summary>
    /// Every series names the renamed sheet by its new name, in its values, its categories and the
    /// cell its name comes from. The <c>rename-*</c> fixture shows Excel renaming all three (D65).
    /// </summary>
    void IWorkbookListener.OnSheetRenamed(string oldSheetName, string newSheetName)
        => RewriteSheet(SheetRewrite.Rename(oldSheetName, newSheetName));

    /// <summary>
    /// Each series reference to the deleted sheet becomes <c>#REF!</c>, where it stands.
    /// </summary>
    /// <remarks>
    /// Excel writes the values as <c>#REF!</c> too, but moves the series name and the categories into
    /// <c>c15:filteredSeriesTitle</c> and <c>c15:filteredCategoryTitle</c> extensions, each
    /// <c>#REF!</c> (the <c>delete-*</c> fixture). XLibur does not move them: the text matches Excel's
    /// and the XML does not, a call recorded in spec 55's Results.
    /// </remarks>
    void IWorkbookListener.OnSheetDeleting(string sheetName)
        => RewriteSheet(SheetRewrite.Delete(_worksheet.Workbook, sheetName));

    private void RewriteSheet(SheetRewrite rewrite)
    {
        foreach (var chart in _charts.OfType<XLChart>())
        {
            foreach (var series in chart.SeriesInternal.Items)
                series.RewriteSheet(_worksheet.Name, rewrite);

            foreach (var series in chart.SecondarySeriesInternal.Items)
                series.RewriteSheet(_worksheet.Name, rewrite);
        }
    }
}
