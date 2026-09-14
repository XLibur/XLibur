using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel.CalcEngine.Visitors;
using XLibur.Excel.Coordinates;
using XLibur.Excel.IO;

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
    /// <remarks>
    /// A ChartEx chart Excel wrote names a hidden name in each reference, and the name is renamed
    /// instead: the <c>chartex-pivotcf-*</c> fixture leaves the chart part as it was.
    /// </remarks>
    void IWorkbookListener.OnSheetRenamed(string oldSheetName, string newSheetName)
        => RewriteSheet(SheetRewrite.Rename(oldSheetName, newSheetName));

    /// <summary>
    /// Each series reference to the deleted sheet becomes <c>#REF!</c>, where it stands.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Excel writes the values as <c>#REF!</c> too, but moves the series name and the categories into
    /// <c>c15:filteredSeriesTitle</c> and <c>c15:filteredCategoryTitle</c> extensions, each
    /// <c>#REF!</c> (the <c>delete-*</c> fixture). XLibur does not move them: the text matches Excel's
    /// and the XML does not, a call recorded in spec 55's Results.
    /// </para>
    /// <para>
    /// A ChartEx chart Excel wrote holds its references in hidden names, and Excel removes the names
    /// that referred to the deleted sheet and the references that named them (the
    /// <c>chartex-pivotcf-*</c> fixture, and see <see cref="DropChartDataNames"/>).
    /// </para>
    /// </remarks>
    void IWorkbookListener.OnSheetDeleting(string sheetName)
    {
        RewriteSheet(SheetRewrite.Delete(_worksheet.Workbook, sheetName));
        DropChartDataNames(_worksheet.Workbook.DefinedNamesInternal.ChartDataNamesGoingWithSheet);
    }

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

    /// <summary>
    /// Takes each reference that names one of <paramref name="going"/> out of a loaded ChartEx chart
    /// (see <see cref="XLChartSeries.DropChartDataNames"/>). The door decided which names go before any
    /// holder heard of the delete (see <see cref="XLDefinedNames.FindChartDataNamesGoingWithSheet"/>).
    /// A chart created in code holds its references itself, as a standard chart does.
    /// </summary>
    private void DropChartDataNames(IReadOnlyDictionary<string, Area> going)
    {
        if (going.Count == 0)
            return;

        foreach (var chart in _charts.OfType<XLChart>())
        {
            if (!chart.LoadedFromFile || !ChartWriter.IsExtendedType(chart.ChartType))
                continue;

            foreach (var series in chart.SeriesInternal.Items)
                series.DropChartDataNames(going);
        }
    }
}
