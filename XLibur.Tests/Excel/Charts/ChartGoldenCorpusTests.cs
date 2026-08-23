using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Charts;

/// <summary>
/// Pins the exact chart-part XML XLibur writes for a representative set of new charts. Spec 22
/// reorganises the code that produces this XML without changing it; any diff here is a finding to
/// investigate, never noise to re-baseline without a written explanation.
/// </summary>
/// <remarks>
/// <para>
/// A missing fixture is written rather than asserted, so adding a case to <see cref="Fixtures"/> and
/// running the test once is how the corpus is widened. A fixture that exists is only ever compared.
/// </para>
/// <para>
/// One fixture has been re-baselined since it was first captured. Spec 22 task 2 gave
/// <c>bar-titled</c> a <c>&lt;c:autoTitleDeleted val="0"/&gt;</c> it did not have before: unifying
/// the build and patch paths of the title means a new chart now keeps that element in step the way
/// a loaded chart always did. The element is what Excel itself writes and repeats the default, so
/// the chart renders identically; the divergence it removes is the whole point of the task.
/// </para>
/// </remarks>
public class ChartGoldenCorpusTests
{
    private const string Values = "Data!$B$1:$B$2";
    private const string Categories = "Data!$A$1:$A$2";

    [Test]
    [Arguments("bar-plain")]
    [Arguments("line-legend-bottom")]
    [Arguments("bar-titled")]
    [Arguments("line-datalabels")]
    [Arguments("bar-secondary-axis")]
    [Arguments("line-series-format")]
    [Arguments("bar-axis-scale")]
    [Arguments("scatter-smooth")]
    public async Task Chart_part_xml_matches_the_golden_fixture(string name)
    {
        var actual = ChartGoldenCorpus.CaptureChartPartXml(ws => Fixtures(name, ws));
        var directory = ChartGoldenCorpus.GoldenDirectory();
        var path = Path.Combine(directory, name + ".xml");

        if (!File.Exists(path))
        {
            Directory.CreateDirectory(directory);
            File.WriteAllText(path, actual);
        }

        await Assert.That(actual).IsEqualTo(File.ReadAllText(path));
    }

    /// <summary>
    /// One builder per fixture name. Between them they reach the legend, the chart title, the axes
    /// and their scaling, the data labels at both series and group level, and the series shape
    /// properties, marker and smoothing — the five concepts spec 22 moves.
    /// </summary>
    private static void Fixtures(string name, IXLWorksheet ws)
    {
        switch (name)
        {
            case "bar-plain":
                AddChart(ws, XLChartType.ColumnClustered);
                break;

            case "line-legend-bottom":
            {
                // Visible has to be assigned too: a legend nobody switched on is not written, which
                // is what makes an unassigned legend round-trip untouched.
                var legend = AddChart(ws, XLChartType.Line).Legend;
                legend.Visible = true;
                legend.Position = XLLegendPosition.Bottom;
                legend.Overlay = true;
                break;
            }

            case "bar-titled":
            {
                var chart = AddChart(ws, XLChartType.ColumnClustered);
                chart.Title = "Quarterly";
                chart.ValueAxis.Title = "Units";
                break;
            }

            case "line-datalabels":
            {
                var chart = AddChart(ws, XLChartType.Line);
                chart.DataLabels.ShowValue = true;
                chart.DataLabels.Position = XLDataLabelPosition.Above;
                chart.Series.First().DataLabels.ShowCategoryName = true;
                break;
            }

            case "bar-secondary-axis":
            {
                var chart = AddChart(ws, XLChartType.ColumnClustered);
                chart.Series.Add("Price", "Data!$C$1:$C$2", Categories).UseSecondaryAxis = true;
                chart.SecondaryValueAxis.MajorGridlines = true;
                chart.ValueAxis.MajorGridlines = true;
                break;
            }

            case "line-series-format":
            {
                var chart = AddChart(ws, XLChartType.LineWithMarkers);
                var series = chart.Series.First();
                series.FillColor = XLColor.Red;
                series.LineColor = XLColor.FromTheme(XLThemeColor.Accent2);
                series.LineWidthPt = 2.25;
                series.MarkerStyle = XLMarkerStyle.Diamond;
                series.MarkerSize = 7;
                series.MarkerFillColor = XLColor.Blue;
                series.Smooth = true;
                break;
            }

            case "bar-axis-scale":
            {
                var chart = AddChart(ws, XLChartType.ColumnClustered);
                chart.ValueAxis.Min = 0;
                chart.ValueAxis.Max = 500;
                chart.ValueAxis.MajorUnit = 100;
                chart.ValueAxis.MinorUnit = 25;
                chart.ValueAxis.NumberFormat = "#,##0";
                chart.ValueAxis.Orientation = XLAxisOrientation.MaxMin;
                chart.CategoryAxis.Visible = false;
                break;
            }

            case "scatter-smooth":
            {
                var chart = AddChart(ws, XLChartType.XYScatterSmoothLinesWithMarkers);
                chart.ValueAxis.LogScale = true;
                chart.ValueAxis.LogBase = 2;
                break;
            }

            default:
                throw new ArgumentOutOfRangeException(nameof(name), name, "Unknown chart fixture.");
        }
    }

    private static IXLChart AddChart(IXLWorksheet ws, XLChartType chartType)
    {
        var chart = ws.Charts.Add(chartType);
        chart.Series.Add("Units", Values, Categories);
        chart.Position.SetColumn(5).SetRow(1);
        chart.SecondPosition.SetColumn(12).SetRow(15);
        return chart;
    }
}
