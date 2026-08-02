using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Linq;
using XLibur.Excel;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Charts;

/// <summary>
/// Data labels, per series and chart-wide.
/// </summary>
public class ChartDataLabelTests
{
    private static IXLWorksheet AddDataSheet(XLWorkbook wb)
    {
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = "Q1";
        ws.Cell("A2").Value = "Q2";
        ws.Cell("B1").Value = 100;
        ws.Cell("B2").Value = 200;
        ws.Cell("C1").Value = 5;
        ws.Cell("C2").Value = 8;
        return ws;
    }

    private static IXLChart AddChart(IXLWorksheet ws, XLChartType type)
    {
        var chart = ws.Charts.Add(type);
        chart.Position.SetColumn(5).SetRow(1);
        chart.SecondPosition.SetColumn(12).SetRow(15);
        return chart;
    }

    private static MemoryStream SaveValidated(XLWorkbook wb)
    {
        var ms = new MemoryStream();
        wb.SaveAs(ms, validate: true);
        ms.Position = 0;
        return ms;
    }

    private static C.ChartSpace ChartSpaceOf(MemoryStream stream)
    {
        stream.Position = 0;
        using var doc = SpreadsheetDocument.Open(stream, false);
        var chartPart = doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!.ChartParts.First();
        return (C.ChartSpace)chartPart.ChartSpace!.CloneNode(true);
    }

    // ── Writing and reading back ────────────────────────────────────────

    [Test]
    public async Task SeriesDataLabelsRoundTrip()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = AddChart(ws, XLChartType.ColumnClustered);
            var series = chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            series.DataLabels.ShowValue = true;
            series.DataLabels.ShowCategoryName = true;
            series.DataLabels.NumberFormat = "#,##0";
            series.DataLabels.Position = XLDataLabelPosition.OutsideEnd;

            using var saved = SaveValidated(wb);
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var labels = wb.Worksheet("Data").Charts.First().Series.First().DataLabels;
            await Assert.That(labels.ShowValue).IsTrue();
            await Assert.That(labels.ShowCategoryName).IsTrue();
            await Assert.That(labels.ShowSeriesName).IsFalse();
            await Assert.That(labels.ShowPercentage).IsFalse();
            await Assert.That(labels.NumberFormat).IsEqualTo("#,##0");
            await Assert.That(labels.Position).IsEqualTo(XLDataLabelPosition.OutsideEnd);
        }
    }

    [Test]
    public async Task ChartWideDataLabelsRoundTrip()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = AddChart(ws, XLChartType.ColumnClustered);
            chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Series.Add("Cost", "Data!$C$1:$C$2", "Data!$A$1:$A$2");
            chart.DataLabels.ShowValue = true;
            chart.DataLabels.Position = XLDataLabelPosition.InsideEnd;

            using var saved = SaveValidated(wb);
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            await Assert.That(chart.DataLabels.ShowValue).IsTrue();
            await Assert.That(chart.DataLabels.Position).IsEqualTo(XLDataLabelPosition.InsideEnd);

            // Nothing was set per series, so the series labels stay at their defaults.
            await Assert.That(chart.Series.First().DataLabels.ShowValue).IsFalse();
        }
    }

    [Test]
    public async Task ChartWideLabelsAreWrittenOnTheChartGroupAndSeriesLabelsOnTheSeries()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
        chart.Series.Add("Cost", "Data!$C$1:$C$2", "Data!$A$1:$A$2").DataLabels.ShowSeriesName = true;
        chart.DataLabels.ShowValue = true;

        using var ms = SaveValidated(wb);
        var chartSpace = ChartSpaceOf(ms);
        var barChart = chartSpace.Descendants<C.BarChart>().Single();

        var groupLabels = barChart.Elements<C.DataLabels>().Single();
        await Assert.That(groupLabels.Elements<C.ShowValue>().Single().Val!.Value).IsTrue();

        var seriesElements = barChart.Elements<C.BarChartSeries>().ToList();
        await Assert.That(seriesElements[0].Elements<C.DataLabels>()).IsEmpty();
        await Assert.That(seriesElements[1].Elements<C.DataLabels>().Single()
            .Elements<C.ShowSeriesName>().Single().Val!.Value).IsTrue();

        // c:dLbls sits after the last c:ser and before the axis ids.
        var children = barChart.ChildElements.Select(e => e.LocalName).ToList();
        await Assert.That(children.LastIndexOf("ser")).IsLessThan(children.IndexOf("dLbls"));
        await Assert.That(children.IndexOf("dLbls")).IsLessThan(children.IndexOf("axId"));
    }

    [Test]
    public async Task NoDataLabelsAreWrittenWhenNothingIsAsked()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");

        using var ms = SaveValidated(wb);
        await Assert.That(ChartSpaceOf(ms).Descendants<C.DataLabels>()).IsEmpty();
    }

    [Test]
    public async Task EveryShowFlagIsWrittenSoTheResultDoesNotDependOnTheChartStyle()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2").DataLabels.ShowValue = true;

        using var ms = SaveValidated(wb);
        var labels = ChartSpaceOf(ms).Descendants<C.DataLabels>().Single();

        await Assert.That(labels.Elements<C.ShowLegendKey>().Single().Val!.Value).IsFalse();
        await Assert.That(labels.Elements<C.ShowValue>().Single().Val!.Value).IsTrue();
        await Assert.That(labels.Elements<C.ShowCategoryName>().Single().Val!.Value).IsFalse();
        await Assert.That(labels.Elements<C.ShowSeriesName>().Single().Val!.Value).IsFalse();
        await Assert.That(labels.Elements<C.ShowPercent>().Single().Val!.Value).IsFalse();
        await Assert.That(labels.Elements<C.ShowBubbleSize>().Single().Val!.Value).IsFalse();
    }

    [Test]
    public async Task PercentageLabelsOnAPieChartRoundTrip()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = AddChart(ws, XLChartType.Pie);
            var series = chart.Series.Add("Share", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            series.DataLabels.ShowPercentage = true;
            series.DataLabels.Position = XLDataLabelPosition.BestFit;

            using var saved = SaveValidated(wb);
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var labels = wb.Worksheet("Data").Charts.First().Series.First().DataLabels;
            await Assert.That(labels.ShowPercentage).IsTrue();
            await Assert.That(labels.Position).IsEqualTo(XLDataLabelPosition.BestFit);
        }
    }

    [Test]
    public async Task LabelsSurviveEveryChartFamilyThatSupportsThem()
    {
        XLChartType[] types =
        [
            XLChartType.ColumnClustered, XLChartType.BarStacked, XLChartType.ColumnClustered3D,
            XLChartType.Line, XLChartType.LineWithMarkers, XLChartType.Area, XLChartType.Radar,
            XLChartType.Pie, XLChartType.Doughnut, XLChartType.XYScatterMarkers, XLChartType.Bubble,
            XLChartType.StockHighLowClose, XLChartType.Surface, XLChartType.ConeClustered
        ];

        foreach (var type in types)
        {
            using var wb = new XLWorkbook();
            var ws = AddDataSheet(wb);
            var chart = AddChart(ws, type);

            var seriesCount = type == XLChartType.StockHighLowClose ? 3 : 1;
            for (var i = 0; i < seriesCount; i++)
            {
                var series = chart.Series.Add($"S{i}", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
                series.DataLabels.ShowValue = true;
                series.DataLabels.NumberFormat = "0.00";
            }

            chart.DataLabels.ShowCategoryName = true;

            await Assert.That(() =>
            {
                using var ms = SaveValidated(wb);
            }).ThrowsNothing().Because($"{type} produced invalid chart XML.");
        }
    }

    [Test]
    public async Task SurfaceChartsDoNotGetDataLabels()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.Surface);
        chart.Series.Add("S1", "Data!$B$1:$B$2", "Data!$A$1:$A$2").DataLabels.ShowValue = true;
        chart.DataLabels.ShowValue = true;

        using var ms = SaveValidated(wb);
        await Assert.That(ChartSpaceOf(ms).Descendants<C.DataLabels>()).IsEmpty().Because("Neither CT_SurfaceChart nor CT_SurfaceSer has a dLbls child.");
    }

    // ── Position validation ─────────────────────────────────────────────

    [Test]
    public async Task OutsideEndIsRejectedOnAStackedColumnChart()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.ColumnStacked);
        var labels = chart.Series.Add("S", "Data!$B$1:$B$2").DataLabels;

        var ex = await Assert.That(() => labels.Position = XLDataLabelPosition.OutsideEnd).Throws<ArgumentException>();
        await Assert.That(ex!.Message).Contains("ColumnStacked");
        await Assert.That(ex!.Message).Contains("InsideBase").Because("The message lists what Excel does offer.");

        await Assert.That(() => labels.Position = XLDataLabelPosition.InsideEnd).ThrowsNothing();
    }

    [Test]
    public async Task MarkerPositionsAreRejectedOnAColumnChart()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var labels = AddChart(ws, XLChartType.ColumnClustered)
            .Series.Add("S", "Data!$B$1:$B$2").DataLabels;

        await Assert.That(() => labels.Position = XLDataLabelPosition.Above).Throws<ArgumentException>();
        await Assert.That(() => labels.Position = XLDataLabelPosition.BestFit).Throws<ArgumentException>();
    }

    [Test]
    public async Task BarPositionsAreRejectedOnALineChart()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var labels = AddChart(ws, XLChartType.Line).Series.Add("S", "Data!$B$1:$B$2").DataLabels;

        await Assert.That(() => labels.Position = XLDataLabelPosition.OutsideEnd).Throws<ArgumentException>();
        await Assert.That(() => labels.Position = XLDataLabelPosition.Above).ThrowsNothing();
    }

    [Test]
    public async Task ChartTypesWithoutPositionsAcceptOnlyAuto()
    {
        XLChartType[] types =
        [
            XLChartType.Area, XLChartType.Doughnut, XLChartType.Bubble,
            XLChartType.StockHighLowClose, XLChartType.ColumnClustered3D, XLChartType.Surface
        ];

        foreach (var type in types)
        {
            using var wb = new XLWorkbook();
            var ws = AddDataSheet(wb);
            var labels = AddChart(ws, type).Series.Add("S", "Data!$B$1:$B$2").DataLabels;

            var ex = await Assert.That(() => labels.Position = XLDataLabelPosition.Center).Throws<ArgumentException>().Because($"{type} should refuse a position.");
            await Assert.That(ex!.Message).Contains("only Auto");
        }
    }

    [Test]
    public async Task AComboChartsSecondarySeriesIsValidatedAgainstTheSecondaryChartType()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.ColumnClustered);
        chart.Series.Add("Units", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
        chart.SecondaryChartType = XLChartType.LineWithMarkers;
        var line = chart.SecondarySeries.Add("Price", "Data!$C$1:$C$2", "Data!$A$1:$A$2");

        // Above is a line position, which the primary column type would refuse.
        await Assert.That(() => line.DataLabels.Position = XLDataLabelPosition.Above).ThrowsNothing();
        await Assert.That(() => line.DataLabels.Position = XLDataLabelPosition.OutsideEnd).Throws<ArgumentException>();

        line.DataLabels.ShowValue = true;
        await Assert.That(() =>
        {
            using var ms = SaveValidated(wb);
        }).ThrowsNothing();
    }

    [Test]
    public async Task APositionThatBecomesInvalidWhenTheChartTypeChangesIsDropped()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.ColumnClustered);
        var series = chart.Series.Add("S", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
        series.DataLabels.ShowValue = true;
        series.DataLabels.Position = XLDataLabelPosition.OutsideEnd;

        // Area charts offer no explicit position. Writing the one set earlier would make Excel
        // refuse the file, so it is left out.
        chart.ChartType = XLChartType.Area;

        using var ms = SaveValidated(wb);
        var labels = ChartSpaceOf(ms).Descendants<C.DataLabels>().Single();
        await Assert.That(labels.Elements<C.DataLabelPosition>()).IsEmpty();
        await Assert.That(labels.Elements<C.ShowValue>().Single().Val!.Value).IsTrue();
    }
}
