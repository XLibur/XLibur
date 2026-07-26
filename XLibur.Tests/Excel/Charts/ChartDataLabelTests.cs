using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;
using XLibur.Excel;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace XLibur.Tests.Excel.Charts;

/// <summary>
/// Data labels, per series and chart-wide.
/// </summary>
[TestFixture]
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

    private static C.ChartSpace ChartSpaceOf(Stream stream)
    {
        stream.Position = 0;
        using var doc = SpreadsheetDocument.Open(stream, false);
        var chartPart = doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!.ChartParts.First();
        return (C.ChartSpace)chartPart.ChartSpace!.CloneNode(true);
    }

    // ── Writing and reading back ────────────────────────────────────────

    [Test]
    public void SeriesDataLabelsRoundTrip()
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
            Assert.That(labels.ShowValue, Is.True);
            Assert.That(labels.ShowCategoryName, Is.True);
            Assert.That(labels.ShowSeriesName, Is.False);
            Assert.That(labels.ShowPercentage, Is.False);
            Assert.That(labels.NumberFormat, Is.EqualTo("#,##0"));
            Assert.That(labels.Position, Is.EqualTo(XLDataLabelPosition.OutsideEnd));
        }
    }

    [Test]
    public void ChartWideDataLabelsRoundTrip()
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
            Assert.That(chart.DataLabels.ShowValue, Is.True);
            Assert.That(chart.DataLabels.Position, Is.EqualTo(XLDataLabelPosition.InsideEnd));

            // Nothing was set per series, so the series labels stay at their defaults.
            Assert.That(chart.Series.First().DataLabels.ShowValue, Is.False);
        }
    }

    [Test]
    public void ChartWideLabelsAreWrittenOnTheChartGroupAndSeriesLabelsOnTheSeries()
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
        Assert.That(groupLabels.Elements<C.ShowValue>().Single().Val!.Value, Is.True);

        var seriesElements = barChart.Elements<C.BarChartSeries>().ToList();
        Assert.That(seriesElements[0].Elements<C.DataLabels>(), Is.Empty);
        Assert.That(seriesElements[1].Elements<C.DataLabels>().Single()
            .Elements<C.ShowSeriesName>().Single().Val!.Value, Is.True);

        // c:dLbls sits after the last c:ser and before the axis ids.
        var children = barChart.ChildElements.Select(e => e.LocalName).ToList();
        Assert.That(children.LastIndexOf("ser"), Is.LessThan(children.IndexOf("dLbls")));
        Assert.That(children.IndexOf("dLbls"), Is.LessThan(children.IndexOf("axId")));
    }

    [Test]
    public void NoDataLabelsAreWrittenWhenNothingIsAsked()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");

        using var ms = SaveValidated(wb);
        Assert.That(ChartSpaceOf(ms).Descendants<C.DataLabels>(), Is.Empty);
    }

    [Test]
    public void EveryShowFlagIsWrittenSoTheResultDoesNotDependOnTheChartStyle()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2").DataLabels.ShowValue = true;

        using var ms = SaveValidated(wb);
        var labels = ChartSpaceOf(ms).Descendants<C.DataLabels>().Single();

        Assert.That(labels.Elements<C.ShowLegendKey>().Single().Val!.Value, Is.False);
        Assert.That(labels.Elements<C.ShowValue>().Single().Val!.Value, Is.True);
        Assert.That(labels.Elements<C.ShowCategoryName>().Single().Val!.Value, Is.False);
        Assert.That(labels.Elements<C.ShowSeriesName>().Single().Val!.Value, Is.False);
        Assert.That(labels.Elements<C.ShowPercent>().Single().Val!.Value, Is.False);
        Assert.That(labels.Elements<C.ShowBubbleSize>().Single().Val!.Value, Is.False);
    }

    [Test]
    public void PercentageLabelsOnAPieChartRoundTrip()
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
            Assert.That(labels.ShowPercentage, Is.True);
            Assert.That(labels.Position, Is.EqualTo(XLDataLabelPosition.BestFit));
        }
    }

    [Test]
    public void LabelsSurviveEveryChartFamilyThatSupportsThem()
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

            Assert.DoesNotThrow(() =>
            {
                using var ms = SaveValidated(wb);
            }, $"{type} produced invalid chart XML.");
        }
    }

    [Test]
    public void SurfaceChartsDoNotGetDataLabels()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.Surface);
        chart.Series.Add("S1", "Data!$B$1:$B$2", "Data!$A$1:$A$2").DataLabels.ShowValue = true;
        chart.DataLabels.ShowValue = true;

        using var ms = SaveValidated(wb);
        Assert.That(ChartSpaceOf(ms).Descendants<C.DataLabels>(), Is.Empty,
            "Neither CT_SurfaceChart nor CT_SurfaceSer has a dLbls child.");
    }

    // ── Position validation ─────────────────────────────────────────────

    [Test]
    public void OutsideEndIsRejectedOnAStackedColumnChart()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.ColumnStacked);
        var labels = chart.Series.Add("S", "Data!$B$1:$B$2").DataLabels;

        var ex = Assert.Throws<ArgumentException>(() => labels.Position = XLDataLabelPosition.OutsideEnd);
        Assert.That(ex!.Message, Does.Contain("ColumnStacked"));
        Assert.That(ex.Message, Does.Contain("InsideBase"), "The message lists what Excel does offer.");

        Assert.DoesNotThrow(() => labels.Position = XLDataLabelPosition.InsideEnd);
    }

    [Test]
    public void MarkerPositionsAreRejectedOnAColumnChart()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var labels = AddChart(ws, XLChartType.ColumnClustered)
            .Series.Add("S", "Data!$B$1:$B$2").DataLabels;

        Assert.Throws<ArgumentException>(() => labels.Position = XLDataLabelPosition.Above);
        Assert.Throws<ArgumentException>(() => labels.Position = XLDataLabelPosition.BestFit);
    }

    [Test]
    public void BarPositionsAreRejectedOnALineChart()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var labels = AddChart(ws, XLChartType.Line).Series.Add("S", "Data!$B$1:$B$2").DataLabels;

        Assert.Throws<ArgumentException>(() => labels.Position = XLDataLabelPosition.OutsideEnd);
        Assert.DoesNotThrow(() => labels.Position = XLDataLabelPosition.Above);
    }

    [Test]
    public void ChartTypesWithoutPositionsAcceptOnlyAuto()
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

            var ex = Assert.Throws<ArgumentException>(
                () => labels.Position = XLDataLabelPosition.Center, $"{type} should refuse a position.");
            Assert.That(ex!.Message, Does.Contain("only Auto"));
        }
    }

    [Test]
    public void AComboChartsSecondarySeriesIsValidatedAgainstTheSecondaryChartType()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.ColumnClustered);
        chart.Series.Add("Units", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
        chart.SecondaryChartType = XLChartType.LineWithMarkers;
        var line = chart.SecondarySeries.Add("Price", "Data!$C$1:$C$2", "Data!$A$1:$A$2");

        // Above is a line position, which the primary column type would refuse.
        Assert.DoesNotThrow(() => line.DataLabels.Position = XLDataLabelPosition.Above);
        Assert.Throws<ArgumentException>(() => line.DataLabels.Position = XLDataLabelPosition.OutsideEnd);

        line.DataLabels.ShowValue = true;
        Assert.DoesNotThrow(() =>
        {
            using var ms = SaveValidated(wb);
        });
    }

    [Test]
    public void APositionThatBecomesInvalidWhenTheChartTypeChangesIsDropped()
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
        Assert.That(labels.Elements<C.DataLabelPosition>(), Is.Empty);
        Assert.That(labels.Elements<C.ShowValue>().Single().Val!.Value, Is.True);
    }
}
