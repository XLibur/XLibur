using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;
using XLibur.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace XLibur.Tests.Excel.Charts;

/// <summary>
/// Series formatting: fill, outline, markers, smoothing and secondary axis binding.
/// </summary>
[TestFixture]
public class ChartSeriesFormattingTests
{
    // ── Helpers ─────────────────────────────────────────────────────────

    private static IXLWorksheet AddDataSheet(XLWorkbook wb, string name = "Data")
    {
        var ws = wb.AddWorksheet(name);
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

    /// <summary>
    /// Saves with the OpenXML validator switched on, so a schema-invalid child order fails the test
    /// rather than only showing up as a repair prompt in Excel.
    /// </summary>
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
        // Detach from the package so the caller can keep using it after the document is closed.
        return (C.ChartSpace)chartPart.ChartSpace!.CloneNode(true);
    }

    // ── Writing and reading back ────────────────────────────────────────

    [Test]
    public void SeriesFillAndLineRoundTrip()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = AddChart(ws, XLChartType.ColumnClustered);
            var series = chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            series.FillColor = XLColor.FromArgb(0xFF, 0x00, 0x00);
            series.LineColor = XLColor.FromArgb(0x00, 0x33, 0x66);
            series.LineWidthPt = 2.25;

            using var saved = SaveValidated(wb);
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var series = wb.Worksheet("Data").Charts.First().Series.First();
            Assert.That(series.FillColor, Is.EqualTo(XLColor.FromArgb(0xFF, 0x00, 0x00)));
            Assert.That(series.LineColor, Is.EqualTo(XLColor.FromArgb(0x00, 0x33, 0x66)));
            Assert.That(series.LineWidthPt, Is.EqualTo(2.25));
        }
    }

    [Test]
    public void SeriesFillIsWrittenAsSolidFillInShapeProperties()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2").FillColor = XLColor.FromArgb(0x12, 0x34, 0x56);

        using var ms = SaveValidated(wb);
        var chartSpace = ChartSpaceOf(ms);

        var seriesElement = chartSpace.Descendants<C.BarChartSeries>().Single();
        var shapeProperties = seriesElement.Elements<C.ChartShapeProperties>().Single();
        var rgb = shapeProperties.Elements<A.SolidFill>().Single().Elements<A.RgbColorModelHex>().Single();
        Assert.That(rgb.Val!.Value, Is.EqualTo("123456"));

        // c:spPr has to come before c:cat and c:val.
        var children = seriesElement.ChildElements.Select(e => e.LocalName).ToList();
        Assert.That(children.IndexOf("spPr"), Is.LessThan(children.IndexOf("cat")));
    }

    [Test]
    public void NoFormattingWritesNoShapeProperties()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");

        using var ms = SaveValidated(wb);
        var chartSpace = ChartSpaceOf(ms);

        var seriesElement = chartSpace.Descendants<C.BarChartSeries>().Single();
        Assert.That(seriesElement.Elements<C.ChartShapeProperties>(), Is.Empty,
            "An unformatted series must not pin down a colour; Excel picks the theme colour.");
    }

    [Test]
    public void ThemeFillColorRoundTrips()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = AddChart(ws, XLChartType.ColumnClustered);
            chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2").FillColor =
                XLColor.FromTheme(XLThemeColor.Accent3);

            using var saved = SaveValidated(wb);
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var series = wb.Worksheet("Data").Charts.First().Series.First();
            Assert.That(series.FillColor!.ColorType, Is.EqualTo(XLColorType.Theme));
            Assert.That(series.FillColor.ThemeColor, Is.EqualTo(XLThemeColor.Accent3));
        }
    }

    [Test]
    public void MarkerRoundTrips()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = AddChart(ws, XLChartType.Line);
            var series = chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            series.MarkerStyle = XLMarkerStyle.Diamond;
            series.MarkerSize = 9;
            series.MarkerFillColor = XLColor.FromArgb(0x00, 0xB0, 0x50);

            using var saved = SaveValidated(wb);
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var series = wb.Worksheet("Data").Charts.First().Series.First();
            Assert.That(series.MarkerStyle, Is.EqualTo(XLMarkerStyle.Diamond));
            Assert.That(series.MarkerSize, Is.EqualTo(9));
            Assert.That(series.MarkerFillColor, Is.EqualTo(XLColor.FromArgb(0x00, 0xB0, 0x50)));
        }
    }

    [Test]
    public void MarkerStyleNoneKeepsChartTypeLine()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = AddChart(ws, XLChartType.Line);
            chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2").MarkerStyle = XLMarkerStyle.None;

            using var saved = SaveValidated(wb);
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.Line));
            Assert.That(chart.Series.First().MarkerStyle, Is.EqualTo(XLMarkerStyle.None));
        }
    }

    [Test]
    public void LineWithMarkersStillWritesAutoMarker()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.LineWithMarkers);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");

        using var ms = SaveValidated(wb);
        var chartSpace = ChartSpaceOf(ms);

        var seriesElement = chartSpace.Descendants<C.LineChartSeries>().Single();
        var marker = seriesElement.Elements<C.Marker>().Single();
        Assert.That(marker.Elements<C.Symbol>().Single().Val!.Value, Is.EqualTo(C.MarkerStyleValues.Auto));

        // c:marker has to sit between c:tx and c:cat.
        var children = seriesElement.ChildElements.Select(e => e.LocalName).ToList();
        Assert.That(children.IndexOf("marker"), Is.LessThan(children.IndexOf("cat")));
    }

    [Test]
    public void SmoothRoundTrips()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = AddChart(ws, XLChartType.Line);
            chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2").Smooth = true;

            using var saved = SaveValidated(wb);
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            Assert.That(wb.Worksheet("Data").Charts.First().Series.First().Smooth, Is.True);
        }
    }

    [Test]
    public void SmoothIsNotWrittenWhenLeftAlone()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.Line);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");

        using var ms = SaveValidated(wb);
        var chartSpace = ChartSpaceOf(ms);

        Assert.That(chartSpace.Descendants<C.Smooth>(), Is.Empty);
    }

    [Test]
    public void SmoothScatterTypeIsWrittenAsSmoothed()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.XYScatterSmoothLinesNoMarkers);
        chart.Series.Add("Points", "Data!$B$1:$B$2", "Data!$A$1:$A$2");

        using var ms = SaveValidated(wb);
        var chartSpace = ChartSpaceOf(ms);

        Assert.That(chartSpace.Descendants<C.Smooth>().Single().Val!.Value, Is.True);
    }

    [Test]
    public void ExplicitFalseOverridesTheSmoothChartTypeDefault()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.XYScatterSmoothLinesNoMarkers);
        chart.Series.Add("Points", "Data!$B$1:$B$2", "Data!$A$1:$A$2").Smooth = false;

        using var ms = SaveValidated(wb);
        var chartSpace = ChartSpaceOf(ms);

        Assert.That(chartSpace.Descendants<C.Smooth>().Single().Val!.Value, Is.False);
    }

    [Test]
    public void StockSeriesAreSmoothedToo()
    {
        // A stock chart is built from CT_LineSer, which takes c:smooth. The writer used to leave it
        // out, so Smooth was honoured on a stock chart read from a file but not on a new one.
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = AddChart(ws, XLChartType.StockHighLowClose);
            foreach (var name in new[] { "High", "Low", "Close" })
                chart.Series.Add(name, "Data!$B$1:$B$2", "Data!$A$1:$A$2").Smooth = true;

            using var saved = SaveValidated(wb);
            var chartSpace = ChartSpaceOf(saved);
            Assert.That(chartSpace.Descendants<C.Smooth>().Select(s => s.Val!.Value),
                Is.EqualTo(new[] { true, true, true }));

            saved.Position = 0;
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            Assert.That(wb.Worksheet("Data").Charts.First().Series.Select(s => s.Smooth),
                Is.EqualTo(new[] { true, true, true }));
        }
    }

    [Test]
    public void FormattingSurvivesEveryStandardChartFamily()
    {
        // The formatting is appended by each chart family's own builder, so every family gets a
        // schema check.
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

            // A stock chart is only valid with three or four series (high/low/close).
            var seriesCount = type == XLChartType.StockHighLowClose ? 3 : 1;
            for (var i = 0; i < seriesCount; i++)
            {
                var series = chart.Series.Add($"S{i}", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
                series.FillColor = XLColor.Red;
                series.LineColor = XLColor.Blue;
                series.LineWidthPt = 1.5;
                series.MarkerStyle = XLMarkerStyle.Square;
                series.MarkerSize = 7;
                series.Smooth = true;
            }

            Assert.DoesNotThrow(() =>
            {
                using var ms = SaveValidated(wb);
            }, $"{type} produced invalid chart XML.");
        }
    }

    // ── Secondary axis ──────────────────────────────────────────────────

    [Test]
    public void SecondaryAxisSeriesGetsItsOwnPlotGroupAndAxes()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.ColumnClustered);
        chart.Series.Add("Units", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
        chart.Series.Add("Price", "Data!$C$1:$C$2", "Data!$A$1:$A$2").UseSecondaryAxis = true;

        using var ms = SaveValidated(wb);
        var chartSpace = ChartSpaceOf(ms);
        var plotArea = chartSpace.Descendants<C.PlotArea>().Single();

        var barCharts = plotArea.Elements<C.BarChart>().ToList();
        Assert.That(barCharts, Has.Count.EqualTo(2), "The secondary series needs its own bar chart group.");

        var primaryAxisIds = barCharts[0].Elements<C.AxisId>().Select(a => a.Val!.Value).ToList();
        var secondaryAxisIds = barCharts[1].Elements<C.AxisId>().Select(a => a.Val!.Value).ToList();
        Assert.That(secondaryAxisIds, Is.Not.EquivalentTo(primaryAxisIds));

        Assert.That(plotArea.Elements<C.ValueAxis>().Count(), Is.EqualTo(2));
        Assert.That(plotArea.Elements<C.CategoryAxis>().Count(), Is.EqualTo(2));

        // The extra category axis is hidden and the extra value axis sits on the right.
        var hiddenCategoryAxis = plotArea.Elements<C.CategoryAxis>()
            .Single(a => a.Elements<C.Delete>().Single().Val!.Value);
        Assert.That(hiddenCategoryAxis.Elements<C.AxisId>().Single().Val!.Value,
            Is.EqualTo(secondaryAxisIds[0]));

        var rightValueAxis = plotArea.Elements<C.ValueAxis>()
            .Single(a => a.Elements<C.AxisPosition>().Single().Val!.Value == C.AxisPositionValues.Right);
        Assert.That(rightValueAxis.Elements<C.Crosses>().Single().Val!.Value,
            Is.EqualTo(C.CrossesValues.Maximum));
    }

    [Test]
    public void SecondaryAxisRoundTrips()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = AddChart(ws, XLChartType.ColumnClustered);
            chart.Series.Add("Units", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Series.Add("Price", "Data!$C$1:$C$2", "Data!$A$1:$A$2").UseSecondaryAxis = true;

            using var saved = SaveValidated(wb);
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var series = wb.Worksheet("Data").Charts.First().Series.ToList();
            Assert.That(series, Has.Count.EqualTo(2));
            Assert.That(series[0].Name, Is.EqualTo("Units"));
            Assert.That(series[0].UseSecondaryAxis, Is.False);
            Assert.That(series[1].Name, Is.EqualTo("Price"));
            Assert.That(series[1].UseSecondaryAxis, Is.True);
        }
    }

    [Test]
    public void ComboChartCanPutItsSecondaryTypeOnTheSecondaryAxis()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = AddChart(ws, XLChartType.ColumnClustered);
            chart.Series.Add("Units", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.SecondaryChartType = XLChartType.Line;
            chart.SecondarySeries.Add("Price", "Data!$C$1:$C$2", "Data!$A$1:$A$2").UseSecondaryAxis = true;

            using var saved = SaveValidated(wb);
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.ColumnClustered));
            Assert.That(chart.SecondaryChartType, Is.EqualTo(XLChartType.Line));
            Assert.That(chart.Series.Single().UseSecondaryAxis, Is.False);
            Assert.That(chart.SecondarySeries.Single().UseSecondaryAxis, Is.True);
        }
    }

    [Test]
    public void SecondaryAxisIsIgnoredForChartTypesWithoutOne()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.Pie);
        chart.Series.Add("Share", "Data!$B$1:$B$2", "Data!$A$1:$A$2").UseSecondaryAxis = true;

        using var ms = SaveValidated(wb);
        var chartSpace = ChartSpaceOf(ms);

        Assert.That(chartSpace.Descendants<C.PieChart>().Count(), Is.EqualTo(1));
        Assert.That(chartSpace.Descendants<C.ValueAxis>(), Is.Empty);
    }

    // ── Validation of the property setters ──────────────────────────────

    [Test]
    public void MarkerSizeOutsideExcelsRangeIsRejected()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var series = AddChart(ws, XLChartType.Line).Series.Add("S", "Data!$B$1:$B$2");

        Assert.Throws<ArgumentOutOfRangeException>(() => series.MarkerSize = 1);
        Assert.Throws<ArgumentOutOfRangeException>(() => series.MarkerSize = 73);
        Assert.DoesNotThrow(() => series.MarkerSize = null);
    }

    [Test]
    public void LineWidthOutsideExcelsRangeIsRejected()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var series = AddChart(ws, XLChartType.Line).Series.Add("S", "Data!$B$1:$B$2");

        Assert.Throws<ArgumentOutOfRangeException>(() => series.LineWidthPt = -1);
        Assert.Throws<ArgumentOutOfRangeException>(() => series.LineWidthPt = 1585);
        Assert.DoesNotThrow(() => series.LineWidthPt = null);
    }

    [Test]
    public void TheFormattedChartExampleProducesAValidWorkbook()
    {
        // Left on disk on purpose: this is the file to open in Excel when checking that the
        // formatting renders the way it is meant to.
        var path = Path.Combine(TestContext.CurrentContext.WorkDirectory, "FormattedChartExamples.xlsx");
        new XLibur.Examples.Charts.FormattedChartExamples().Create(path);
        TestContext.Out.WriteLine($"Formatted chart example: {path}");

        // Reloading through XLibur and saving with validation on checks both that the example reads
        // back and that what it wrote is schema-valid.
        using var wb = new XLWorkbook(path);
        using var ms = SaveValidated(wb);
        Assert.That(wb.Worksheets.Count, Is.EqualTo(7));
    }

    [Test]
    public void SecondaryAxisCannotBeChangedOnALoadedChart()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = AddChart(ws, XLChartType.ColumnClustered);
            chart.Series.Add("Units", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            using var saved = SaveValidated(wb);
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var series = wb.Worksheet("Data").Charts.First().Series.First();
            var ex = Assert.Throws<NotSupportedException>(() => series.UseSecondaryAxis = true);
            Assert.That(ex!.Message, Does.Contain("loaded from a file"));

            // Assigning the value it already has is not a change, so it is allowed.
            Assert.DoesNotThrow(() => series.UseSecondaryAxis = false);
        }
    }
}
