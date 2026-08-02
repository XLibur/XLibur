using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Linq;
using XLibur.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;

namespace XLibur.Tests.Excel.Charts;

/// <summary>
/// Series formatting: fill, outline, markers, smoothing and secondary axis binding.
/// </summary>
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

    private static C.ChartSpace ChartSpaceOf(MemoryStream stream)
    {
        stream.Position = 0;
        using var doc = SpreadsheetDocument.Open(stream, false);
        var chartPart = doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!.ChartParts.First();
        // Detach from the package so the caller can keep using it after the document is closed.
        return (C.ChartSpace)chartPart.ChartSpace!.CloneNode(true);
    }

    // ── Writing and reading back ────────────────────────────────────────

    [Test]
    public async Task SeriesFillAndLineRoundTrip()
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
            await Assert.That(series.FillColor).IsEqualTo(XLColor.FromArgb(0xFF, 0x00, 0x00));
            await Assert.That(series.LineColor).IsEqualTo(XLColor.FromArgb(0x00, 0x33, 0x66));
            await Assert.That(series.LineWidthPt).IsEqualTo(2.25);
        }
    }

    [Test]
    public async Task SeriesFillIsWrittenAsSolidFillInShapeProperties()
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
        await Assert.That(rgb.Val!.Value).IsEqualTo("123456");

        // c:spPr has to come before c:cat and c:val.
        var children = seriesElement.ChildElements.Select(e => e.LocalName).ToList();
        await Assert.That(children.IndexOf("spPr")).IsLessThan(children.IndexOf("cat"));
    }

    [Test]
    public async Task NoFormattingWritesNoShapeProperties()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");

        using var ms = SaveValidated(wb);
        var chartSpace = ChartSpaceOf(ms);

        var seriesElement = chartSpace.Descendants<C.BarChartSeries>().Single();
        await Assert.That(seriesElement.Elements<C.ChartShapeProperties>()).IsEmpty().Because("An unformatted series must not pin down a colour; Excel picks the theme colour.");
    }

    [Test]
    public async Task ThemeFillColorRoundTrips()
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
            await Assert.That(series.FillColor!.ColorType).IsEqualTo(XLColorType.Theme);
            await Assert.That(series.FillColor.ThemeColor).IsEqualTo(XLThemeColor.Accent3);
        }
    }

    [Test]
    public async Task MarkerRoundTrips()
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
            await Assert.That(series.MarkerStyle).IsEqualTo(XLMarkerStyle.Diamond);
            await Assert.That(series.MarkerSize).IsEqualTo(9);
            await Assert.That(series.MarkerFillColor).IsEqualTo(XLColor.FromArgb(0x00, 0xB0, 0x50));
        }
    }

    [Test]
    public async Task MarkerStyleNoneKeepsChartTypeLine()
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
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.Line);
            await Assert.That(chart.Series.First().MarkerStyle).IsEqualTo(XLMarkerStyle.None);
        }
    }

    [Test]
    public async Task LineWithMarkersStillWritesAutoMarker()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.LineWithMarkers);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");

        using var ms = SaveValidated(wb);
        var chartSpace = ChartSpaceOf(ms);

        var seriesElement = chartSpace.Descendants<C.LineChartSeries>().Single();
        var marker = seriesElement.Elements<C.Marker>().Single();
        await Assert.That(marker.Elements<C.Symbol>().Single().Val!.Value).IsEqualTo(C.MarkerStyleValues.Auto);

        // c:marker has to sit between c:tx and c:cat.
        var children = seriesElement.ChildElements.Select(e => e.LocalName).ToList();
        await Assert.That(children.IndexOf("marker")).IsLessThan(children.IndexOf("cat"));
    }

    [Test]
    public async Task SmoothRoundTrips()
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
            await Assert.That(wb.Worksheet("Data").Charts.First().Series.First().Smooth).IsTrue();
        }
    }

    [Test]
    public async Task SmoothIsNotWrittenWhenLeftAlone()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.Line);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");

        using var ms = SaveValidated(wb);
        var chartSpace = ChartSpaceOf(ms);

        await Assert.That(chartSpace.Descendants<C.Smooth>()).IsEmpty();
    }

    [Test]
    public async Task SmoothScatterTypeIsWrittenAsSmoothed()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.XYScatterSmoothLinesNoMarkers);
        chart.Series.Add("Points", "Data!$B$1:$B$2", "Data!$A$1:$A$2");

        using var ms = SaveValidated(wb);
        var chartSpace = ChartSpaceOf(ms);

        await Assert.That(chartSpace.Descendants<C.Smooth>().Single().Val!.Value).IsTrue();
    }

    [Test]
    public async Task ExplicitFalseOverridesTheSmoothChartTypeDefault()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.XYScatterSmoothLinesNoMarkers);
        chart.Series.Add("Points", "Data!$B$1:$B$2", "Data!$A$1:$A$2").Smooth = false;

        using var ms = SaveValidated(wb);
        var chartSpace = ChartSpaceOf(ms);

        await Assert.That(chartSpace.Descendants<C.Smooth>().Single().Val!.Value).IsFalse();
    }

    [Test]
    public async Task StockSeriesAreSmoothedToo()
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
            await Assert.That(chartSpace.Descendants<C.Smooth>().Select(s => s.Val!.Value)).IsEquivalentTo(new[] { true, true, true }, CollectionOrdering.Matching);

            saved.Position = 0;
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            await Assert.That(wb.Worksheet("Data").Charts.First().Series.Select(s => s.Smooth)).IsEquivalentTo(new[] { true, true, true }, CollectionOrdering.Matching);
        }
    }

    [Test]
    public async Task FormattingSurvivesEveryStandardChartFamily()
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

            await Assert.That(() =>
            {
                using var ms = SaveValidated(wb);
            }).ThrowsNothing().Because($"{type} produced invalid chart XML.");
        }
    }

    // ── Secondary axis ──────────────────────────────────────────────────

    [Test]
    public async Task SecondaryAxisSeriesGetsItsOwnPlotGroupAndAxes()
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
        await Assert.That(barCharts.Count).IsEqualTo(2).Because("The secondary series needs its own bar chart group.");

        var primaryAxisIds = barCharts[0].Elements<C.AxisId>().Select(a => a.Val!.Value).ToList();
        var secondaryAxisIds = barCharts[1].Elements<C.AxisId>().Select(a => a.Val!.Value).ToList();
        await Assert.That(secondaryAxisIds).IsNotEquivalentTo(primaryAxisIds);

        await Assert.That(plotArea.Elements<C.ValueAxis>().Count()).IsEqualTo(2);
        await Assert.That(plotArea.Elements<C.CategoryAxis>().Count()).IsEqualTo(2);

        // The extra category axis is hidden and the extra value axis sits on the right.
        var hiddenCategoryAxis = plotArea.Elements<C.CategoryAxis>()
            .Single(a => a.Elements<C.Delete>().Single().Val!.Value);
        await Assert.That(hiddenCategoryAxis.Elements<C.AxisId>().Single().Val!.Value).IsEqualTo(secondaryAxisIds[0]);

        var rightValueAxis = plotArea.Elements<C.ValueAxis>()
            .Single(a => a.Elements<C.AxisPosition>().Single().Val!.Value == C.AxisPositionValues.Right);
        await Assert.That(rightValueAxis.Elements<C.Crosses>().Single().Val!.Value).IsEqualTo(C.CrossesValues.Maximum);
    }

    [Test]
    public async Task SecondaryAxisRoundTrips()
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
            await Assert.That(series.Count).IsEqualTo(2);
            await Assert.That(series[0].Name).IsEqualTo("Units");
            await Assert.That(series[0].UseSecondaryAxis).IsFalse();
            await Assert.That(series[1].Name).IsEqualTo("Price");
            await Assert.That(series[1].UseSecondaryAxis).IsTrue();
        }
    }

    [Test]
    public async Task ComboChartCanPutItsSecondaryTypeOnTheSecondaryAxis()
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
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.ColumnClustered);
            await Assert.That(chart.SecondaryChartType).IsEqualTo(XLChartType.Line);
            await Assert.That(chart.Series.Single().UseSecondaryAxis).IsFalse();
            await Assert.That(chart.SecondarySeries.Single().UseSecondaryAxis).IsTrue();
        }
    }

    [Test]
    public async Task SecondaryAxisIsIgnoredForChartTypesWithoutOne()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.Pie);
        chart.Series.Add("Share", "Data!$B$1:$B$2", "Data!$A$1:$A$2").UseSecondaryAxis = true;

        using var ms = SaveValidated(wb);
        var chartSpace = ChartSpaceOf(ms);

        await Assert.That(chartSpace.Descendants<C.PieChart>().Count()).IsEqualTo(1);
        await Assert.That(chartSpace.Descendants<C.ValueAxis>()).IsEmpty();
    }

    // ── Validation of the property setters ──────────────────────────────

    [Test]
    public async Task MarkerSizeOutsideExcelsRangeIsRejected()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var series = AddChart(ws, XLChartType.Line).Series.Add("S", "Data!$B$1:$B$2");

        await Assert.That(() => series.MarkerSize = 1).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => series.MarkerSize = 73).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => series.MarkerSize = null).ThrowsNothing();
    }

    [Test]
    public async Task LineWidthOutsideExcelsRangeIsRejected()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var series = AddChart(ws, XLChartType.Line).Series.Add("S", "Data!$B$1:$B$2");

        await Assert.That(() => series.LineWidthPt = -1).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => series.LineWidthPt = 1585).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => series.LineWidthPt = null).ThrowsNothing();
    }

    [Test]
    public async Task TheFormattedChartExampleProducesAValidWorkbook()
    {
        // Left on disk on purpose: this is the file to open in Excel when checking that the
        // formatting renders the way it is meant to.
        var path = Path.Combine(TestContext.TestDirectory!, "FormattedChartExamples.xlsx");
        new XLibur.Examples.Charts.FormattedChartExamples().Create(path);
        Console.Out.WriteLine($"Formatted chart example: {path}");

        // Reloading through XLibur and saving with validation on checks both that the example reads
        // back and that what it wrote is schema-valid.
        using var wb = new XLWorkbook(path);
        using var ms = SaveValidated(wb);
        await Assert.That(wb.Worksheets.Count).IsEqualTo(7);
    }

    [Test]
    public async Task SecondaryAxisCannotBeChangedOnALoadedChart()
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
            var ex = await Assert.That(() => series.UseSecondaryAxis = true).Throws<NotSupportedException>();
            await Assert.That(ex!.Message).Contains("loaded from a file");

            // Assigning the value it already has is not a change, so it is allowed.
            await Assert.That(() => series.UseSecondaryAxis = false).ThrowsNothing();
        }
    }
}
