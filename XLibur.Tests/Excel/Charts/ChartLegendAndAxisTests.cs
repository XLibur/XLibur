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
/// The legend, and the axis title, scale, number format, gridlines and orientation.
/// </summary>
public class ChartLegendAndAxisTests
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

    // ── Legend ──────────────────────────────────────────────────────────

    [Test]
    public async Task NoLegendIsWrittenUntilItIsAskedFor()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");

        using var ms = SaveValidated(wb);
        await Assert.That(ChartSpaceOf(ms).Descendants<C.Legend>()).IsEmpty();
    }

    [Test]
    public async Task LegendRoundTrips()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = AddChart(ws, XLChartType.ColumnClustered);
            chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Legend.Visible = true;
            chart.Legend.Position = XLLegendPosition.Bottom;
            chart.Legend.Overlay = true;

            using var saved = SaveValidated(wb);
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var legend = wb.Worksheet("Data").Charts.First().Legend;
            await Assert.That(legend.Visible).IsTrue();
            await Assert.That(legend.Position).IsEqualTo(XLLegendPosition.Bottom);
            await Assert.That(legend.Overlay).IsTrue();
        }
    }

    [Test]
    public async Task LegendSitsBetweenThePlotAreaAndPlotVisibleOnly()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
        chart.Legend.Visible = true;

        using var ms = SaveValidated(wb);
        var chartElement = ChartSpaceOf(ms).Elements<C.Chart>().Single();
        var children = chartElement.ChildElements.Select(e => e.LocalName).ToList();

        await Assert.That(children.IndexOf("plotArea")).IsLessThan(children.IndexOf("legend"));
        await Assert.That(children.IndexOf("legend")).IsLessThan(children.IndexOf("plotVisOnly"));
    }

    [Test]
    public async Task HidingTheLegendOfALoadedChartRemovesIt()
    {
        using var withLegend = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = AddChart(ws, XLChartType.ColumnClustered);
            chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Legend.Visible = true;
            using var saved = SaveValidated(wb);
            saved.CopyTo(withLegend);
        }

        using var withoutLegend = new MemoryStream();
        withLegend.Position = 0;
        using (var wb = new XLWorkbook(withLegend))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            await Assert.That(chart.Legend.Visible).IsTrue();
            chart.Legend.Visible = false;
            wb.SaveAs(withoutLegend, validate: true);
        }

        await Assert.That(ChartSpaceOf(withoutLegend).Descendants<C.Legend>()).IsEmpty();
    }

    [Test]
    public async Task PositioningALegendThatIsNotThereDoesNotCreateOne()
    {
        using var original = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            AddChart(ws, XLChartType.ColumnClustered)
                .Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            using var saved = SaveValidated(wb);
            saved.CopyTo(original);
        }

        using var edited = new MemoryStream();
        original.Position = 0;
        using (var wb = new XLWorkbook(original))
        {
            // Position is documented as ignored while the legend is hidden, and a new chart gets no
            // legend from it either. A loaded one must behave the same way.
            var chart = wb.Worksheet("Data").Charts.First();
            await Assert.That(chart.Legend.Visible).IsFalse();
            chart.Legend.Position = XLLegendPosition.Bottom;
            chart.Legend.Overlay = true;
            wb.SaveAs(edited, validate: true);
        }

        await Assert.That(ChartSpaceOf(edited).Descendants<C.Legend>()).IsEmpty();
    }

    [Test]
    public async Task ShowingALegendOnALoadedChartStillHonoursThePositionSetWithIt()
    {
        using var original = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            AddChart(ws, XLChartType.ColumnClustered)
                .Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            using var saved = SaveValidated(wb);
            saved.CopyTo(original);
        }

        using var edited = new MemoryStream();
        original.Position = 0;
        using (var wb = new XLWorkbook(original))
        {
            var legend = wb.Worksheet("Data").Charts.First().Legend;
            legend.Visible = true;
            legend.Position = XLLegendPosition.Left;
            wb.SaveAs(edited, validate: true);
        }

        edited.Position = 0;
        using (var wb = new XLWorkbook(edited))
        {
            var legend = wb.Worksheet("Data").Charts.First().Legend;
            await Assert.That(legend.Visible).IsTrue();
            await Assert.That(legend.Position).IsEqualTo(XLLegendPosition.Left);
        }
    }

    // ── Axes ────────────────────────────────────────────────────────────

    [Test]
    public async Task AxisTitleAndNumberFormatRoundTrip()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = AddChart(ws, XLChartType.ColumnClustered);
            chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.CategoryAxis.Title = "Quarter";
            chart.ValueAxis.Title = "Revenue";
            chart.ValueAxis.NumberFormat = "$ #,##0";

            using var saved = SaveValidated(wb);
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            await Assert.That(chart.CategoryAxis.Title).IsEqualTo("Quarter");
            await Assert.That(chart.ValueAxis.Title).IsEqualTo("Revenue");
            await Assert.That(chart.ValueAxis.NumberFormat).IsEqualTo("$ #,##0");
            await Assert.That(chart.CategoryAxis.NumberFormat).IsNull();
        }
    }

    [Test]
    public async Task AxisScaleRoundTrips()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = AddChart(ws, XLChartType.ColumnClustered);
            chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.ValueAxis.Min = 0;
            chart.ValueAxis.Max = 250;
            chart.ValueAxis.MajorUnit = 50;
            chart.ValueAxis.MinorUnit = 10;
            chart.ValueAxis.MajorGridlines = true;

            using var saved = SaveValidated(wb);
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var axis = wb.Worksheet("Data").Charts.First().ValueAxis;
            await Assert.That(axis.Min).IsEqualTo(0);
            await Assert.That(axis.Max).IsEqualTo(250);
            await Assert.That(axis.MajorUnit).IsEqualTo(50);
            await Assert.That(axis.MinorUnit).IsEqualTo(10);
            await Assert.That(axis.MajorGridlines).IsTrue();
        }
    }

    [Test]
    public async Task ScalingChildrenAreWrittenInSchemaOrder()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
        chart.ValueAxis.LogScale = true;
        chart.ValueAxis.LogBase = 2;
        chart.ValueAxis.Orientation = XLAxisOrientation.MaxMin;
        chart.ValueAxis.Min = 1;
        chart.ValueAxis.Max = 1000;

        using var ms = SaveValidated(wb);
        var valueAxis = ChartSpaceOf(ms).Descendants<C.ValueAxis>().Single();
        var scaling = valueAxis.Elements<C.Scaling>().Single();

        await Assert.That(scaling.ChildElements.Select(e => e.LocalName)).IsEquivalentTo(new[] { "logBase", "orientation", "max", "min" }, CollectionOrdering.Matching);
        await Assert.That(scaling.Elements<C.LogBase>().Single().Val!.Value).IsEqualTo(2);
        await Assert.That(scaling.Elements<C.Orientation>().Single().Val!.Value).IsEqualTo(C.OrientationValues.MaxMin);
    }

    [Test]
    public async Task LogScaleRoundTrips()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = AddChart(ws, XLChartType.ColumnClustered);
            chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.ValueAxis.LogScale = true;
            chart.ValueAxis.LogBase = 2;

            using var saved = SaveValidated(wb);
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var axis = wb.Worksheet("Data").Charts.First().ValueAxis;
            await Assert.That(axis.LogScale).IsTrue();
            await Assert.That(axis.LogBase).IsEqualTo(2);
        }
    }

    [Test]
    public async Task ValueAxisOnlyPropertiesAreSkippedOnTheCategoryAxis()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
        chart.CategoryAxis.MajorUnit = 1;
        chart.CategoryAxis.LogScale = true;

        // CT_CatAx has neither a unit nor room for c:logBase, so writing them would make Excel
        // refuse the file.
        using var ms = SaveValidated(wb);
        var categoryAxis = ChartSpaceOf(ms).Descendants<C.CategoryAxis>().Single();

        await Assert.That(categoryAxis.Elements<C.MajorUnit>()).IsEmpty();
        await Assert.That(categoryAxis.Elements<C.Scaling>().Single().Elements<C.LogBase>()).IsEmpty();
    }

    [Test]
    public async Task HiddenAxisRoundTrips()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = AddChart(ws, XLChartType.ColumnClustered);
            chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.ValueAxis.Visible = false;

            using var saved = SaveValidated(wb);
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            await Assert.That(chart.ValueAxis.Visible).IsFalse();
            await Assert.That(chart.CategoryAxis.Visible).IsTrue();
        }
    }

    [Test]
    public async Task SecondaryValueAxisCanBeTitledAndScaled()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = AddChart(ws, XLChartType.ColumnClustered);
            chart.Series.Add("Units", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Series.Add("Price", "Data!$C$1:$C$2", "Data!$A$1:$A$2").UseSecondaryAxis = true;

            chart.ValueAxis.Title = "Units";
            chart.SecondaryValueAxis.Title = "Price";
            chart.SecondaryValueAxis.Max = 10;

            using var saved = SaveValidated(wb);
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            await Assert.That(chart.ValueAxis.Title).IsEqualTo("Units");
            await Assert.That(chart.SecondaryValueAxis.Title).IsEqualTo("Price");
            await Assert.That(chart.SecondaryValueAxis.Max).IsEqualTo(10);
        }
    }

    [Test]
    public async Task ScatterChartsTakeANumberFormatOnBothAxes()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = AddChart(ws, XLChartType.XYScatterMarkers);
            chart.Series.Add("Points", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.CategoryAxis.Title = "X";
            chart.CategoryAxis.MajorUnit = 25;
            chart.ValueAxis.Title = "Y";

            using var saved = SaveValidated(wb);
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            await Assert.That(chart.CategoryAxis.Title).IsEqualTo("X");
            await Assert.That(chart.CategoryAxis.MajorUnit).IsEqualTo(25).Because("The horizontal axis of a scatter chart is a value axis, so it takes units.");
            await Assert.That(chart.ValueAxis.Title).IsEqualTo("Y");
        }
    }

    [Test]
    public async Task AxisTitleTextIsWrittenAsRichText()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
        chart.ValueAxis.Title = "Revenue";

        using var ms = SaveValidated(wb);
        var valueAxis = ChartSpaceOf(ms).Descendants<C.ValueAxis>().Single();
        var title = valueAxis.Elements<C.Title>().Single();

        await Assert.That(title.Descendants<A.Text>().Single().Text).IsEqualTo("Revenue");

        // c:title has to follow c:axPos and precede c:crossAx.
        var children = valueAxis.ChildElements.Select(e => e.LocalName).ToList();
        await Assert.That(children.IndexOf("axPos")).IsLessThan(children.IndexOf("title"));
        await Assert.That(children.IndexOf("title")).IsLessThan(children.IndexOf("crossAx"));
    }

    [Test]
    public async Task AxesSurviveEveryChartFamilyThatHasThem()
    {
        XLChartType[] types =
        [
            XLChartType.ColumnClustered, XLChartType.BarStacked, XLChartType.ColumnClustered3D,
            XLChartType.Line, XLChartType.Area, XLChartType.Radar, XLChartType.XYScatterMarkers,
            XLChartType.Bubble, XLChartType.StockHighLowClose, XLChartType.Surface,
            XLChartType.ConeClustered
        ];

        foreach (var type in types)
        {
            using var wb = new XLWorkbook();
            var ws = AddDataSheet(wb);
            var chart = AddChart(ws, type);

            var seriesCount = type == XLChartType.StockHighLowClose ? 3 : 1;
            for (var i = 0; i < seriesCount; i++)
                chart.Series.Add($"S{i}", "Data!$B$1:$B$2", "Data!$A$1:$A$2");

            chart.Legend.Visible = true;
            chart.CategoryAxis.Title = "Category";
            chart.ValueAxis.Title = "Value";
            chart.ValueAxis.NumberFormat = "#,##0";
            chart.ValueAxis.MajorGridlines = true;
            chart.ValueAxis.Min = 0;
            chart.ValueAxis.MajorUnit = 50;
            chart.ValueAxis.Orientation = XLAxisOrientation.MaxMin;

            await Assert.That(() =>
            {
                using var ms = SaveValidated(wb);
            }).ThrowsNothing().Because($"{type} produced invalid chart XML.");
        }
    }

    // ── Validation ──────────────────────────────────────────────────────

    [Test]
    public async Task AxisUnitsMustBePositive()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var axis = AddChart(ws, XLChartType.ColumnClustered).ValueAxis;

        await Assert.That(() => axis.MajorUnit = 0).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => axis.MinorUnit = -1).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => axis.MajorUnit = null).ThrowsNothing();
    }

    [Test]
    public async Task LogBaseOutsideExcelsRangeIsRejected()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var axis = AddChart(ws, XLChartType.ColumnClustered).ValueAxis;

        await Assert.That(() => axis.LogBase = 1).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => axis.LogBase = 1001).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => axis.LogBase = 1000).ThrowsNothing();
    }

    // ── New chart and loaded chart must agree ───────────────────────────

    /// <summary>
    /// A new chart and a chart loaded from a file with no <c>c:legend</c> must reach the same legend
    /// XML from the same model. Before spec 22 these were two functions — <c>BuildLegend</c> and the
    /// null-element branch of <c>PatchLegend</c> — agreeing by hand and by comment; afterwards they
    /// are the two branches of one <c>Apply</c>.
    /// </summary>
    [Test]
    [Arguments(XLLegendPosition.Bottom)]
    [Arguments(XLLegendPosition.Left)]
    [Arguments(XLLegendPosition.TopRight)]
    public async Task A_new_chart_and_a_legendless_loaded_chart_write_the_same_legend(
        XLLegendPosition position)
    {
        var fromNew = ChartGoldenCorpus.CaptureChartPartXml(ws =>
        {
            var chart = ws.Charts.Add(XLChartType.ColumnClustered);
            chart.Series.Add("Units", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Legend.Visible = true;
            chart.Legend.Position = position;
        });

        // A chart saved with no legend at all, reloaded, then given the same legend.
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = ws.Charts.Add(XLChartType.ColumnClustered);
            chart.Series.Add("Units", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var reloaded = new XLWorkbook(ms);
        var loadedLegend = reloaded.Worksheet("Data").Charts.First().Legend;
        loadedLegend.Visible = true;
        loadedLegend.Position = position;

        using var patched = new MemoryStream();
        reloaded.SaveAs(patched, validate: true);

        await Assert.That(LegendPositionIn(ChartGoldenCorpus.FirstChartPartXml(patched)))
            .IsEqualTo(LegendPositionIn(fromNew))
            .Because("The legend a new chart writes and the legend a loaded chart is patched to are the same element.");
        await Assert.That(LegendPositionIn(fromNew)).IsNotNull();
    }

    /// <summary>
    /// The axis properties assigned on a new chart and the same properties assigned on a reloaded one
    /// must reach the same XML. Before spec 22 the two travelled through <c>BuildScaling</c> /
    /// <c>AppendAxisBody</c> / <c>AppendAxisUnits</c> and <c>PatchAxis</c> independently.
    /// </summary>
    [Test]
    [Arguments(true)]
    [Arguments(false)]
    public async Task An_axis_agrees_between_a_new_chart_and_a_reloaded_one(bool gridlines)
    {
        var fromNew = ChartGoldenCorpus.CaptureChartPartXml(ws =>
        {
            var chart = ws.Charts.Add(XLChartType.ColumnClustered);
            chart.Series.Add("Units", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            AssignAxis(chart.ValueAxis, gridlines);
        });

        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = ws.Charts.Add(XLChartType.ColumnClustered);
            chart.Series.Add("Units", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var reloaded = new XLWorkbook(ms);
        AssignAxis(reloaded.Worksheet("Data").Charts.First().ValueAxis, gridlines);

        using var patched = new MemoryStream();
        reloaded.SaveAs(patched, validate: true);
        var fromLoaded = ChartGoldenCorpus.FirstChartPartXml(patched);

        await Assert.That(ValueAxisIn(fromLoaded)).IsEqualTo(ValueAxisIn(fromNew))
            .Because("A new value axis and a patched one are the same element built from the same model.");
    }

    private static void AssignAxis(IXLChartAxis axis, bool gridlines)
    {
        axis.MajorGridlines = gridlines;
        axis.Title = "Units";
        axis.NumberFormat = "#,##0";
        axis.Min = 0;
        axis.Max = 400;
        axis.MajorUnit = 100;
        axis.MinorUnit = 20;
        axis.Orientation = XLAxisOrientation.MaxMin;
    }

    /// <summary>The <c>c:valAx</c> of a chart part, as XML, with the axis identifiers stripped.</summary>
    private static string ValueAxisIn(string chartPartXml)
    {
        var doc = System.Xml.Linq.XDocument.Parse(chartPartXml);
        System.Xml.Linq.XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        var axis = doc.Descendants(c + "valAx").Single();
        foreach (var id in axis.Descendants(c + "axId").Concat(axis.Descendants(c + "crossAx")).ToList())
            id.Remove();
        return axis.ToString();
    }

    private static string? LegendPositionIn(string chartPartXml)
    {
        var doc = System.Xml.Linq.XDocument.Parse(chartPartXml);
        System.Xml.Linq.XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        return doc.Descendants(c + "legendPos").FirstOrDefault()?.Attribute("val")?.Value;
    }
}
