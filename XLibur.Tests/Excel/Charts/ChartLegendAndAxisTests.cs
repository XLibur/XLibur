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
/// The legend, and the axis title, scale, number format, gridlines and orientation.
/// </summary>
[TestFixture]
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

    private static C.ChartSpace ChartSpaceOf(Stream stream)
    {
        stream.Position = 0;
        using var doc = SpreadsheetDocument.Open(stream, false);
        var chartPart = doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!.ChartParts.First();
        return (C.ChartSpace)chartPart.ChartSpace!.CloneNode(true);
    }

    // ── Legend ──────────────────────────────────────────────────────────

    [Test]
    public void NoLegendIsWrittenUntilItIsAskedFor()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");

        using var ms = SaveValidated(wb);
        Assert.That(ChartSpaceOf(ms).Descendants<C.Legend>(), Is.Empty);
    }

    [Test]
    public void LegendRoundTrips()
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
            Assert.That(legend.Visible, Is.True);
            Assert.That(legend.Position, Is.EqualTo(XLLegendPosition.Bottom));
            Assert.That(legend.Overlay, Is.True);
        }
    }

    [Test]
    public void LegendSitsBetweenThePlotAreaAndPlotVisibleOnly()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
        chart.Legend.Visible = true;

        using var ms = SaveValidated(wb);
        var chartElement = ChartSpaceOf(ms).Elements<C.Chart>().Single();
        var children = chartElement.ChildElements.Select(e => e.LocalName).ToList();

        Assert.That(children.IndexOf("plotArea"), Is.LessThan(children.IndexOf("legend")));
        Assert.That(children.IndexOf("legend"), Is.LessThan(children.IndexOf("plotVisOnly")));
    }

    [Test]
    public void HidingTheLegendOfALoadedChartRemovesIt()
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
            Assert.That(chart.Legend.Visible, Is.True);
            chart.Legend.Visible = false;
            wb.SaveAs(withoutLegend, validate: true);
        }

        Assert.That(ChartSpaceOf(withoutLegend).Descendants<C.Legend>(), Is.Empty);
    }

    [Test]
    public void PositioningALegendThatIsNotThereDoesNotCreateOne()
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
            Assert.That(chart.Legend.Visible, Is.False);
            chart.Legend.Position = XLLegendPosition.Bottom;
            chart.Legend.Overlay = true;
            wb.SaveAs(edited, validate: true);
        }

        Assert.That(ChartSpaceOf(edited).Descendants<C.Legend>(), Is.Empty);
    }

    [Test]
    public void ShowingALegendOnALoadedChartStillHonoursThePositionSetWithIt()
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
            Assert.That(legend.Visible, Is.True);
            Assert.That(legend.Position, Is.EqualTo(XLLegendPosition.Left));
        }
    }

    // ── Axes ────────────────────────────────────────────────────────────

    [Test]
    public void AxisTitleAndNumberFormatRoundTrip()
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
            Assert.That(chart.CategoryAxis.Title, Is.EqualTo("Quarter"));
            Assert.That(chart.ValueAxis.Title, Is.EqualTo("Revenue"));
            Assert.That(chart.ValueAxis.NumberFormat, Is.EqualTo("$ #,##0"));
            Assert.That(chart.CategoryAxis.NumberFormat, Is.Null);
        }
    }

    [Test]
    public void AxisScaleRoundTrips()
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
            Assert.That(axis.Min, Is.EqualTo(0));
            Assert.That(axis.Max, Is.EqualTo(250));
            Assert.That(axis.MajorUnit, Is.EqualTo(50));
            Assert.That(axis.MinorUnit, Is.EqualTo(10));
            Assert.That(axis.MajorGridlines, Is.True);
        }
    }

    [Test]
    public void ScalingChildrenAreWrittenInSchemaOrder()
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

        Assert.That(scaling.ChildElements.Select(e => e.LocalName),
            Is.EqualTo(new[] { "logBase", "orientation", "max", "min" }));
        Assert.That(scaling.Elements<C.LogBase>().Single().Val!.Value, Is.EqualTo(2));
        Assert.That(scaling.Elements<C.Orientation>().Single().Val!.Value,
            Is.EqualTo(C.OrientationValues.MaxMin));
    }

    [Test]
    public void LogScaleRoundTrips()
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
            Assert.That(axis.LogScale, Is.True);
            Assert.That(axis.LogBase, Is.EqualTo(2));
        }
    }

    [Test]
    public void ValueAxisOnlyPropertiesAreSkippedOnTheCategoryAxis()
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

        Assert.That(categoryAxis.Elements<C.MajorUnit>(), Is.Empty);
        Assert.That(categoryAxis.Elements<C.Scaling>().Single().Elements<C.LogBase>(), Is.Empty);
    }

    [Test]
    public void HiddenAxisRoundTrips()
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
            Assert.That(chart.ValueAxis.Visible, Is.False);
            Assert.That(chart.CategoryAxis.Visible, Is.True);
        }
    }

    [Test]
    public void SecondaryValueAxisCanBeTitledAndScaled()
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
            Assert.That(chart.ValueAxis.Title, Is.EqualTo("Units"));
            Assert.That(chart.SecondaryValueAxis.Title, Is.EqualTo("Price"));
            Assert.That(chart.SecondaryValueAxis.Max, Is.EqualTo(10));
        }
    }

    [Test]
    public void ScatterChartsTakeANumberFormatOnBothAxes()
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
            Assert.That(chart.CategoryAxis.Title, Is.EqualTo("X"));
            Assert.That(chart.CategoryAxis.MajorUnit, Is.EqualTo(25),
                "The horizontal axis of a scatter chart is a value axis, so it takes units.");
            Assert.That(chart.ValueAxis.Title, Is.EqualTo("Y"));
        }
    }

    [Test]
    public void AxisTitleTextIsWrittenAsRichText()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = AddChart(ws, XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
        chart.ValueAxis.Title = "Revenue";

        using var ms = SaveValidated(wb);
        var valueAxis = ChartSpaceOf(ms).Descendants<C.ValueAxis>().Single();
        var title = valueAxis.Elements<C.Title>().Single();

        Assert.That(title.Descendants<A.Text>().Single().Text, Is.EqualTo("Revenue"));

        // c:title has to follow c:axPos and precede c:crossAx.
        var children = valueAxis.ChildElements.Select(e => e.LocalName).ToList();
        Assert.That(children.IndexOf("axPos"), Is.LessThan(children.IndexOf("title")));
        Assert.That(children.IndexOf("title"), Is.LessThan(children.IndexOf("crossAx")));
    }

    [Test]
    public void AxesSurviveEveryChartFamilyThatHasThem()
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

            Assert.DoesNotThrow(() =>
            {
                using var ms = SaveValidated(wb);
            }, $"{type} produced invalid chart XML.");
        }
    }

    // ── Validation ──────────────────────────────────────────────────────

    [Test]
    public void AxisUnitsMustBePositive()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var axis = AddChart(ws, XLChartType.ColumnClustered).ValueAxis;

        Assert.Throws<ArgumentOutOfRangeException>(() => axis.MajorUnit = 0);
        Assert.Throws<ArgumentOutOfRangeException>(() => axis.MinorUnit = -1);
        Assert.DoesNotThrow(() => axis.MajorUnit = null);
    }

    [Test]
    public void LogBaseOutsideExcelsRangeIsRejected()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var axis = AddChart(ws, XLChartType.ColumnClustered).ValueAxis;

        Assert.Throws<ArgumentOutOfRangeException>(() => axis.LogBase = 1);
        Assert.Throws<ArgumentOutOfRangeException>(() => axis.LogBase = 1001);
        Assert.DoesNotThrow(() => axis.LogBase = 1000);
    }
}
