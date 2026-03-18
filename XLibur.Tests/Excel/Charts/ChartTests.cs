using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using System.IO;
using System.Linq;
using XLibur.Excel;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace XLibur.Tests.Excel.Charts;

[TestFixture]
public class ChartTests
{
    [Test]
    public void CanCreateColumnClusteredChart()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = "Category";
        ws.Cell("A2").Value = "Q1";
        ws.Cell("A3").Value = "Q2";
        ws.Cell("B1").Value = "Sales";
        ws.Cell("B2").Value = 100;
        ws.Cell("B3").Value = 200;

        var chart = ws.Charts.Add(XLChartType.ColumnClustered);
        chart.SetTitle("Sales Chart");
        chart.Series.Add("Sales", "Data!$B$2:$B$3", "Data!$A$2:$A$3");
        chart.Position.SetColumn(3).SetRow(1);
        chart.SecondPosition.SetColumn(10).SetRow(15);

        Assert.That(ws.Charts.Count, Is.EqualTo(1));
        Assert.That(chart.ChartType, Is.EqualTo(XLChartType.ColumnClustered));
        Assert.That(chart.Title, Is.EqualTo("Sales Chart"));
        Assert.That(chart.Series.Count, Is.EqualTo(1));
    }

    [Test]
    public void CanSaveAndLoadChart()
    {
        using var ms = new MemoryStream();

        // Create workbook with chart
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Category";
            ws.Cell("A2").Value = "Q1";
            ws.Cell("A3").Value = "Q2";
            ws.Cell("B1").Value = "Sales";
            ws.Cell("B2").Value = 100;
            ws.Cell("B3").Value = 200;

            var chart = ws.Charts.Add(XLChartType.ColumnClustered);
            chart.SetTitle("Test Chart");
            chart.Series.Add("Sales", "Data!$B$2:$B$3", "Data!$A$2:$A$3");
            chart.Position.SetColumn(3).SetRow(1);
            chart.SecondPosition.SetColumn(10).SetRow(15);

            wb.SaveAs(ms);
        }

        // Reload and verify
        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheet("Data");
            Assert.That(ws.Charts.Count, Is.EqualTo(1));

            var chart = ws.Charts.First();
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.ColumnClustered));
            Assert.That(chart.Title, Is.EqualTo("Test Chart"));
            Assert.That(chart.Series.Count, Is.EqualTo(1));

            var series = chart.Series.First();
            Assert.That(series.Name, Is.EqualTo("Sales"));
            Assert.That(series.ValueReferences, Is.EqualTo("Data!$B$2:$B$3"));
            Assert.That(series.CategoryReferences, Is.EqualTo("Data!$A$2:$A$3"));
        }
    }

    [Test]
    public void SavedChartHasValidOpenXmlStructure()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = 10;
            ws.Cell("A2").Value = 20;

            var chart = ws.Charts.Add(XLChartType.ColumnClustered);
            chart.Series.Add("Values", "Sheet1!$A$1:$A$2");
            chart.Position.SetColumn(2).SetRow(0);
            chart.SecondPosition.SetColumn(8).SetRow(12);

            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var wsPart = doc.WorkbookPart!.WorksheetParts.First();
        var drawingsPart = wsPart.DrawingsPart;
        Assert.That(drawingsPart, Is.Not.Null);

        var chartParts = drawingsPart!.ChartParts.ToList();
        Assert.That(chartParts, Has.Count.EqualTo(1));

        var chartSpace = chartParts[0].ChartSpace;
        Assert.That(chartSpace, Is.Not.Null);

        var chartEl = chartSpace!.Elements<C.Chart>().FirstOrDefault();
        Assert.That(chartEl, Is.Not.Null);

        var barChart = chartEl!.PlotArea!.Elements<BarChart>().FirstOrDefault();
        Assert.That(barChart, Is.Not.Null);
        Assert.That(barChart!.BarDirection!.Val!.Value, Is.EqualTo(BarDirectionValues.Column));
        Assert.That(barChart.BarGrouping!.Val!.Value, Is.EqualTo(BarGroupingValues.Clustered));
    }

    [Test]
    public void MultipleSeries()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Q1";
            ws.Cell("A2").Value = "Q2";
            ws.Cell("B1").Value = 100;
            ws.Cell("B2").Value = 200;
            ws.Cell("C1").Value = 150;
            ws.Cell("C2").Value = 250;

            var chart = ws.Charts.Add(XLChartType.ColumnClustered);
            chart.Series.Add("Series1", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Series.Add("Series2", "Data!$C$1:$C$2", "Data!$A$1:$A$2");
            chart.Position.SetColumn(0).SetRow(5);
            chart.SecondPosition.SetColumn(8).SetRow(20);

            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheet("Data");
            var chart = ws.Charts.First();
            Assert.That(chart.Series.Count, Is.EqualTo(2));

            var series = chart.Series.ToList();
            Assert.That(series[0].Name, Is.EqualTo("Series1"));
            Assert.That(series[1].Name, Is.EqualTo("Series2"));
            Assert.That(series[0].Index, Is.EqualTo(0u));
            Assert.That(series[1].Index, Is.EqualTo(1u));
        }
    }

    [Test]
    public void ChartWithoutTitle()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = 10;

            var chart = ws.Charts.Add(XLChartType.ColumnClustered);
            chart.Series.Add("Values", "Data!$A$1:$A$1");
            chart.Position.SetColumn(2).SetRow(0);
            chart.SecondPosition.SetColumn(8).SetRow(12);

            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.Title, Is.Null);
        }
    }

    [Test]
    public void ChartPositionsArePreserved()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = 10;

            var chart = ws.Charts.Add(XLChartType.ColumnClustered);
            chart.Series.Add("Values", "Data!$A$1:$A$1");
            chart.Position.SetColumn(3).SetRow(5);
            chart.SecondPosition.SetColumn(10).SetRow(20);

            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.Position.Column, Is.EqualTo(3));
            Assert.That(chart.Position.Row, Is.EqualTo(5));
            Assert.That(chart.SecondPosition.Column, Is.EqualTo(10));
            Assert.That(chart.SecondPosition.Row, Is.EqualTo(20));
        }
    }

    [Test]
    public void ChartDoesNotPreventPictureWriting()
    {
        // Ensure charts and pictures can coexist
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = 10;

            var chart = ws.Charts.Add(XLChartType.ColumnClustered);
            chart.Series.Add("Values", "Data!$A$1:$A$1");
            chart.Position.SetColumn(0).SetRow(0);
            chart.SecondPosition.SetColumn(5).SetRow(10);

            // Should not throw
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            Assert.That(wb.Worksheet("Data").Charts.Count, Is.EqualTo(1));
        }
    }

    [Test]
    public void CanSaveAndLoadPieChart()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Alpha";
            ws.Cell("A2").Value = "Beta";
            ws.Cell("A3").Value = "Gamma";
            ws.Cell("B1").Value = 40;
            ws.Cell("B2").Value = 35;
            ws.Cell("B3").Value = 25;

            var chart = ws.Charts.Add(XLChartType.Pie);
            chart.SetTitle("Distribution");
            chart.Series.Add("Values", "Data!$B$1:$B$3", "Data!$A$1:$A$3");
            chart.Position.SetColumn(0).SetRow(5);
            chart.SecondPosition.SetColumn(8).SetRow(18);

            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.Pie));
            Assert.That(chart.Title, Is.EqualTo("Distribution"));
            Assert.That(chart.Series.Count, Is.EqualTo(1));
            Assert.That(chart.Series.First().ValueReferences, Is.EqualTo("Data!$B$1:$B$3"));
        }
    }

    [Test]
    public void CanSaveAndLoadStackedBarChart()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "X";
            ws.Cell("A2").Value = "Y";
            ws.Cell("B1").Value = 10;
            ws.Cell("B2").Value = 20;
            ws.Cell("C1").Value = 30;
            ws.Cell("C2").Value = 40;

            var chart = ws.Charts.Add(XLChartType.BarStacked);
            chart.SetTitle("Stacked");
            chart.Series.Add("S1", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Series.Add("S2", "Data!$C$1:$C$2", "Data!$A$1:$A$2");
            chart.Position.SetColumn(0).SetRow(4);
            chart.SecondPosition.SetColumn(8).SetRow(18);

            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.BarStacked));
            Assert.That(chart.Series.Count, Is.EqualTo(2));
        }
    }

    [Test]
    public void CanSaveAndLoadLineChart()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Jan";
            ws.Cell("A2").Value = "Feb";
            ws.Cell("A3").Value = "Mar";
            ws.Cell("B1").Value = 10;
            ws.Cell("B2").Value = 20;
            ws.Cell("B3").Value = 15;

            var chart = ws.Charts.Add(XLChartType.Line);
            chart.SetTitle("Trend");
            chart.Series.Add("Values", "Data!$B$1:$B$3", "Data!$A$1:$A$3");
            chart.Position.SetColumn(0).SetRow(5);
            chart.SecondPosition.SetColumn(8).SetRow(18);

            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.Line));
            Assert.That(chart.Title, Is.EqualTo("Trend"));
            Assert.That(chart.Series.Count, Is.EqualTo(1));
        }
    }

    [Test]
    public void CanSaveAndLoadRadarChart()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Skill1";
            ws.Cell("A2").Value = "Skill2";
            ws.Cell("A3").Value = "Skill3";
            ws.Cell("B1").Value = 8;
            ws.Cell("B2").Value = 6;
            ws.Cell("B3").Value = 9;

            var chart = ws.Charts.Add(XLChartType.Radar);
            chart.SetTitle("Skills");
            chart.Series.Add("Person", "Data!$B$1:$B$3", "Data!$A$1:$A$3");
            chart.Position.SetColumn(0).SetRow(5);
            chart.SecondPosition.SetColumn(8).SetRow(18);

            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.Radar));
            Assert.That(chart.Title, Is.EqualTo("Skills"));
            Assert.That(chart.Series.Count, Is.EqualTo(1));
        }
    }

    [Test]
    public void CanSaveAndLoadComboChart()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Q1";
            ws.Cell("A2").Value = "Q2";
            ws.Cell("B1").Value = 100;
            ws.Cell("B2").Value = 200;
            ws.Cell("C1").Value = 5.5;
            ws.Cell("C2").Value = 6.0;

            var chart = ws.Charts.Add(XLChartType.ColumnClustered);
            chart.SetTitle("Combo");
            chart.Series.Add("Units", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.SecondaryChartType = XLChartType.Line;
            chart.SecondarySeries.Add("Price", "Data!$C$1:$C$2", "Data!$A$1:$A$2");
            chart.Position.SetColumn(0).SetRow(4);
            chart.SecondPosition.SetColumn(8).SetRow(18);

            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.ColumnClustered));
            Assert.That(chart.Series.Count, Is.EqualTo(1));
            Assert.That(chart.Series.First().Name, Is.EqualTo("Units"));

            Assert.That(chart.SecondaryChartType, Is.EqualTo(XLChartType.Line));
            Assert.That(chart.SecondarySeries.Count, Is.EqualTo(1));
            Assert.That(chart.SecondarySeries.First().Name, Is.EqualTo("Price"));
        }
    }

    [Test]
    public void CanSaveAndLoadScatterChart()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = 1.0; ws.Cell("B1").Value = 2.0;
            ws.Cell("A2").Value = 3.0; ws.Cell("B2").Value = 4.0;

            var chart = ws.Charts.Add(XLChartType.XYScatterMarkers);
            chart.SetTitle("XY");
            chart.Series.Add("Points", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Position.SetColumn(0).SetRow(4);
            chart.SecondPosition.SetColumn(8).SetRow(18);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.XYScatterMarkers));
            Assert.That(chart.Title, Is.EqualTo("XY"));
            Assert.That(chart.Series.Count, Is.EqualTo(1));
        }
    }

    [Test]
    public void CanSaveAndLoadStockChart()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Mon"; ws.Cell("B1").Value = 105; ws.Cell("C1").Value = 98; ws.Cell("D1").Value = 102;
            ws.Cell("A2").Value = "Tue"; ws.Cell("B2").Value = 108; ws.Cell("C2").Value = 100; ws.Cell("D2").Value = 104;

            var chart = ws.Charts.Add(XLChartType.StockHighLowClose);
            chart.Series.Add("High", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Series.Add("Low", "Data!$C$1:$C$2", "Data!$A$1:$A$2");
            chart.Series.Add("Close", "Data!$D$1:$D$2", "Data!$A$1:$A$2");
            chart.Position.SetColumn(0).SetRow(4);
            chart.SecondPosition.SetColumn(8).SetRow(18);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.StockHighLowClose));
            Assert.That(chart.Series.Count, Is.EqualTo(3));
        }
    }

    [Test]
    public void CanSaveAndLoadSurfaceChart()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "R1"; ws.Cell("B1").Value = 10; ws.Cell("C1").Value = 20;
            ws.Cell("A2").Value = "R2"; ws.Cell("B2").Value = 30; ws.Cell("C2").Value = 40;

            var chart = ws.Charts.Add(XLChartType.Surface);
            chart.SetTitle("Surface");
            chart.Series.Add("S1", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Series.Add("S2", "Data!$C$1:$C$2", "Data!$A$1:$A$2");
            chart.Position.SetColumn(0).SetRow(4);
            chart.SecondPosition.SetColumn(8).SetRow(18);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.Surface));
            Assert.That(chart.Series.Count, Is.EqualTo(2));
        }
    }

    [Test]
    public void CanSaveAndLoadWaterfallChart()
    {
        // Also write to disk for manual Excel inspection
        var filePath = Path.Combine(Path.GetTempPath(), "WaterfallTest.xlsx");

        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Start"; ws.Cell("B1").Value = 1000;
            ws.Cell("A2").Value = "Add"; ws.Cell("B2").Value = 500;
            ws.Cell("A3").Value = "End"; ws.Cell("B3").Value = 1500;

            var chart = ws.Charts.Add(XLChartType.Waterfall);
            chart.SetTitle("WF");
            chart.Series.Add("Amount", "Data!$B$1:$B$3", "Data!$A$1:$A$3");
            chart.Position.SetColumn(0).SetRow(5);
            chart.SecondPosition.SetColumn(8).SetRow(18);
            wb.SaveAs(ms);
            ms.Position = 0;
            wb.SaveAs(filePath);
        }
        TestContext.Out.WriteLine($"Waterfall test file: {filePath}");

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.Waterfall));
            Assert.That(chart.Title, Is.EqualTo("WF"));
            Assert.That(chart.Series.Count, Is.EqualTo(1));
        }
    }

    [Test]
    public void CanSaveAndLoadFunnelChart()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Stage1"; ws.Cell("B1").Value = 100;
            ws.Cell("A2").Value = "Stage2"; ws.Cell("B2").Value = 60;

            var chart = ws.Charts.Add(XLChartType.Funnel);
            chart.Series.Add("Count", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Position.SetColumn(0).SetRow(4);
            chart.SecondPosition.SetColumn(8).SetRow(18);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.Funnel));
            Assert.That(chart.Series.Count, Is.EqualTo(1));
        }
    }

    [Test]
    public void CanSaveAndLoadSunburstChart()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "A"; ws.Cell("B1").Value = 40;
            ws.Cell("A2").Value = "B"; ws.Cell("B2").Value = 60;

            var chart = ws.Charts.Add(XLChartType.Sunburst);
            chart.SetTitle("SB");
            chart.Series.Add("Values", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Position.SetColumn(0).SetRow(4);
            chart.SecondPosition.SetColumn(8).SetRow(18);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.Sunburst));
            Assert.That(chart.Title, Is.EqualTo("SB"));
        }
    }

    [Test]
    public void CanSaveAndLoadTreemapChart()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "X"; ws.Cell("B1").Value = 50;
            ws.Cell("A2").Value = "Y"; ws.Cell("B2").Value = 30;

            var chart = ws.Charts.Add(XLChartType.Treemap);
            chart.Series.Add("Rev", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Position.SetColumn(0).SetRow(4);
            chart.SecondPosition.SetColumn(8).SetRow(18);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.Treemap));
        }
    }

    [Test]
    public void CanSaveAndLoadBoxWhiskerChart()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "G"; ws.Cell("B1").Value = 10;
            ws.Cell("A2").Value = "G"; ws.Cell("B2").Value = 20;

            var chart = ws.Charts.Add(XLChartType.BoxWhisker);
            chart.SetTitle("BW");
            chart.Series.Add("Val", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Position.SetColumn(0).SetRow(4);
            chart.SecondPosition.SetColumn(8).SetRow(18);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.BoxWhisker));
            Assert.That(chart.Title, Is.EqualTo("BW"));
        }
    }

    [Test]
    public void CanSaveAndLoadAreaChart()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Q1"; ws.Cell("B1").Value = 10; ws.Cell("C1").Value = 20;
            ws.Cell("A2").Value = "Q2"; ws.Cell("B2").Value = 15; ws.Cell("C2").Value = 25;

            var chart = ws.Charts.Add(XLChartType.AreaStacked);
            chart.SetTitle("Area");
            chart.Series.Add("S1", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Series.Add("S2", "Data!$C$1:$C$2", "Data!$A$1:$A$2");
            chart.Position.SetColumn(0).SetRow(4);
            chart.SecondPosition.SetColumn(8).SetRow(18);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.AreaStacked));
            Assert.That(chart.Series.Count, Is.EqualTo(2));
        }
    }

    [Test]
    public void CanSaveAndLoadDoughnutChart()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "X"; ws.Cell("B1").Value = 60;
            ws.Cell("A2").Value = "Y"; ws.Cell("B2").Value = 40;

            var chart = ws.Charts.Add(XLChartType.Doughnut);
            chart.SetTitle("Ring");
            chart.Series.Add("Values", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Position.SetColumn(0).SetRow(4);
            chart.SecondPosition.SetColumn(8).SetRow(18);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.Doughnut));
            Assert.That(chart.Title, Is.EqualTo("Ring"));
        }
    }

    [Test]
    public void CanSaveAndLoadBubbleChart()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = 10; ws.Cell("B1").Value = 20;
            ws.Cell("A2").Value = 30; ws.Cell("B2").Value = 40;

            var chart = ws.Charts.Add(XLChartType.Bubble);
            chart.SetTitle("Bubbles");
            chart.Series.Add("Points", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Position.SetColumn(0).SetRow(4);
            chart.SecondPosition.SetColumn(8).SetRow(18);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.Bubble));
            Assert.That(chart.Title, Is.EqualTo("Bubbles"));
            Assert.That(chart.Series.Count, Is.EqualTo(1));
        }
    }

    [Test]
    public void CanSaveAndLoadConeChart()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "X"; ws.Cell("B1").Value = 10;
            ws.Cell("A2").Value = "Y"; ws.Cell("B2").Value = 20;

            var chart = ws.Charts.Add(XLChartType.ConeClustered);
            chart.Series.Add("S1", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Position.SetColumn(0).SetRow(4);
            chart.SecondPosition.SetColumn(8).SetRow(18);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.ConeClustered));
        }
    }

    [Test]
    public void CanSaveAndLoadCylinderChart()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "X"; ws.Cell("B1").Value = 15;
            ws.Cell("A2").Value = "Y"; ws.Cell("B2").Value = 25;

            var chart = ws.Charts.Add(XLChartType.CylinderHorizontalStacked);
            chart.Series.Add("S1", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Position.SetColumn(0).SetRow(4);
            chart.SecondPosition.SetColumn(8).SetRow(18);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.CylinderHorizontalStacked));
        }
    }

    [Test]
    public void CanSaveAndLoadPyramidChart()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "X"; ws.Cell("B1").Value = 30;
            ws.Cell("A2").Value = "Y"; ws.Cell("B2").Value = 50;

            var chart = ws.Charts.Add(XLChartType.PyramidStacked100Percent);
            chart.Series.Add("S1", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Position.SetColumn(0).SetRow(4);
            chart.SecondPosition.SetColumn(8).SetRow(18);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.PyramidStacked100Percent));
        }
    }
}
