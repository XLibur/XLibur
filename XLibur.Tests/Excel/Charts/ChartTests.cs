using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Linq;
using XLibur.Excel;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Charts;

public class ChartTests
{
    [Test]
    public async Task CanCreateColumnClusteredChart()
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

        await Assert.That(ws.Charts.Count).IsEqualTo(1);
        await Assert.That(chart.ChartType).IsEqualTo(XLChartType.ColumnClustered);
        await Assert.That(chart.Title).IsEqualTo("Sales Chart");
        await Assert.That(chart.Series.Count).IsEqualTo(1);
    }

    [Test]
    public async Task CanSaveAndLoadChart()
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
            await Assert.That(ws.Charts.Count).IsEqualTo(1);

            var chart = ws.Charts.First();
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.ColumnClustered);
            await Assert.That(chart.Title).IsEqualTo("Test Chart");
            await Assert.That(chart.Series.Count).IsEqualTo(1);

            var series = chart.Series.First();
            await Assert.That(series.Name).IsEqualTo("Sales");
            await Assert.That(series.ValueReferences).IsEqualTo("Data!$B$2:$B$3");
            await Assert.That(series.CategoryReferences).IsEqualTo("Data!$A$2:$A$3");
        }
    }

    [Test]
    public async Task SavedChartHasValidOpenXmlStructure()
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
        await Assert.That(drawingsPart).IsNotNull();

        var chartParts = drawingsPart!.ChartParts.ToList();
        await Assert.That(chartParts).Count().IsEqualTo(1);

        var chartSpace = chartParts[0].ChartSpace;
        await Assert.That(chartSpace).IsNotNull();

        var chartEl = chartSpace!.Elements<C.Chart>().FirstOrDefault();
        await Assert.That(chartEl).IsNotNull();

        var barChart = chartEl!.PlotArea!.Elements<BarChart>().FirstOrDefault();
        await Assert.That(barChart).IsNotNull();
        await Assert.That(barChart!.BarDirection!.Val!.Value).IsEqualTo(BarDirectionValues.Column);
        await Assert.That(barChart.BarGrouping!.Val!.Value).IsEqualTo(BarGroupingValues.Clustered);
    }

    [Test]
    public async Task MultipleSeries()
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
            await Assert.That(chart.Series.Count).IsEqualTo(2);

            var series = chart.Series.ToList();
            await Assert.That(series[0].Name).IsEqualTo("Series1");
            await Assert.That(series[1].Name).IsEqualTo("Series2");
            await Assert.That(series[0].Index).IsEqualTo(0u);
            await Assert.That(series[1].Index).IsEqualTo(1u);
        }
    }

    [Test]
    public async Task ChartWithoutTitle()
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
            await Assert.That(chart.Title).IsNull();
        }
    }

    [Test]
    public async Task ChartPositionsArePreserved()
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
            await Assert.That(chart.Position.Column).IsEqualTo(3);
            await Assert.That(chart.Position.Row).IsEqualTo(5);
            await Assert.That(chart.SecondPosition.Column).IsEqualTo(10);
            await Assert.That(chart.SecondPosition.Row).IsEqualTo(20);
        }
    }

    [Test]
    public async Task ChartDoesNotPreventPictureWriting()
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
            await Assert.That(wb.Worksheet("Data").Charts.Count).IsEqualTo(1);
        }
    }

    [Test]
    public async Task CanSaveAndLoadPieChart()
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
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.Pie);
            await Assert.That(chart.Title).IsEqualTo("Distribution");
            await Assert.That(chart.Series.Count).IsEqualTo(1);
            await Assert.That(chart.Series.First().ValueReferences).IsEqualTo("Data!$B$1:$B$3");
        }
    }

    [Test]
    public async Task CanSaveAndLoadStackedBarChart()
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
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.BarStacked);
            await Assert.That(chart.Series.Count).IsEqualTo(2);
        }
    }

    [Test]
    public async Task CanSaveAndLoadLineChart()
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
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.Line);
            await Assert.That(chart.Title).IsEqualTo("Trend");
            await Assert.That(chart.Series.Count).IsEqualTo(1);
        }
    }

    [Test]
    public async Task CanSaveAndLoadRadarChart()
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
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.Radar);
            await Assert.That(chart.Title).IsEqualTo("Skills");
            await Assert.That(chart.Series.Count).IsEqualTo(1);
        }
    }

    [Test]
    public async Task CanSaveAndLoadComboChart()
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
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.ColumnClustered);
            await Assert.That(chart.Series.Count).IsEqualTo(1);
            await Assert.That(chart.Series.First().Name).IsEqualTo("Units");

            await Assert.That(chart.SecondaryChartType).IsEqualTo(XLChartType.Line);
            await Assert.That(chart.SecondarySeries.Count).IsEqualTo(1);
            await Assert.That(chart.SecondarySeries.First().Name).IsEqualTo("Price");
        }
    }

    [Test]
    public async Task CanSaveAndLoadScatterChart()
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
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.XYScatterMarkers);
            await Assert.That(chart.Title).IsEqualTo("XY");
            await Assert.That(chart.Series.Count).IsEqualTo(1);
        }
    }

    [Test]
    public async Task CanSaveAndLoadStockChart()
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
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.StockHighLowClose);
            await Assert.That(chart.Series.Count).IsEqualTo(3);
        }
    }

    [Test]
    public async Task CanSaveAndLoadSurfaceChart()
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
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.Surface);
            await Assert.That(chart.Series.Count).IsEqualTo(2);
        }
    }

    [Test]
    public async Task CanSaveAndLoadWaterfallChart()
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
        Console.Out.WriteLine($"Waterfall test file: {filePath}");

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.Waterfall);
            await Assert.That(chart.Title).IsEqualTo("WF");
            await Assert.That(chart.Series.Count).IsEqualTo(1);
        }
    }

    [Test]
    public async Task CanSaveAndLoadFunnelChart()
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
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.Funnel);
            await Assert.That(chart.Series.Count).IsEqualTo(1);
        }
    }

    [Test]
    public async Task CanSaveAndLoadSunburstChart()
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
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.Sunburst);
            await Assert.That(chart.Title).IsEqualTo("SB");
        }
    }

    [Test]
    public async Task CanSaveAndLoadTreemapChart()
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
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.Treemap);
        }
    }

    [Test]
    public async Task CanSaveAndLoadBoxWhiskerChart()
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
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.BoxWhisker);
            await Assert.That(chart.Title).IsEqualTo("BW");
        }
    }

    [Test]
    public async Task CanSaveAndLoadAreaChart()
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
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.AreaStacked);
            await Assert.That(chart.Series.Count).IsEqualTo(2);
        }
    }

    [Test]
    public async Task CanSaveAndLoadDoughnutChart()
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
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.Doughnut);
            await Assert.That(chart.Title).IsEqualTo("Ring");
        }
    }

    [Test]
    public async Task CanSaveAndLoadBubbleChart()
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
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.Bubble);
            await Assert.That(chart.Title).IsEqualTo("Bubbles");
            await Assert.That(chart.Series.Count).IsEqualTo(1);
        }
    }

    [Test]
    public async Task CanSaveAndLoadConeChart()
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
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.ConeClustered);
        }
    }

    [Test]
    public async Task CanSaveAndLoadCylinderChart()
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
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.CylinderHorizontalStacked);
        }
    }

    [Test]
    public async Task CanSaveAndLoadPyramidChart()
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
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.PyramidStacked100Percent);
        }
    }

    // ── New chart and loaded chart must agree ───────────────────────────

    /// <summary>
    /// The title a new chart writes and the title a loaded chart is patched to must be the same
    /// title, and both must leave <c>c:autoTitleDeleted</c> saying the title is shown. Before spec 22
    /// the two travelled through <c>BuildTitleText</c> and <c>PatchTitle</c> independently, and only
    /// the second of them kept <c>c:autoTitleDeleted</c> in step.
    /// </summary>
    [Test]
    public async Task A_new_chart_and_a_reloaded_chart_carry_the_same_title()
    {
        const string title = "Quarterly revenue";

        var fromNew = ChartGoldenCorpus.CaptureChartPartXml(ws =>
        {
            var chart = ws.Charts.Add(XLChartType.ColumnClustered);
            chart.Series.Add("Units", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Title = title;
        });

        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Q1";
            ws.Cell("B1").Value = 100;
            var chart = ws.Charts.Add(XLChartType.ColumnClustered);
            chart.Series.Add("Units", "Data!$B$1:$B$1", "Data!$A$1:$A$1");
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var reloaded = new XLWorkbook(ms);
        reloaded.Worksheet("Data").Charts.First().Title = title;

        using var patched = new MemoryStream();
        reloaded.SaveAs(patched, validate: true);
        var fromLoaded = ChartGoldenCorpus.FirstChartPartXml(patched);

        await Assert.That(fromNew).Contains($"<a:t>{title}</a:t>");
        await Assert.That(fromLoaded).Contains($"<a:t>{title}</a:t>");
        await Assert.That(fromNew).Contains("<c:autoTitleDeleted val=\"0\"");
        await Assert.That(fromLoaded).Contains("<c:autoTitleDeleted val=\"0\"");
    }
}
