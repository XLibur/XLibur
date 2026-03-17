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
}
