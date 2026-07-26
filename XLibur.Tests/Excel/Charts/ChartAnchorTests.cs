using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using System.IO;
using System.Linq;
using XLibur.Excel;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Tests.Excel.Charts;

/// <summary>
/// The three ways a chart can be anchored to the sheet: two-cell, one-cell and absolute.
/// </summary>
[TestFixture]
public class ChartAnchorTests
{
    private static IXLWorksheet AddDataSheet(XLWorkbook wb)
    {
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = "Q1";
        ws.Cell("A2").Value = "Q2";
        ws.Cell("B1").Value = 100;
        ws.Cell("B2").Value = 200;
        return ws;
    }

    private static MemoryStream SaveValidated(XLWorkbook wb)
    {
        var ms = new MemoryStream();
        wb.SaveAs(ms, validate: true);
        ms.Position = 0;
        return ms;
    }

    private static Xdr.WorksheetDrawing DrawingOf(Stream stream)
    {
        stream.Position = 0;
        using var doc = SpreadsheetDocument.Open(stream, false);
        var drawing = doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!.WorksheetDrawing!;
        return (Xdr.WorksheetDrawing)drawing.CloneNode(true);
    }

    [Test]
    public void TwoCellAnchorIsStillTheDefault()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = ws.Charts.Add(XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
        chart.Position.SetColumn(3).SetRow(2);
        chart.SecondPosition.SetColumn(10).SetRow(16);

        Assert.That(chart.Anchor, Is.EqualTo(XLDrawingAnchor.MoveAndSizeWithCells));

        using var ms = SaveValidated(wb);
        var drawing = DrawingOf(ms);
        Assert.That(drawing.Elements<Xdr.TwoCellAnchor>().Count(), Is.EqualTo(1));
        Assert.That(drawing.Elements<Xdr.OneCellAnchor>(), Is.Empty);
    }

    [Test]
    public void OneCellAnchoredChartRoundTrips()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = ws.Charts.Add(XLChartType.ColumnClustered);
            chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Anchor = XLDrawingAnchor.MoveWithCells;
            chart.Position.SetColumn(4).SetRow(3);
            chart.Width = 480;
            chart.Height = 288;

            using var saved = SaveValidated(wb);
            var drawing = DrawingOf(saved);
            Assert.That(drawing.Elements<Xdr.OneCellAnchor>().Count(), Is.EqualTo(1));

            saved.Position = 0;
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.Single();
            Assert.That(chart.Anchor, Is.EqualTo(XLDrawingAnchor.MoveWithCells));
            Assert.That(chart.Position.Column, Is.EqualTo(4));
            Assert.That(chart.Position.Row, Is.EqualTo(3));
            Assert.That(chart.Width, Is.EqualTo(480));
            Assert.That(chart.Height, Is.EqualTo(288));
            Assert.That(chart.Series.Single().Name, Is.EqualTo("Sales"));
        }
    }

    [Test]
    public void AbsolutelyAnchoredChartRoundTrips()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = ws.Charts.Add(XLChartType.Line);
            chart.SetTitle("Pinned");
            chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Anchor = XLDrawingAnchor.Absolute;
            chart.Left = 200;
            chart.Top = 120;
            chart.Width = 400;
            chart.Height = 250;

            using var saved = SaveValidated(wb);
            var drawing = DrawingOf(saved);
            Assert.That(drawing.Elements<Xdr.AbsoluteAnchor>().Count(), Is.EqualTo(1));

            saved.Position = 0;
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.Single();
            Assert.That(chart.Anchor, Is.EqualTo(XLDrawingAnchor.Absolute));
            Assert.That(chart.Title, Is.EqualTo("Pinned"));
            Assert.That(chart.Left, Is.EqualTo(200));
            Assert.That(chart.Top, Is.EqualTo(120));
            Assert.That(chart.Width, Is.EqualTo(400));
            Assert.That(chart.Height, Is.EqualTo(250));
        }
    }

    [Test]
    public void ChartsUnderEveryAnchorKindAreFoundOnOneSheet()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);

            var twoCell = ws.Charts.Add(XLChartType.ColumnClustered);
            twoCell.SetTitle("Two cell");
            twoCell.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            twoCell.Position.SetColumn(3).SetRow(1);
            twoCell.SecondPosition.SetColumn(9).SetRow(14);

            var oneCell = ws.Charts.Add(XLChartType.Line);
            oneCell.SetTitle("One cell");
            oneCell.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            oneCell.Anchor = XLDrawingAnchor.MoveWithCells;
            oneCell.Position.SetColumn(3).SetRow(16);
            oneCell.Width = 400;
            oneCell.Height = 250;

            var absolute = ws.Charts.Add(XLChartType.Pie);
            absolute.SetTitle("Absolute");
            absolute.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            absolute.Anchor = XLDrawingAnchor.Absolute;
            absolute.Left = 700;
            absolute.Top = 20;
            absolute.Width = 320;
            absolute.Height = 240;

            using var saved = SaveValidated(wb);
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var charts = wb.Worksheet("Data").Charts.ToList();
            Assert.That(charts, Has.Count.EqualTo(3),
                "A one-cell or absolute anchored chart used to be skipped on read.");
            Assert.That(charts.Select(c => c.Title),
                Is.EquivalentTo(new[] { "Two cell", "One cell", "Absolute" }));
            Assert.That(charts.Select(c => c.Anchor), Is.EquivalentTo(new[]
            {
                XLDrawingAnchor.MoveAndSizeWithCells,
                XLDrawingAnchor.MoveWithCells,
                XLDrawingAnchor.Absolute
            }));
        }
    }

    [Test]
    public void FormattingAnAnchorlessChartStillReachesItsChartPart()
    {
        // The patcher finds the chart part through its relationship id, not through the anchor, so
        // editing a one-cell anchored chart has to work the same way.
        using var original = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = AddDataSheet(wb);
            var chart = ws.Charts.Add(XLChartType.ColumnClustered);
            chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Anchor = XLDrawingAnchor.MoveWithCells;
            chart.Position.SetColumn(3).SetRow(1);
            chart.Width = 400;
            chart.Height = 250;

            using var saved = SaveValidated(wb);
            saved.CopyTo(original);
        }

        using var edited = new MemoryStream();
        original.Position = 0;
        using (var wb = new XLWorkbook(original))
        {
            wb.Worksheet("Data").Charts.Single().Series.Single().FillColor = XLColor.FromHtml("#C00000");
            wb.SaveAs(edited, validate: true);
        }

        edited.Position = 0;
        using (var wb = new XLWorkbook(edited))
        {
            var chart = wb.Worksheet("Data").Charts.Single();
            Assert.That(chart.Anchor, Is.EqualTo(XLDrawingAnchor.MoveWithCells));
            Assert.That(chart.Series.Single().FillColor, Is.EqualTo(XLColor.FromHtml("#C00000")));
        }
    }
}
