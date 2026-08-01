using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Linq;
using XLibur.Excel;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;

namespace XLibur.Tests.Excel.Charts;

/// <summary>
/// The three ways a chart can be anchored to the sheet: two-cell, one-cell and absolute.
/// </summary>
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

    private static Xdr.WorksheetDrawing DrawingOf(MemoryStream stream)
    {
        stream.Position = 0;
        using var doc = SpreadsheetDocument.Open(stream, false);
        var drawing = doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!.WorksheetDrawing!;
        return (Xdr.WorksheetDrawing)drawing.CloneNode(true);
    }

    [Test]
    public async Task TwoCellAnchorIsStillTheDefault()
    {
        using var wb = new XLWorkbook();
        var ws = AddDataSheet(wb);
        var chart = ws.Charts.Add(XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
        chart.Position.SetColumn(3).SetRow(2);
        chart.SecondPosition.SetColumn(10).SetRow(16);

        await Assert.That(chart.Anchor).IsEqualTo(XLDrawingAnchor.MoveAndSizeWithCells);

        using var ms = SaveValidated(wb);
        var drawing = DrawingOf(ms);
        await Assert.That(drawing.Elements<Xdr.TwoCellAnchor>().Count()).IsEqualTo(1);
        await Assert.That(drawing.Elements<Xdr.OneCellAnchor>()).IsEmpty();
    }

    [Test]
    public async Task OneCellAnchoredChartRoundTrips()
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
            await Assert.That(drawing.Elements<Xdr.OneCellAnchor>().Count()).IsEqualTo(1);

            saved.Position = 0;
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.Single();
            await Assert.That(chart.Anchor).IsEqualTo(XLDrawingAnchor.MoveWithCells);
            await Assert.That(chart.Position.Column).IsEqualTo(4);
            await Assert.That(chart.Position.Row).IsEqualTo(3);
            await Assert.That(chart.Width).IsEqualTo(480);
            await Assert.That(chart.Height).IsEqualTo(288);
            await Assert.That(chart.Series.Single().Name).IsEqualTo("Sales");
        }
    }

    [Test]
    public async Task AbsolutelyAnchoredChartRoundTrips()
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
            await Assert.That(drawing.Elements<Xdr.AbsoluteAnchor>().Count()).IsEqualTo(1);

            saved.Position = 0;
            saved.CopyTo(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var chart = wb.Worksheet("Data").Charts.Single();
            await Assert.That(chart.Anchor).IsEqualTo(XLDrawingAnchor.Absolute);
            await Assert.That(chart.Title).IsEqualTo("Pinned");
            await Assert.That(chart.Left).IsEqualTo(200);
            await Assert.That(chart.Top).IsEqualTo(120);
            await Assert.That(chart.Width).IsEqualTo(400);
            await Assert.That(chart.Height).IsEqualTo(250);
        }
    }

    [Test]
    public async Task ChartsUnderEveryAnchorKindAreFoundOnOneSheet()
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
            await Assert.That(charts.Count).IsEqualTo(3).Because("A one-cell or absolute anchored chart used to be skipped on read.");
            await Assert.That(charts.Select(c => c.Title)).IsEquivalentTo(new[] { "Two cell", "One cell", "Absolute" }, CollectionOrdering.Matching);
            await Assert.That(charts.Select(c => c.Anchor)).IsEquivalentTo(new[]
            {
                XLDrawingAnchor.MoveAndSizeWithCells,
                XLDrawingAnchor.MoveWithCells,
                XLDrawingAnchor.Absolute
            }, CollectionOrdering.Matching);
        }
    }

    [Test]
    public async Task FormattingAnAnchorlessChartStillReachesItsChartPart()
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
            await Assert.That(chart.Anchor).IsEqualTo(XLDrawingAnchor.MoveWithCells);
            await Assert.That(chart.Series.Single().FillColor).IsEqualTo(XLColor.FromHtml("#C00000"));
        }
    }
}
