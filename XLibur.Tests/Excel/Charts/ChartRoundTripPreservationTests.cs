using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;
using System.Text;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Charts;

/// <summary>
/// Reading Excel-shaped chart XML, and keeping the parts of it that XLibur does not model when a
/// chart is edited and saved again.
/// </summary>
/// <remarks>
/// The fixture below is hand-written to match the XML Excel emits — the extra <c>a:effectLst</c>,
/// <c>a:round</c> and <c>cap</c> attributes, a scheme colour, a series name held in a
/// <c>c:strRef</c> cache, a trendline, and a secondary axis pair — because no Excel is available to
/// author a resource file in CI.
/// </remarks>
[TestFixture]
public class ChartRoundTripPreservationTests
{
    private const string ExcelShapedChartXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <c:chart>
            <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Units and price</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
            <c:autoTitleDeleted val="0"/>
            <c:plotArea>
              <c:layout/>
              <c:barChart>
                <c:barDir val="col"/>
                <c:grouping val="clustered"/>
                <c:varyColors val="0"/>
                <c:ser>
                  <c:idx val="0"/>
                  <c:order val="0"/>
                  <c:tx><c:strRef><c:f>Data!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Units</c:v></c:pt></c:strCache></c:strRef></c:tx>
                  <c:spPr>
                    <a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill>
                    <a:ln w="19050" cap="rnd"><a:solidFill><a:srgbClr val="203864"/></a:solidFill><a:round/></a:ln>
                    <a:effectLst/>
                  </c:spPr>
                  <c:invertIfNegative val="0"/>
                  <c:trendline><c:trendlineType val="linear"/></c:trendline>
                  <c:cat><c:strRef><c:f>Data!$A$1:$A$2</c:f></c:strRef></c:cat>
                  <c:val><c:numRef><c:f>Data!$B$1:$B$2</c:f></c:numRef></c:val>
                </c:ser>
                <c:gapWidth val="219"/>
                <c:overlap val="-27"/>
                <c:axId val="111111111"/>
                <c:axId val="222222222"/>
              </c:barChart>
              <c:lineChart>
                <c:grouping val="standard"/>
                <c:varyColors val="0"/>
                <c:ser>
                  <c:idx val="1"/>
                  <c:order val="1"/>
                  <c:tx><c:v>Price</c:v></c:tx>
                  <c:spPr>
                    <a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill><a:round/></a:ln>
                    <a:effectLst/>
                  </c:spPr>
                  <c:marker>
                    <c:symbol val="circle"/>
                    <c:size val="7"/>
                    <c:spPr><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr>
                  </c:marker>
                  <c:cat><c:strRef><c:f>Data!$A$1:$A$2</c:f></c:strRef></c:cat>
                  <c:val><c:numRef><c:f>Data!$C$1:$C$2</c:f></c:numRef></c:val>
                  <c:smooth val="1"/>
                </c:ser>
                <c:marker val="1"/>
                <c:axId val="333333333"/>
                <c:axId val="444444444"/>
              </c:lineChart>
              <c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:catAx>
              <c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111111111"/></c:valAx>
              <c:valAx><c:axId val="444444444"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="r"/><c:crossAx val="333333333"/><c:crosses val="max"/></c:valAx>
              <c:catAx><c:axId val="333333333"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:crossAx val="444444444"/></c:catAx>
            </c:plotArea>
            <c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>
            <c:plotVisOnly val="1"/>
          </c:chart>
        </c:chartSpace>
        """;

    /// <summary>
    /// Produces a workbook whose single chart part holds <see cref="ExcelShapedChartXml"/>.
    /// </summary>
    private static MemoryStream CreateWorkbookWithExcelShapedChart()
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Q1";
            ws.Cell("A2").Value = "Q2";
            ws.Cell("B1").Value = 100;
            ws.Cell("B2").Value = 200;
            ws.Cell("C1").Value = 5;
            ws.Cell("C2").Value = 8;

            var chart = ws.Charts.Add(XLChartType.ColumnClustered);
            chart.Series.Add("Units", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Position.SetColumn(5).SetRow(1);
            chart.SecondPosition.SetColumn(12).SetRow(15);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var doc = SpreadsheetDocument.Open(ms, true))
        {
            var chartPart = doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!.ChartParts.First();
            using var source = new MemoryStream(Encoding.UTF8.GetBytes(ExcelShapedChartXml));
            chartPart.FeedData(source);
        }

        ms.Position = 0;
        return ms;
    }

    private static string ReadChartXml(Stream stream)
    {
        stream.Position = 0;
        using var doc = SpreadsheetDocument.Open(stream, false);
        var chartPart = doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!.ChartParts.First();
        using var reader = new StreamReader(chartPart.GetStream(FileMode.Open, FileAccess.Read));
        return reader.ReadToEnd();
    }

    // ── Reading Excel-shaped XML ────────────────────────────────────────

    [Test]
    public void ExcelShapedSeriesFormattingIsRead()
    {
        using var ms = CreateWorkbookWithExcelShapedChart();
        using var wb = new XLWorkbook(ms);
        var chart = wb.Worksheet("Data").Charts.First();

        Assert.That(chart.Title, Is.EqualTo("Units and price"));
        Assert.That(chart.ChartType, Is.EqualTo(XLChartType.ColumnClustered));
        Assert.That(chart.SecondaryChartType, Is.EqualTo(XLChartType.LineWithMarkers));

        var bar = chart.Series.Single();
        Assert.That(bar.Name, Is.EqualTo("Units"), "The name lives in the c:strRef cache.");
        Assert.That(bar.FillColor, Is.EqualTo(XLColor.FromHtml("#ED7D31")));
        Assert.That(bar.LineColor, Is.EqualTo(XLColor.FromHtml("#203864")));
        Assert.That(bar.LineWidthPt, Is.EqualTo(1.5));
        Assert.That(bar.UseSecondaryAxis, Is.False);

        var line = chart.SecondarySeries.Single();
        Assert.That(line.Name, Is.EqualTo("Price"), "The name is a literal c:v.");
        Assert.That(line.LineColor!.ColorType, Is.EqualTo(XLColorType.Theme));
        Assert.That(line.LineColor.ThemeColor, Is.EqualTo(XLThemeColor.Accent2));
        Assert.That(line.LineWidthPt, Is.EqualTo(2.25));
        Assert.That(line.MarkerStyle, Is.EqualTo(XLMarkerStyle.Circle));
        Assert.That(line.MarkerSize, Is.EqualTo(7));
        Assert.That(line.MarkerFillColor, Is.EqualTo(XLColor.FromHtml("#70AD47")));
        Assert.That(line.Smooth, Is.True);
        Assert.That(line.UseSecondaryAxis, Is.True,
            "The line group is plotted against the value axis on the right.");
    }

    // ── Preservation ────────────────────────────────────────────────────

    [Test]
    public void LoadAndSaveWithoutEditsLeavesTheChartPartUntouched()
    {
        using var original = CreateWorkbookWithExcelShapedChart();
        var before = ReadChartXml(original);

        using var saved = new MemoryStream();
        original.Position = 0;
        using (var wb = new XLWorkbook(original))
        {
            wb.SaveAs(saved);
        }

        Assert.That(ReadChartXml(saved), Is.EqualTo(before));
    }

    [Test]
    public void EditingOneSeriesKeepsEverythingElseInTheChartPart()
    {
        using var original = CreateWorkbookWithExcelShapedChart();
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            var series = wb.Worksheet("Data").Charts.First().Series.Single();
            series.FillColor = XLColor.FromHtml("#00B050");
            series.LineWidthPt = 3;
            wb.SaveAs(saved);
        }

        var xml = ReadChartXml(saved);

        // The edits landed.
        Assert.That(xml, Does.Contain("00B050"));
        Assert.That(xml, Does.Not.Contain("ED7D31"), "The old fill colour must be replaced, not doubled up.");
        Assert.That(xml, Does.Contain("w=\"38100\""), "3 pt is 38100 EMU.");

        // Everything XLibur does not model survived.
        Assert.That(xml, Does.Contain("<c:trendline>"));
        Assert.That(xml, Does.Contain("cap=\"rnd\""));
        Assert.That(xml, Does.Contain("<a:round"));
        Assert.That(xml, Does.Contain("<a:effectLst"));
        Assert.That(xml, Does.Contain("<c:gapWidth val=\"219\""));
        Assert.That(xml, Does.Contain("<c:legend>"));

        // The untouched line series kept its own formatting.
        Assert.That(xml, Does.Contain("accent2"));
        Assert.That(xml, Does.Contain("<c:size val=\"7\""));
        Assert.That(xml, Does.Contain("70AD47"));
    }

    [Test]
    public void EditedSeriesFormattingReloadsWithTheNewValues()
    {
        using var original = CreateWorkbookWithExcelShapedChart();
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            chart.Series.Single().FillColor = XLColor.FromHtml("#00B050");

            var line = chart.SecondarySeries.Single();
            line.MarkerStyle = XLMarkerStyle.Square;
            line.MarkerSize = 10;
            line.Smooth = false;
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using (var wb = new XLWorkbook(saved))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            Assert.That(chart.Series.Single().FillColor, Is.EqualTo(XLColor.FromHtml("#00B050")));

            var line = chart.SecondarySeries.Single();
            Assert.That(line.MarkerStyle, Is.EqualTo(XLMarkerStyle.Square));
            Assert.That(line.MarkerSize, Is.EqualTo(10));
            Assert.That(line.Smooth, Is.False);
            Assert.That(line.MarkerFillColor, Is.EqualTo(XLColor.FromHtml("#70AD47")),
                "A property that was not assigned keeps the value from the file.");
        }
    }

    [Test]
    public void ClearingAColorRemovesTheFillRatherThanWritingBlack()
    {
        using var original = CreateWorkbookWithExcelShapedChart();
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            wb.Worksheet("Data").Charts.First().Series.Single().FillColor = null;
            wb.SaveAs(saved);
        }

        var xml = ReadChartXml(saved);
        Assert.That(xml, Does.Not.Contain("ED7D31"));
        Assert.That(xml, Does.Not.Contain("<a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill>"));

        saved.Position = 0;
        using (var wb = new XLWorkbook(saved))
        {
            Assert.That(wb.Worksheet("Data").Charts.First().Series.Single().FillColor, Is.Null);
        }
    }

    [Test]
    public void MarkerAndSmoothAreNotPatchedIntoASeriesTypeThatCannotHoldThem()
    {
        using var original = CreateWorkbookWithExcelShapedChart();
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            // A bar series has neither c:marker nor c:smooth in its schema.
            var bar = wb.Worksheet("Data").Charts.First().Series.Single();
            bar.MarkerStyle = XLMarkerStyle.Circle;
            bar.MarkerSize = 10;
            bar.Smooth = true;
            bar.FillColor = XLColor.FromHtml("#00B050");

            wb.SaveAs(saved, validate: true);
        }

        var xml = ReadChartXml(saved);
        var barSeries = xml[xml.IndexOf("<c:barChart>", StringComparison.Ordinal)..
                            xml.IndexOf("<c:lineChart>", StringComparison.Ordinal)];
        Assert.That(barSeries, Does.Not.Contain("<c:marker"));
        Assert.That(barSeries, Does.Not.Contain("<c:smooth"));
        Assert.That(barSeries, Does.Contain("00B050"), "The fill still applies.");
    }

    [Test]
    public void SavingANewChartTwiceDoesNotDuplicateIt()
    {
        using var first = new MemoryStream();
        using var second = new MemoryStream();

        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Q1";
            ws.Cell("B1").Value = 100;

            var chart = ws.Charts.Add(XLChartType.ColumnClustered);
            var series = chart.Series.Add("Units", "Data!$B$1:$B$1", "Data!$A$1:$A$1");
            chart.Position.SetColumn(5).SetRow(1);
            chart.SecondPosition.SetColumn(12).SetRow(15);
            wb.SaveAs(first);

            // The second save has to patch the part written by the first one.
            series.FillColor = XLColor.FromHtml("#00B050");
            wb.SaveAs(second);
        }

        second.Position = 0;
        using (var wb = new XLWorkbook(second))
        {
            var ws = wb.Worksheet("Data");
            Assert.That(ws.Charts.Count, Is.EqualTo(1));
            Assert.That(ws.Charts.First().Series.Single().FillColor, Is.EqualTo(XLColor.FromHtml("#00B050")));
        }
    }
}
