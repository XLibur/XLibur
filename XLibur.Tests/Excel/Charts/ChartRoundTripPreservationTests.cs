using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Linq;
using System.Text;
using XLibur.Excel;
using System.Threading.Tasks;

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
                  <c:dLbls>
                    <c:dLbl><c:idx val="0"/><c:delete val="1"/></c:dLbl>
                    <c:numFmt formatCode="#,##0" sourceLinked="0"/>
                    <c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
                    <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>
                    <c:dLblPos val="outEnd"/>
                    <c:showLegendKey val="0"/>
                    <c:showVal val="1"/>
                    <c:showCatName val="0"/>
                    <c:showSerName val="0"/>
                    <c:showPercent val="0"/>
                    <c:showBubbleSize val="0"/>
                  </c:dLbls>
                  <c:trendline><c:trendlineType val="linear"/></c:trendline>
                  <c:errBars>
                    <c:errDir val="y"/>
                    <c:errBarType val="both"/>
                    <c:errValType val="percentage"/>
                    <c:noEndCap val="0"/>
                    <c:val val="5"/>
                  </c:errBars>
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
              <c:valAx>
                <c:axId val="222222222"/>
                <c:scaling><c:orientation val="minMax"/></c:scaling>
                <c:delete val="0"/>
                <c:axPos val="l"/>
                <c:majorGridlines/>
                <c:title><c:tx><c:rich><a:bodyPr rot="-5400000" vert="horz"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Units</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
                <c:numFmt formatCode="#,##0" sourceLinked="0"/>
                <c:majorTickMark val="out"/>
                <c:minorTickMark val="none"/>
                <c:tickLblPos val="nextTo"/>
                <c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
                <c:txPr><a:bodyPr rot="-60000000" spcFirstLastPara="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>
                <c:crossAx val="111111111"/>
                <c:crossBetween val="between"/>
              </c:valAx>
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
    private static MemoryStream CreateWorkbookWithExcelShapedChart() =>
        CreateWorkbookWithExcelShapedChart(ExcelShapedChartXml);

    /// <summary>
    /// Produces a workbook whose single chart part holds the given XML.
    /// </summary>
    private static MemoryStream CreateWorkbookWithExcelShapedChart(string chartXml)
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
            using var source = new MemoryStream(Encoding.UTF8.GetBytes(chartXml));
            chartPart.FeedData(source);
        }

        ms.Position = 0;
        return ms;
    }

    private static string ReadChartXml(MemoryStream stream)
    {
        stream.Position = 0;
        using var doc = SpreadsheetDocument.Open(stream, false);
        var chartPart = doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!.ChartParts.First();
        using var reader = new StreamReader(chartPart.GetStream(FileMode.Open, FileAccess.Read));
        return reader.ReadToEnd();
    }

    // ── Reading Excel-shaped XML ────────────────────────────────────────

    [Test]
    public async Task ExcelShapedSeriesFormattingIsRead()
    {
        using var ms = CreateWorkbookWithExcelShapedChart();
        using var wb = new XLWorkbook(ms);
        var chart = wb.Worksheet("Data").Charts.First();

        await Assert.That(chart.Title).IsEqualTo("Units and price");
        await Assert.That(chart.ChartType).IsEqualTo(XLChartType.ColumnClustered);
        await Assert.That(chart.SecondaryChartType).IsEqualTo(XLChartType.LineWithMarkers);

        var bar = chart.Series.Single();
        await Assert.That(bar.Name).IsEqualTo("Units").Because("The name lives in the c:strRef cache.");
        await Assert.That(bar.FillColor).IsEqualTo(XLColor.FromHtml("#ED7D31"));
        await Assert.That(bar.LineColor).IsEqualTo(XLColor.FromHtml("#203864"));
        await Assert.That(bar.LineWidthPt).IsEqualTo(1.5);
        await Assert.That(bar.UseSecondaryAxis).IsFalse();

        await Assert.That(bar.DataLabels.ShowValue).IsTrue();
        await Assert.That(bar.DataLabels.ShowCategoryName).IsFalse();
        await Assert.That(bar.DataLabels.NumberFormat).IsEqualTo("#,##0");
        await Assert.That(bar.DataLabels.Position).IsEqualTo(XLDataLabelPosition.OutsideEnd);

        var line = chart.SecondarySeries.Single();
        await Assert.That(line.Name).IsEqualTo("Price").Because("The name is a literal c:v.");
        await Assert.That(line.LineColor!.ColorType).IsEqualTo(XLColorType.Theme);
        await Assert.That(line.LineColor.ThemeColor).IsEqualTo(XLThemeColor.Accent2);
        await Assert.That(line.LineWidthPt).IsEqualTo(2.25);
        await Assert.That(line.MarkerStyle).IsEqualTo(XLMarkerStyle.Circle);
        await Assert.That(line.MarkerSize).IsEqualTo(7);
        await Assert.That(line.MarkerFillColor).IsEqualTo(XLColor.FromHtml("#70AD47"));
        await Assert.That(line.Smooth).IsTrue();
        await Assert.That(line.UseSecondaryAxis).IsTrue().Because("The line group is plotted against the value axis on the right.");
    }

    [Test]
    public async Task ExcelShapedLegendAndAxesAreRead()
    {
        using var ms = CreateWorkbookWithExcelShapedChart();
        using var wb = new XLWorkbook(ms);
        var chart = wb.Worksheet("Data").Charts.First();

        await Assert.That(chart.Legend.Visible).IsTrue();
        await Assert.That(chart.Legend.Position).IsEqualTo(XLLegendPosition.Bottom);
        await Assert.That(chart.Legend.Overlay).IsFalse();

        await Assert.That(chart.CategoryAxis.Visible).IsTrue();
        await Assert.That(chart.CategoryAxis.Title).IsNull();

        await Assert.That(chart.ValueAxis.Title).IsEqualTo("Units");
        await Assert.That(chart.ValueAxis.NumberFormat).IsEqualTo("#,##0");
        await Assert.That(chart.ValueAxis.MajorGridlines).IsTrue();
        await Assert.That(chart.ValueAxis.Visible).IsTrue();
        await Assert.That(chart.ValueAxis.LogScale).IsFalse();

        await Assert.That(chart.SecondaryValueAxis.Visible).IsTrue().Because("The fixture's line group hangs off a second value axis.");
    }

    // ── Preservation ────────────────────────────────────────────────────

    [Test]
    public async Task LoadAndSaveWithoutEditsLeavesTheChartPartUntouched()
    {
        using var original = CreateWorkbookWithExcelShapedChart();
        var before = ReadChartXml(original);

        using var saved = new MemoryStream();
        original.Position = 0;
        using (var wb = new XLWorkbook(original))
        {
            wb.SaveAs(saved);
        }

        await Assert.That(ReadChartXml(saved)).IsEqualTo(before);
    }

    [Test]
    public async Task EditingOneSeriesKeepsEverythingElseInTheChartPart()
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
        await Assert.That(xml).Contains("00B050");
        await Assert.That(xml).DoesNotContain("ED7D31").Because("The old fill colour must be replaced, not doubled up.");
        await Assert.That(xml).Contains("w=\"38100\"").Because("3 pt is 38100 EMU.");

        // Everything XLibur does not model survived.
        await Assert.That(xml).Contains("<c:trendline>");
        await Assert.That(xml).Contains("<c:errBars>");
        await Assert.That(xml).Contains("<c:errValType val=\"percentage\"");
        await Assert.That(xml).Contains("cap=\"rnd\"");
        await Assert.That(xml).Contains("<a:round");
        await Assert.That(xml).Contains("<a:effectLst");
        await Assert.That(xml).Contains("<c:gapWidth val=\"219\"");
        await Assert.That(xml).Contains("<c:legend>");

        // The untouched line series kept its own formatting.
        await Assert.That(xml).Contains("accent2");
        await Assert.That(xml).Contains("<c:size val=\"7\"");
        await Assert.That(xml).Contains("70AD47");
    }

    [Test]
    public async Task EditedSeriesFormattingReloadsWithTheNewValues()
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
            await Assert.That(chart.Series.Single().FillColor).IsEqualTo(XLColor.FromHtml("#00B050"));

            var line = chart.SecondarySeries.Single();
            await Assert.That(line.MarkerStyle).IsEqualTo(XLMarkerStyle.Square);
            await Assert.That(line.MarkerSize).IsEqualTo(10);
            await Assert.That(line.Smooth).IsFalse();
            await Assert.That(line.MarkerFillColor).IsEqualTo(XLColor.FromHtml("#70AD47")).Because("A property that was not assigned keeps the value from the file.");
        }
    }

    [Test]
    public async Task ClearingAColorRemovesTheFillRatherThanWritingBlack()
    {
        using var original = CreateWorkbookWithExcelShapedChart();
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            wb.Worksheet("Data").Charts.First().Series.Single().FillColor = null;
            wb.SaveAs(saved);
        }

        var xml = ReadChartXml(saved);
        await Assert.That(xml).DoesNotContain("ED7D31");
        await Assert.That(xml).DoesNotContain("<a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill>");

        saved.Position = 0;
        using (var wb = new XLWorkbook(saved))
        {
            await Assert.That(wb.Worksheet("Data").Charts.First().Series.Single().FillColor).IsNull();
        }
    }

    [Test]
    public async Task EditingDataLabelsKeepsTheirTextAndPerPointOverrides()
    {
        using var original = CreateWorkbookWithExcelShapedChart();
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            var labels = wb.Worksheet("Data").Charts.First().Series.Single().DataLabels;
            labels.ShowCategoryName = true;
            labels.Position = XLDataLabelPosition.InsideEnd;
            wb.SaveAs(saved, validate: true);
        }

        var xml = ReadChartXml(saved);

        await Assert.That(xml).Contains("val=\"inEnd\"");
        await Assert.That(xml).DoesNotContain("val=\"outEnd\"");
        await Assert.That(xml).Contains("<c:showCatName val=\"1\"");
        await Assert.That(xml).Contains("<c:showVal val=\"1\"").Because("A flag nobody touched keeps its value.");

        // Label formatting XLibur does not model, and the hidden first point, are still there.
        await Assert.That(xml).Contains("<c:txPr>");
        await Assert.That(xml).Contains("sz=\"900\"");
        await Assert.That(xml).Contains("<c:dLbl>");
        await Assert.That(xml).Contains("formatCode=\"#,##0\"").Because("The number format was not touched either.");
    }

    [Test]
    public async Task DataLabelsCanBeAddedToASeriesThatHadNone()
    {
        using var original = CreateWorkbookWithExcelShapedChart();
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            // The line series in the fixture carries no c:dLbls at all.
            var labels = wb.Worksheet("Data").Charts.First().SecondarySeries.Single().DataLabels;
            labels.ShowValue = true;
            labels.Position = XLDataLabelPosition.Above;
            wb.SaveAs(saved, validate: true);
        }

        saved.Position = 0;
        using (var wb = new XLWorkbook(saved))
        {
            var labels = wb.Worksheet("Data").Charts.First().SecondarySeries.Single().DataLabels;
            await Assert.That(labels.ShowValue).IsTrue();
            await Assert.That(labels.Position).IsEqualTo(XLDataLabelPosition.Above);
        }
    }

    [Test]
    public async Task ChartWideLabelsReachEveryGroupOfALoadedComboChart()
    {
        using var original = CreateWorkbookWithExcelShapedChart();
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            // Chart-wide labels apply to the whole chart, and the writer puts them on each group of a
            // new combo chart. A loaded one used to get them on the c:barChart only.
            wb.Worksheet("Data").Charts.First().DataLabels.ShowValue = true;
            wb.SaveAs(saved, validate: true);
        }

        var xml = ReadChartXml(saved);
        var barChart = xml[xml.IndexOf("<c:barChart>", StringComparison.Ordinal)..xml.IndexOf("</c:barChart>", StringComparison.Ordinal)];
        var lineChart = xml[xml.IndexOf("<c:lineChart>", StringComparison.Ordinal)..xml.IndexOf("</c:lineChart>", StringComparison.Ordinal)];

        // The group-level c:dLbls follows the last c:ser, so it is what comes after </c:ser>.
        await Assert.That(barChart[barChart.LastIndexOf("</c:ser>", StringComparison.Ordinal)..]).Contains("<c:showVal val=\"1\"");
        await Assert.That(lineChart[lineChart.LastIndexOf("</c:ser>", StringComparison.Ordinal)..]).Contains("<c:showVal val=\"1\"");
    }

    [Test]
    public async Task TurningLabelsOnClearsAnExistingDeleteFlag()
    {
        // Excel writes <c:dLbls><c:delete val="1"/></c:dLbls> for a series whose labels are switched
        // off. A c:delete left in place overrides every show flag next to it.
        var start = ExcelShapedChartXml.IndexOf("<c:dLbls>", StringComparison.Ordinal);
        var end = ExcelShapedChartXml.IndexOf("</c:dLbls>", StringComparison.Ordinal) + "</c:dLbls>".Length;
        var chartXml = ExcelShapedChartXml[..start]
                       + "<c:dLbls><c:delete val=\"1\"/></c:dLbls>"
                       + ExcelShapedChartXml[end..];

        using var original = CreateWorkbookWithExcelShapedChart(chartXml);
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            wb.Worksheet("Data").Charts.First().Series.Single().DataLabels.ShowValue = true;
            wb.SaveAs(saved, validate: true);
        }

        var xml = ReadChartXml(saved);
        var labels = xml[xml.IndexOf("<c:dLbls>", StringComparison.Ordinal)..xml.IndexOf("</c:dLbls>", StringComparison.Ordinal)];
        await Assert.That(labels).Contains("<c:showVal val=\"1\"");
        await Assert.That(labels).DoesNotContain("<c:delete");
    }

    [Test]
    public async Task EditingAnAxisKeepsItsTickMarksAndTextFormatting()
    {
        using var original = CreateWorkbookWithExcelShapedChart();
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            chart.ValueAxis.Title = "Units sold";
            chart.ValueAxis.Min = 0;
            chart.ValueAxis.Max = 250;
            chart.ValueAxis.MajorGridlines = false;
            chart.Legend.Position = XLLegendPosition.Right;
            wb.SaveAs(saved, validate: true);
        }

        var xml = ReadChartXml(saved);

        // The edits landed.
        await Assert.That(xml).Contains("Units sold");
        await Assert.That(xml).DoesNotContain("<a:t>Units</a:t>");
        await Assert.That(xml).Contains("<c:max val=\"250\"");
        await Assert.That(xml).Contains("<c:min val=\"0\"");
        await Assert.That(xml).DoesNotContain("<c:majorGridlines");
        await Assert.That(xml).Contains("<c:legendPos val=\"r\"");

        // Everything on the axis that XLibur does not model is untouched.
        await Assert.That(xml).Contains("<c:majorTickMark val=\"out\"");
        await Assert.That(xml).Contains("<c:tickLblPos val=\"nextTo\"");
        await Assert.That(xml).Contains("<c:crossBetween val=\"between\"");
        await Assert.That(xml).Contains("spcFirstLastPara=\"1\"");
        await Assert.That(xml).Contains("formatCode=\"#,##0\"").Because("The number format was not touched.");
    }

    [Test]
    public async Task EditedAxisAndLegendReloadWithTheNewValues()
    {
        using var original = CreateWorkbookWithExcelShapedChart();
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            chart.CategoryAxis.Title = "Quarter";
            chart.ValueAxis.Orientation = XLAxisOrientation.MaxMin;
            chart.ValueAxis.MajorUnit = 25;
            chart.SecondaryValueAxis.Title = "Price";
            chart.Legend.Visible = false;
            wb.SaveAs(saved, validate: true);
        }

        saved.Position = 0;
        using (var wb = new XLWorkbook(saved))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            await Assert.That(chart.CategoryAxis.Title).IsEqualTo("Quarter");
            await Assert.That(chart.ValueAxis.Orientation).IsEqualTo(XLAxisOrientation.MaxMin);
            await Assert.That(chart.ValueAxis.MajorUnit).IsEqualTo(25);
            await Assert.That(chart.ValueAxis.Title).IsEqualTo("Units").Because("An untouched property is kept.");
            await Assert.That(chart.SecondaryValueAxis.Title).IsEqualTo("Price");
            await Assert.That(chart.Legend.Visible).IsFalse();
        }
    }

    [Test]
    public async Task MarkerAndSmoothAreNotPatchedIntoASeriesTypeThatCannotHoldThem()
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
        var barSeries = xml[xml.IndexOf("<c:barChart>", StringComparison.Ordinal)..xml.IndexOf("<c:lineChart>", StringComparison.Ordinal)];
        await Assert.That(barSeries).DoesNotContain("<c:marker");
        await Assert.That(barSeries).DoesNotContain("<c:smooth");
        await Assert.That(barSeries).Contains("00B050").Because("The fill still applies.");
    }

    /// <summary>
    /// The fixture's line series without its <c>c:marker</c> and <c>c:smooth</c>, so the patcher has
    /// to add them rather than update them.
    /// </summary>
    private static string ChartXmlWithPlainLineSeries()
    {
        var marker = ExcelShapedChartXml.IndexOf("<c:marker>", StringComparison.Ordinal);
        var afterMarker = ExcelShapedChartXml.IndexOf("</c:marker>", StringComparison.Ordinal)
                          + "</c:marker>".Length;
        var withoutMarker = ExcelShapedChartXml[..marker] + ExcelShapedChartXml[afterMarker..];
        return withoutMarker.Replace("<c:smooth val=\"1\"/>", string.Empty);
    }

    [Test]
    public async Task MarkerAndSmoothCanBeAddedToASeriesThatHadNeither()
    {
        using var original = CreateWorkbookWithExcelShapedChart(ChartXmlWithPlainLineSeries());
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            var line = wb.Worksheet("Data").Charts.First().SecondarySeries.Single();
            await Assert.That(line.MarkerStyle).IsEqualTo(XLMarkerStyle.Auto);
            await Assert.That(line.Smooth).IsFalse();

            line.MarkerStyle = XLMarkerStyle.Triangle;
            line.MarkerSize = 9;
            line.MarkerFillColor = XLColor.FromHtml("#FFC000");
            line.Smooth = true;
            wb.SaveAs(saved, validate: true);
        }

        saved.Position = 0;
        using (var wb = new XLWorkbook(saved))
        {
            var line = wb.Worksheet("Data").Charts.First().SecondarySeries.Single();
            await Assert.That(line.MarkerStyle).IsEqualTo(XLMarkerStyle.Triangle);
            await Assert.That(line.MarkerSize).IsEqualTo(9);
            await Assert.That(line.MarkerFillColor).IsEqualTo(XLColor.FromHtml("#FFC000"));
            await Assert.That(line.Smooth).IsTrue();
            await Assert.That(line.LineColor!.ThemeColor).IsEqualTo(XLThemeColor.Accent2).Because("The outline the series came with is untouched.");
        }
    }

    [Test]
    public async Task ClearingEveryMarkerPropertyLeavesNoEmptyMarkerBehind()
    {
        using var original = CreateWorkbookWithExcelShapedChart(ChartXmlWithPlainLineSeries());
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            var line = wb.Worksheet("Data").Charts.First().SecondarySeries.Single();
            line.MarkerSize = null;
            line.MarkerFillColor = null;
            wb.SaveAs(saved, validate: true);
        }

        // An empty c:marker on the series would read back as "this series has markers", which would
        // turn the chart type into LineWithMarkers. The group-level c:marker flag is a different
        // element and is left alone.
        var xml = ReadChartXml(saved);
        var lineSeries = xml[xml.IndexOf("<c:lineChart>", StringComparison.Ordinal)..xml.IndexOf("</c:ser>", xml.IndexOf("<c:lineChart>", StringComparison.Ordinal),
                                 StringComparison.Ordinal)];
        await Assert.That(lineSeries).DoesNotContain("<c:marker");

        saved.Position = 0;
        using (var wb = new XLWorkbook(saved))
        {
            await Assert.That(wb.Worksheet("Data").Charts.First().SecondaryChartType).IsEqualTo(XLChartType.Line);
        }
    }

    [Test]
    public async Task SmoothingCanBeTurnedOffOnASeriesThatHadIt()
    {
        using var original = CreateWorkbookWithExcelShapedChart();
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            var line = wb.Worksheet("Data").Charts.First().SecondarySeries.Single();
            await Assert.That(line.Smooth).IsTrue();
            line.Smooth = false;
            wb.SaveAs(saved, validate: true);
        }

        saved.Position = 0;
        using (var wb = new XLWorkbook(saved))
        {
            await Assert.That(wb.Worksheet("Data").Charts.First().SecondarySeries.Single().Smooth).IsFalse();
        }
    }

    [Test]
    public async Task AnOutlineCanBeAddedToASeriesThatOnlyHadAFill()
    {
        // The bar series' spPr carries a solidFill and an a:ln; strip the a:ln so the patcher has to
        // insert one, in the right place, next to the fill and the effect list.
        var chartXml = ExcelShapedChartXml.Replace(
            "<a:ln w=\"19050\" cap=\"rnd\"><a:solidFill><a:srgbClr val=\"203864\"/></a:solidFill><a:round/></a:ln>",
            string.Empty);

        using var original = CreateWorkbookWithExcelShapedChart(chartXml);
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            var bar = wb.Worksheet("Data").Charts.First().Series.Single();
            await Assert.That(bar.LineColor).IsNull();
            bar.LineColor = XLColor.FromHtml("#203864");
            bar.LineWidthPt = 2;
            wb.SaveAs(saved, validate: true);
        }

        var xml = ReadChartXml(saved);
        await Assert.That(xml).Contains("w=\"25400\"");
        await Assert.That(xml).Contains("<a:effectLst").Because("The effect list is still after the outline.");

        saved.Position = 0;
        using (var wb = new XLWorkbook(saved))
        {
            var bar = wb.Worksheet("Data").Charts.First().Series.Single();
            await Assert.That(bar.LineColor).IsEqualTo(XLColor.FromHtml("#203864"));
            await Assert.That(bar.LineWidthPt).IsEqualTo(2);
            await Assert.That(bar.FillColor).IsEqualTo(XLColor.FromHtml("#ED7D31")).Because("The fill it came with is untouched.");
        }
    }

    [Test]
    public async Task ClearingTheOutlineWidthLeavesTheColourAlone()
    {
        using var original = CreateWorkbookWithExcelShapedChart();
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            var bar = wb.Worksheet("Data").Charts.First().Series.Single();
            bar.LineWidthPt = null;
            wb.SaveAs(saved, validate: true);
        }

        saved.Position = 0;
        using (var wb = new XLWorkbook(saved))
        {
            var bar = wb.Worksheet("Data").Charts.First().Series.Single();
            await Assert.That(bar.LineWidthPt).IsNull();
            await Assert.That(bar.LineColor).IsEqualTo(XLColor.FromHtml("#203864"));
        }
    }

    [Test]
    public async Task EveryMarkerStyleRoundTrips()
    {
        XLMarkerStyle[] styles =
        [
            XLMarkerStyle.None, XLMarkerStyle.Circle, XLMarkerStyle.Dash, XLMarkerStyle.Diamond,
            XLMarkerStyle.Dot, XLMarkerStyle.Plus, XLMarkerStyle.Square, XLMarkerStyle.Star,
            XLMarkerStyle.Triangle, XLMarkerStyle.X
        ];

        foreach (var style in styles)
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Data");
                ws.Cell("A1").Value = "Q1";
                ws.Cell("B1").Value = 100;

                var chart = ws.Charts.Add(XLChartType.LineWithMarkers);
                chart.Series.Add("Sales", "Data!$B$1:$B$1", "Data!$A$1:$A$1").MarkerStyle = style;
                chart.Position.SetColumn(5).SetRow(1);
                chart.SecondPosition.SetColumn(12).SetRow(15);
                wb.SaveAs(ms, validate: true);
            }

            ms.Position = 0;
            using (var wb = new XLWorkbook(ms))
            {
                await Assert.That(wb.Worksheet("Data").Charts.First().Series.Single().MarkerStyle).IsEqualTo(style).Because($"{style} did not round-trip.");
            }
        }
    }

    // ── The chart title ─────────────────────────────────────────────────

    /// <summary>
    /// A chart part whose <c>c:chart</c> opens with the given head — the <c>c:title</c> block and its
    /// <c>c:autoTitleDeleted</c>, which is what the title tests vary — over a plain bar series.
    /// </summary>
    private static string ChartXmlWithTitleHead(string head) => $"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <c:chart>
            {head}
            <c:plotArea>
              <c:layout/>
              <c:barChart>
                <c:barDir val="col"/>
                <c:grouping val="clustered"/>
                <c:varyColors val="0"/>
                <c:ser>
                  <c:idx val="0"/>
                  <c:order val="0"/>
                  <c:tx><c:v>Units</c:v></c:tx>
                  <c:cat><c:strRef><c:f>Data!$A$1:$A$2</c:f></c:strRef></c:cat>
                  <c:val><c:numRef><c:f>Data!$B$1:$B$2</c:f></c:numRef></c:val>
                </c:ser>
                <c:axId val="111111111"/>
                <c:axId val="222222222"/>
              </c:barChart>
              <c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:catAx>
              <c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111111111"/></c:valAx>
            </c:plotArea>
            <c:plotVisOnly val="1"/>
          </c:chart>
        </c:chartSpace>
        """;

    [Test]
    public async Task RetitlingALoadedChartIsWrittenToTheChartPart()
    {
        using var original = CreateWorkbookWithExcelShapedChart();
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            await Assert.That(chart.Title).IsEqualTo("Units and price");

            chart.SetTitle("Edited title");

            // A property that has always been patched, edited in the same save.
            chart.Legend.Position = XLLegendPosition.Top;
            wb.SaveAs(saved, validate: true);
        }

        var xml = ReadChartXml(saved);
        await Assert.That(xml).Contains("<a:t>Edited title</a:t>");
        await Assert.That(xml).DoesNotContain("Units and price");
        await Assert.That(xml).Contains("<c:legendPos val=\"t\"");
        await Assert.That(xml).Contains("<a:t>Units</a:t>").Because("The value axis title is a different title.");
    }

    [Test]
    public async Task AnEditedTitleReloadsWithTheNewValue()
    {
        using var original = CreateWorkbookWithExcelShapedChart();
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            wb.Worksheet("Data").Charts.First().SetTitle("Edited title");
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using (var wb = new XLWorkbook(saved))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            await Assert.That(chart.Title).IsEqualTo("Edited title");
            await Assert.That(chart.ValueAxis.Title).IsEqualTo("Units").Because("The axis title is untouched.");
        }
    }

    [Test]
    public async Task RetitlingALoadedChartKeepsTheFormattingTheTitleCameWith()
    {
        const string styledTitle = """
            <c:title>
              <c:tx><c:rich><a:bodyPr rot="0" vert="horz"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1800"/></a:pPr><a:r><a:rPr lang="en-US" sz="1600" b="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Original title</a:t></a:r></a:p></c:rich></c:tx>
              <c:layout/>
              <c:overlay val="0"/>
              <c:spPr><a:noFill/></c:spPr>
            </c:title>
            <c:autoTitleDeleted val="0"/>
            """;

        using var original = CreateWorkbookWithExcelShapedChart(ChartXmlWithTitleHead(styledTitle));
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            wb.Worksheet("Data").Charts.First().SetTitle("Edited title");
            wb.SaveAs(saved, validate: true);
        }

        var xml = ReadChartXml(saved);
        await Assert.That(xml).Contains("<a:t>Edited title</a:t>");
        await Assert.That(xml).DoesNotContain("Original title");

        // Everything around the text survived.
        await Assert.That(xml).Contains("sz=\"1600\"").Because("The run properties of the old run are kept.");
        await Assert.That(xml).Contains("b=\"1\"");
        await Assert.That(xml).Contains("FF0000");
        await Assert.That(xml).Contains("<a:defRPr sz=\"1800\"");
        await Assert.That(xml).Contains("rot=\"0\"");
        await Assert.That(xml).Contains("<c:overlay val=\"0\"");
        await Assert.That(xml).Contains("<a:noFill");
    }

    [Test]
    public async Task ATitleCanBeAddedToALoadedChartThatHadNone()
    {
        using var original = CreateWorkbookWithExcelShapedChart(
            ChartXmlWithTitleHead("<c:autoTitleDeleted val=\"1\"/>"));
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            await Assert.That(chart.Title).IsNull();

            chart.SetTitle("New title");
            wb.SaveAs(saved, validate: true);
        }

        var xml = ReadChartXml(saved);
        await Assert.That(xml).Contains("<a:t>New title</a:t>");
        await Assert.That(xml).Contains("<c:autoTitleDeleted val=\"0\"").Because("Excel hides the title while this is set.");
        await Assert.That(xml.IndexOf("<c:title>", StringComparison.Ordinal))
            .IsLessThan(xml.IndexOf("<c:autoTitleDeleted", StringComparison.Ordinal))
            .Because("c:title precedes c:autoTitleDeleted in CT_Chart.");
    }

    [Test]
    public async Task ClearingTheTitleOfALoadedChartRemovesIt()
    {
        using var original = CreateWorkbookWithExcelShapedChart();
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            wb.Worksheet("Data").Charts.First().SetTitle(null);
            wb.SaveAs(saved, validate: true);
        }

        var xml = ReadChartXml(saved);
        await Assert.That(xml).DoesNotContain("Units and price");
        await Assert.That(xml).Contains("<c:autoTitleDeleted val=\"1\"").Because("Excel would otherwise put an automatic title back.");
        await Assert.That(xml).Contains("<a:t>Units</a:t>").Because("The value axis keeps its own title.");

        saved.Position = 0;
        using (var wb = new XLWorkbook(saved))
        {
            await Assert.That(wb.Worksheet("Data").Charts.First().Title).IsNull();
        }
    }

    [Test]
    public async Task ATitleLinkedToACellIsReplacedByTheLiteralOne()
    {
        const string linkedTitle = """
            <c:title>
              <c:tx><c:strRef><c:f>Data!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:tx>
              <c:overlay val="0"/>
            </c:title>
            <c:autoTitleDeleted val="0"/>
            """;

        using var original = CreateWorkbookWithExcelShapedChart(ChartXmlWithTitleHead(linkedTitle));
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            wb.Worksheet("Data").Charts.First().SetTitle("Literal title");
            wb.SaveAs(saved, validate: true);
        }

        var xml = ReadChartXml(saved);
        await Assert.That(xml).Contains("<a:t>Literal title</a:t>");
        await Assert.That(xml).DoesNotContain("<c:f>Data!$A$1</c:f>").Because("The link the title had is gone with the text it produced.");
        await Assert.That(xml).Contains("<c:overlay val=\"0\"").Because("The rest of c:title is left alone.");
    }

    /// <summary>
    /// A 3D chart part with no title of its own. <c>c:view3D</c>, <c>c:floor</c>, <c>c:sideWall</c>
    /// and <c>c:backWall</c> sit between <c>c:autoTitleDeleted</c> and <c>c:plotArea</c> in
    /// <c>CT_Chart</c>, so they are what a newly inserted title has to be placed ahead of.
    /// </summary>
    private const string ThreeDimensionalChartXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <c:chart>
            <c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:rAngAx val="1"/></c:view3D>
            <c:floor><c:thickness val="0"/></c:floor>
            <c:sideWall><c:thickness val="0"/></c:sideWall>
            <c:backWall><c:thickness val="0"/></c:backWall>
            <c:plotArea>
              <c:layout/>
              <c:bar3DChart>
                <c:barDir val="col"/>
                <c:grouping val="clustered"/>
                <c:varyColors val="0"/>
                <c:ser>
                  <c:idx val="0"/>
                  <c:order val="0"/>
                  <c:tx><c:v>Units</c:v></c:tx>
                  <c:cat><c:strRef><c:f>Data!$A$1:$A$2</c:f></c:strRef></c:cat>
                  <c:val><c:numRef><c:f>Data!$B$1:$B$2</c:f></c:numRef></c:val>
                </c:ser>
                <c:axId val="111111111"/>
                <c:axId val="222222222"/>
              </c:bar3DChart>
              <c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:catAx>
              <c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111111111"/></c:valAx>
            </c:plotArea>
            <c:plotVisOnly val="1"/>
          </c:chart>
        </c:chartSpace>
        """;

    [Test]
    public async Task ATitleAddedToA3DChartGoesAheadOfTheView3DElements()
    {
        using var original = CreateWorkbookWithExcelShapedChart(ThreeDimensionalChartXml);
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            await Assert.That(chart.Title).IsNull();

            chart.SetTitle("3D title");
            wb.SaveAs(saved, validate: true);
        }

        var xml = ReadChartXml(saved);
        await Assert.That(xml).Contains("<a:t>3D title</a:t>");

        // CT_Chart orders title, autoTitleDeleted, pivotFmts, view3D, floor, sideWall, backWall,
        // plotArea — so both inserted elements belong ahead of the 3D block, not merely ahead of
        // the plot area.
        var title = xml.IndexOf("<c:title>", StringComparison.Ordinal);
        var autoTitleDeleted = xml.IndexOf("<c:autoTitleDeleted", StringComparison.Ordinal);
        var view3D = xml.IndexOf("<c:view3D", StringComparison.Ordinal);

        await Assert.That(title).IsLessThan(autoTitleDeleted);
        await Assert.That(autoTitleDeleted).IsLessThan(view3D);
        await Assert.That(view3D).IsLessThan(xml.IndexOf("<c:floor", StringComparison.Ordinal)).Because("The 3D block itself is left in its own order.");

        saved.Position = 0;
        using (var wb = new XLWorkbook(saved))
        {
            await Assert.That(wb.Worksheet("Data").Charts.First().Title).IsEqualTo("3D title");
        }
    }

    // ── The title of an extended (ChartEx) chart ────────────────────────

    /// <summary>
    /// A workbook holding one waterfall chart, which XLibur writes as an extended chart part. An
    /// empty <paramref name="title"/> leaves the chart untitled.
    /// </summary>
    private static MemoryStream CreateWorkbookWithWaterfallChart(string title)
    {
        var ms = new MemoryStream();
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = "Start";
        ws.Cell("A2").Value = "Add";
        ws.Cell("B1").Value = 1000;
        ws.Cell("B2").Value = 500;

        var chart = ws.Charts.Add(XLChartType.Waterfall);
        if (title.Length != 0)
            chart.SetTitle(title);
        chart.Series.Add("Amount", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
        chart.Position.SetColumn(5).SetRow(1);
        chart.SecondPosition.SetColumn(12).SetRow(15);
        wb.SaveAs(ms);

        ms.Position = 0;
        return ms;
    }

    private static string ReadExtendedChartXml(MemoryStream stream)
    {
        stream.Position = 0;
        using var doc = SpreadsheetDocument.Open(stream, false);
        var chartPart = doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!
            .ExtendedChartParts.First();
        using var reader = new StreamReader(chartPart.GetStream(FileMode.Open, FileAccess.Read));
        return reader.ReadToEnd();
    }

    [Test]
    public async Task RetitlingALoadedExtendedChartIsWrittenToTheChartPart()
    {
        using var original = CreateWorkbookWithWaterfallChart("Original title");
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.Waterfall);
            await Assert.That(chart.Title).IsEqualTo("Original title");

            chart.SetTitle("Edited title");
            wb.SaveAs(saved);
        }

        var xml = ReadExtendedChartXml(saved);
        await Assert.That(xml).Contains("<a:t>Edited title</a:t>");
        await Assert.That(xml).DoesNotContain("Original title");
        await Assert.That(xml).Contains("<cx:plotAreaRegion>").Because("The rest of the chart part is untouched.");

        saved.Position = 0;
        using (var wb = new XLWorkbook(saved))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            await Assert.That(chart.Title).IsEqualTo("Edited title");
            await Assert.That(chart.Series.Count).IsEqualTo(1);
        }
    }

    [Test]
    public async Task ATitleCanBeAddedToALoadedExtendedChartThatHadNone()
    {
        using var original = CreateWorkbookWithWaterfallChart(string.Empty);
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            await Assert.That(chart.Title).IsNull();

            chart.SetTitle("New title");
            wb.SaveAs(saved);
        }

        var xml = ReadExtendedChartXml(saved);
        await Assert.That(xml).Contains("<a:t>New title</a:t>");
        await Assert.That(xml.IndexOf("<cx:title", StringComparison.Ordinal))
            .IsLessThan(xml.IndexOf("<cx:plotArea>", StringComparison.Ordinal))
            .Because("cx:title opens CT_Chart, ahead of cx:plotArea.");

        saved.Position = 0;
        using (var wb = new XLWorkbook(saved))
        {
            await Assert.That(wb.Worksheet("Data").Charts.First().Title).IsEqualTo("New title");
        }
    }

    [Test]
    public async Task ClearingTheTitleOfALoadedExtendedChartRemovesIt()
    {
        using var original = CreateWorkbookWithWaterfallChart("Original title");
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            wb.Worksheet("Data").Charts.First().SetTitle(null);
            wb.SaveAs(saved);
        }

        var xml = ReadExtendedChartXml(saved);
        await Assert.That(xml).DoesNotContain("Original title");
        await Assert.That(xml).DoesNotContain("<cx:title");

        saved.Position = 0;
        using (var wb = new XLWorkbook(saved))
        {
            await Assert.That(wb.Worksheet("Data").Charts.First().Title).IsNull();
        }
    }

    [Test]
    public async Task ReadingATitleIsNotAnEdit()
    {
        using var original = CreateWorkbookWithExcelShapedChart();
        var before = ReadChartXml(original);

        using var saved = new MemoryStream();
        original.Position = 0;
        using (var wb = new XLWorkbook(original))
        {
            var chart = wb.Worksheet("Data").Charts.First();
            await Assert.That(chart.Title).IsEqualTo("Units and price");
            wb.SaveAs(saved);
        }

        await Assert.That(ReadChartXml(saved)).IsEqualTo(before);
    }

    [Test]
    public async Task SavingANewChartTwiceDoesNotDuplicateIt()
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
            await Assert.That(ws.Charts.Count).IsEqualTo(1);
            await Assert.That(ws.Charts.First().Series.Single().FillColor).IsEqualTo(XLColor.FromHtml("#00B050"));
        }
    }
}
