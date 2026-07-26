using NUnit.Framework;
using System.Linq;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Charts;

/// <summary>
/// The chart group elements Excel writes that XLibur does not write itself: the 3D variants and
/// pie-of-pie / bar-of-pie. They used to read as a chart with no series at all.
/// </summary>
[TestFixture]
public class ChartGroupKindReaderTests
{
    private const string CategoryAndValueAxes = ChartPartFixture.CategoryAndValueAxes;

    private static string CategoryAndValueSeries(string name) =>
        ChartPartFixture.CategoryAndValueSeries(name);

    private static IXLChart LoadChart(string plotAreaBody, out XLWorkbook workbook) =>
        ChartPartFixture.LoadChart(plotAreaBody, out workbook);

    [Test]
    public void Pie3DChartIsRead()
    {
        var chart = LoadChart($"<c:pie3DChart><c:varyColors val=\"1\"/>{CategoryAndValueSeries("Share")}</c:pie3DChart>",
            out var wb);
        using (wb)
        {
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.Pie3D));
            Assert.That(chart.Series.Single().Name, Is.EqualTo("Share"));
            Assert.That(chart.Series.Single().ValueReferences, Is.EqualTo("Data!$B$1:$B$2"));
        }
    }

    [Test]
    public void PieOfPieAndBarOfPieAreToldApart()
    {
        var pieOfPie = LoadChart(
            $"<c:ofPieChart><c:ofPieType val=\"pie\"/>{CategoryAndValueSeries("Share")}</c:ofPieChart>", out var wb1);
        using (wb1)
        {
            Assert.That(pieOfPie.ChartType, Is.EqualTo(XLChartType.PieToPie));
        }

        var barOfPie = LoadChart(
            $"<c:ofPieChart><c:ofPieType val=\"bar\"/>{CategoryAndValueSeries("Share")}</c:ofPieChart>", out var wb2);
        using (wb2)
        {
            Assert.That(barOfPie.ChartType, Is.EqualTo(XLChartType.PieToBar));
        }
    }

    [Test]
    public void Line3DChartIsRead()
    {
        var chart = LoadChart(
            $"<c:line3DChart><c:grouping val=\"standard\"/>{CategoryAndValueSeries("Trend")}<c:axId val=\"1\"/><c:axId val=\"2\"/></c:line3DChart>{CategoryAndValueAxes}",
            out var wb);
        using (wb)
        {
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.Line3D));
            Assert.That(chart.Series.Single().Name, Is.EqualTo("Trend"));
        }
    }

    [TestCase("standard", XLChartType.Area3D)]
    [TestCase("stacked", XLChartType.AreaStacked3D)]
    [TestCase("percentStacked", XLChartType.AreaStacked100Percent3D)]
    public void Area3DGroupingIsRead(string grouping, XLChartType expected)
    {
        var chart = LoadChart(
            $"<c:area3DChart><c:grouping val=\"{grouping}\"/>{CategoryAndValueSeries("Area")}<c:axId val=\"1\"/><c:axId val=\"2\"/></c:area3DChart>{CategoryAndValueAxes}",
            out var wb);
        using (wb)
        {
            Assert.That(chart.ChartType, Is.EqualTo(expected));
            Assert.That(chart.Series, Has.Exactly(1).Items);
        }
    }

    [Test]
    public void Surface3DChartIsRead()
    {
        var chart = LoadChart(
            $"<c:surface3DChart>{CategoryAndValueSeries("Height")}<c:axId val=\"1\"/><c:axId val=\"2\"/><c:axId val=\"3\"/></c:surface3DChart>{CategoryAndValueAxes}<c:serAx><c:axId val=\"3\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:crossAx val=\"2\"/></c:serAx>",
            out var wb);
        using (wb)
        {
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.Surface));
            Assert.That(chart.Series.Single().Name, Is.EqualTo("Height"));
        }
    }

    [Test]
    public void A3DBarChartStillReadsAsBefore()
    {
        var chart = LoadChart(
            $"<c:bar3DChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>{CategoryAndValueSeries("Sales")}<c:shape val=\"box\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:bar3DChart>{CategoryAndValueAxes}",
            out var wb);
        using (wb)
        {
            Assert.That(chart.ChartType, Is.EqualTo(XLChartType.ColumnClustered3D));
        }
    }

    [Test]
    public void SeriesFormattingOnA3DGroupIsRead()
    {
        var series = """
            <c:ser>
              <c:idx val="0"/>
              <c:order val="0"/>
              <c:tx><c:v>Trend</c:v></c:tx>
              <c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></a:ln></c:spPr>
              <c:marker><c:symbol val="square"/><c:size val="6"/></c:marker>
              <c:cat><c:strRef><c:f>Data!$A$1:$A$2</c:f></c:strRef></c:cat>
              <c:val><c:numRef><c:f>Data!$B$1:$B$2</c:f></c:numRef></c:val>
              <c:smooth val="1"/>
            </c:ser>
            """;

        var chart = LoadChart(
            $"<c:line3DChart><c:grouping val=\"standard\"/>{series}<c:axId val=\"1\"/><c:axId val=\"2\"/></c:line3DChart>{CategoryAndValueAxes}",
            out var wb);
        using (wb)
        {
            var s = chart.Series.Single();
            Assert.That(s.LineColor, Is.EqualTo(XLColor.FromHtml("#ED7D31")));
            Assert.That(s.LineWidthPt, Is.EqualTo(1.5));
            Assert.That(s.MarkerStyle, Is.EqualTo(XLMarkerStyle.Square));
            Assert.That(s.MarkerSize, Is.EqualTo(6));
            Assert.That(s.Smooth, Is.True);
        }
    }
}
