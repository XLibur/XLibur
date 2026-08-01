using System.Linq;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Charts;

/// <summary>
/// The chart group elements Excel writes that XLibur does not write itself: the 3D variants and
/// pie-of-pie / bar-of-pie. They used to read as a chart with no series at all.
/// </summary>
public class ChartGroupKindReaderTests
{
    private const string CategoryAndValueAxes = ChartPartFixture.CategoryAndValueAxes;

    private static string CategoryAndValueSeries(string name) =>
        ChartPartFixture.CategoryAndValueSeries(name);

    private static IXLChart LoadChart(string plotAreaBody, out XLWorkbook workbook) =>
        ChartPartFixture.LoadChart(plotAreaBody, out workbook);

    [Test]
    public async Task Pie3DChartIsRead()
    {
        var chart = LoadChart($"<c:pie3DChart><c:varyColors val=\"1\"/>{CategoryAndValueSeries("Share")}</c:pie3DChart>",
            out var wb);
        using (wb)
        {
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.Pie3D);
            await Assert.That(chart.Series.Single().Name).IsEqualTo("Share");
            await Assert.That(chart.Series.Single().ValueReferences).IsEqualTo("Data!$B$1:$B$2");
        }
    }

    [Test]
    public async Task PieOfPieAndBarOfPieAreToldApart()
    {
        var pieOfPie = LoadChart(
            $"<c:ofPieChart><c:ofPieType val=\"pie\"/>{CategoryAndValueSeries("Share")}</c:ofPieChart>", out var wb1);
        using (wb1)
        {
            await Assert.That(pieOfPie.ChartType).IsEqualTo(XLChartType.PieToPie);
        }

        var barOfPie = LoadChart(
            $"<c:ofPieChart><c:ofPieType val=\"bar\"/>{CategoryAndValueSeries("Share")}</c:ofPieChart>", out var wb2);
        using (wb2)
        {
            await Assert.That(barOfPie.ChartType).IsEqualTo(XLChartType.PieToBar);
        }
    }

    [Test]
    public async Task Line3DChartIsRead()
    {
        var chart = LoadChart(
            $"<c:line3DChart><c:grouping val=\"standard\"/>{CategoryAndValueSeries("Trend")}<c:axId val=\"1\"/><c:axId val=\"2\"/></c:line3DChart>{CategoryAndValueAxes}",
            out var wb);
        using (wb)
        {
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.Line3D);
            await Assert.That(chart.Series.Single().Name).IsEqualTo("Trend");
        }
    }

    [Test]
    [Arguments("standard", XLChartType.Area3D)]
    [Arguments("stacked", XLChartType.AreaStacked3D)]
    [Arguments("percentStacked", XLChartType.AreaStacked100Percent3D)]
    public async Task Area3DGroupingIsRead(string grouping, XLChartType expected)
    {
        var chart = LoadChart(
            $"<c:area3DChart><c:grouping val=\"{grouping}\"/>{CategoryAndValueSeries("Area")}<c:axId val=\"1\"/><c:axId val=\"2\"/></c:area3DChart>{CategoryAndValueAxes}",
            out var wb);
        using (wb)
        {
            await Assert.That(chart.ChartType).IsEqualTo(expected);
            await Assert.That(chart.Series.Count).IsEqualTo(1);
        }
    }

    [Test]
    public async Task Surface3DChartIsRead()
    {
        var chart = LoadChart(
            $"<c:surface3DChart>{CategoryAndValueSeries("Height")}<c:axId val=\"1\"/><c:axId val=\"2\"/><c:axId val=\"3\"/></c:surface3DChart>{CategoryAndValueAxes}<c:serAx><c:axId val=\"3\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:crossAx val=\"2\"/></c:serAx>",
            out var wb);
        using (wb)
        {
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.Surface);
            await Assert.That(chart.Series.Single().Name).IsEqualTo("Height");
        }
    }

    [Test]
    public async Task A3DBarChartStillReadsAsBefore()
    {
        var chart = LoadChart(
            $"<c:bar3DChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>{CategoryAndValueSeries("Sales")}<c:shape val=\"box\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:bar3DChart>{CategoryAndValueAxes}",
            out var wb);
        using (wb)
        {
            await Assert.That(chart.ChartType).IsEqualTo(XLChartType.ColumnClustered3D);
        }
    }

    [Test]
    public async Task SeriesFormattingOnA3DGroupIsRead()
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
            await Assert.That(s.LineColor).IsEqualTo(XLColor.FromHtml("#ED7D31"));
            await Assert.That(s.LineWidthPt).IsEqualTo(1.5);
            await Assert.That(s.MarkerStyle).IsEqualTo(XLMarkerStyle.Square);
            await Assert.That(s.MarkerSize).IsEqualTo(6);
            await Assert.That(s.Smooth).IsTrue();
        }
    }
}
