using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Linq;
using System.Text;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Charts;

/// <summary>
/// Builds a workbook whose chart part holds hand-written XML. XLibur's own writer only produces the
/// shapes it knows how to produce, so anything Excel may legitimately emit but XLibur never writes —
/// 3D groups, an unusual group order — has to be fed to the reader this way.
/// </summary>
internal static class ChartPartFixture
{
    /// <summary>A <c>c:catAx</c>/<c>c:valAx</c> pair with ids 1 and 2.</summary>
    internal const string CategoryAndValueAxes = """
        <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
        <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
        """;

    /// <summary>A <c>c:ser</c> over the <c>Data</c> sheet the fixture fills in.</summary>
    internal static string CategoryAndValueSeries(string name) => $"""
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:v>{name}</c:v></c:tx>
          <c:cat><c:strRef><c:f>Data!$A$1:$A$2</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Data!$B$1:$B$2</c:f></c:numRef></c:val>
        </c:ser>
        """;

    /// <summary>
    /// Wraps a plot area body in the rest of a chart part, and puts it in a workbook.
    /// </summary>
    internal static MemoryStream CreateWorkbookWithPlotArea(string plotAreaBody)
    {
        var chartXml = $"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <c:chart>
                <c:plotArea>
                  <c:layout/>
                  {plotAreaBody}
                </c:plotArea>
                <c:plotVisOnly val="1"/>
              </c:chart>
            </c:chartSpace>
            """;

        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Q1";
            ws.Cell("A2").Value = "Q2";
            ws.Cell("B1").Value = 100;
            ws.Cell("B2").Value = 200;

            var chart = ws.Charts.Add(XLChartType.ColumnClustered);
            chart.Series.Add("Placeholder", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
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

    /// <summary>
    /// Loads the chart of <see cref="CreateWorkbookWithPlotArea"/>. The caller owns
    /// <paramref name="workbook"/>, which has to outlive the chart it hands back.
    /// </summary>
    internal static IXLChart LoadChart(string plotAreaBody, out XLWorkbook workbook)
    {
        var ms = CreateWorkbookWithPlotArea(plotAreaBody);
        workbook = new XLWorkbook(ms);
        return workbook.Worksheet("Data").Charts.Single();
    }
}
