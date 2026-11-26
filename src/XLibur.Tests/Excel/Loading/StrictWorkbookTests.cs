using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Loading;

public class StrictWorkbookTests
{
    private const string StrictWorkbook = @"TryToLoad\LO\xlsx\universal-content-strict.xlsx";

    private static readonly XNamespace TransitionalMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private static readonly XNamespace TransitionalDrawingMain = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static readonly XNamespace TransitionalChart = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    [Test]
    public async Task The_cells_of_a_strict_workbook_load()
    {
        using var wb = TestHelper.LoadWorkbook(StrictWorkbook);

        await AssertSheet1Cells(wb.Worksheet("Sheet1"));
    }

    [Test]
    public async Task The_cells_of_a_strict_workbook_survive_a_save()
    {
        using var saved = TestHelper.LoadAndSave(StrictWorkbook);

        var cellsWithValue = SavedCells(saved, "xl/worksheets/sheet1.xml")
            .Where(c => c.Element(TransitionalMain + "v") is not null)
            .Select(c => c.Attribute("r")!.Value);
        await Assert.That(cellsWithValue).IsEquivalentTo(["A1", "B1", "C1", "A2", "B2", "C2", "C3", "C4"]);

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await AssertSheet1Cells(reloaded.Worksheet("Sheet1"));
    }

    [Test]
    public async Task The_theme_colors_of_a_strict_workbook_load()
    {
        using var wb = TestHelper.LoadWorkbook(StrictWorkbook);

        await Assert.That(wb.Theme.Accent1.Color.ToArgb()).IsEqualTo(unchecked((int)0xFF5B9BD5));
        await Assert.That(wb.Theme.Accent5.Color.ToArgb()).IsEqualTo(unchecked((int)0xFF4472C4));
        await Assert.That(wb.Theme.Text2.Color.ToArgb()).IsEqualTo(unchecked((int)0xFF44546A));
    }

    /// <summary>
    /// A Strict DrawingML percentage is written <c>"120%"</c>; the Transitional namespace the save
    /// writes requires thousandths of a percent, <c>"120000"</c>.
    /// </summary>
    [Test]
    [Arguments("xl/theme/theme1.xml")]
    [Arguments("xl/drawings/drawing1.xml")]
    public async Task A_saved_strict_drawingml_part_holds_transitional_percentages(string partName)
    {
        using var saved = TestHelper.LoadAndSave(StrictWorkbook);

        var percentages = SavedPart(saved, partName).Descendants()
            .Where(e => e.Name.Namespace == TransitionalDrawingMain)
            .SelectMany(e => e.Attributes())
            .Where(a => a.Value.EndsWith('%'))
            .Select(a => $"{a.Parent!.Name.LocalName}/@{a.Name.LocalName}={a.Value}");
        await Assert.That(percentages).IsEmpty();

        saved.Position = 0;
        using var document = SpreadsheetDocument.Open(saved, false);
        var part = document.WorkbookPart!.GetPartsOfType<ThemePart>()
            .Concat<OpenXmlPart>(document.WorkbookPart.WorksheetParts.Select(w => w.DrawingsPart).OfType<OpenXmlPart>())
            .Single(p => p.Uri.OriginalString == "/" + partName);
        var errors = new OpenXmlValidator().Validate(part).Select(e => $"{e.Path?.XPath}: {e.Description}");
        await Assert.That(errors).IsEmpty();
    }

    [Test]
    public async Task A_saved_strict_percentage_keeps_its_value()
    {
        using var saved = TestHelper.LoadAndSave(StrictWorkbook);

        var stops = SavedPart(saved, "xl/theme/theme1.xml").Descendants(TransitionalDrawingMain + "gs")
            .Select(gs => gs.Attribute("pos")!.Value).Distinct();
        await Assert.That(stops).IsEquivalentTo(["0", "50000", "100000"]);

        var miter = SavedPart(saved, "xl/drawings/drawing1.xml").Descendants(TransitionalDrawingMain + "miter").Single();
        await Assert.That(miter.Attribute("lim")!.Value).IsEqualTo("800000");
    }

    /// <summary>
    /// A Strict chart percentage is written <c>"150%"</c>; the Transitional chart namespace the save
    /// writes requires the whole number, <c>"150"</c>. A text attribute that ends in <c>%</c>, such as a
    /// number format code, keeps its value.
    /// </summary>
    [Test]
    public async Task A_saved_strict_chart_holds_transitional_percentages()
    {
        using var source = StrictWorkbookWithCharts();
        using var wb = new XLWorkbook(source);
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        await Assert.That(ChartValues(saved, "xl/charts/chart1.xml"))
            .IsEqualTo("hPercent=80 depthPercent=120 thickness=0 gapWidth=150 overlap=-25 lblOffset=100 numFmt/formatCode=0%");
        await Assert.That(ChartValues(saved, "xl/charts/chart2.xml")).IsEqualTo("holeSize=50");

        saved.Position = 0;
        using var document = SpreadsheetDocument.Open(saved, false);
        var errors = document.WorkbookPart!.WorksheetParts
            .SelectMany(w => w.DrawingsPart?.ChartParts ?? [])
            .SelectMany(p => new OpenXmlValidator().Validate(p))
            .Select(e => $"{e.Part?.Uri} {e.Path?.XPath}: {e.Description}");
        await Assert.That(errors).IsEmpty();

        static string ChartValues(MemoryStream package, string partName) => string.Join(" ",
            SavedPart(package, partName).Descendants()
                .Where(e => e.Name.Namespace == TransitionalChart)
                .SelectMany(e => e.Attributes())
                .Where(a => a.Parent!.Name.LocalName is "hPercent" or "depthPercent" or "thickness" or "gapWidth"
                                or "overlap" or "lblOffset" or "holeSize" ||
                            a.Value.EndsWith('%'))
                .Select(a => a.Name.LocalName == "val"
                    ? $"{a.Parent!.Name.LocalName}={a.Value}"
                    : $"{a.Parent!.Name.LocalName}/{a.Name.LocalName}={a.Value}"));
    }

    /// <summary>
    /// The Strict fixture with a bar chart and a doughnut chart on Sheet2, written in the Strict chart
    /// namespace as Excel writes a Strict workbook. No fixture holds a Strict chart.
    /// </summary>
    private static MemoryStream StrictWorkbookWithCharts()
    {
        const string strictNamespaces =
            """xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart" xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" """;
        const string series =
            "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:val><c:numRef><c:f>Sheet1!$A$1:$A$2</c:f></c:numRef></c:val></c:ser>";
        const string barChart =
            $"""<c:chartSpace {strictNamespaces}><c:chart>""" +
            """<c:view3D><c:hPercent val="80%"/><c:depthPercent val="120%"/></c:view3D><c:floor><c:thickness val="0%"/></c:floor>""" +
            $"""<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>{series}""" +
            """<c:gapWidth val="150%"/><c:overlap val="-25%"/><c:axId val="1"/><c:axId val="2"/></c:barChart>""" +
            """<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>""" +
            """<c:axPos val="b"/><c:crossAx val="2"/><c:lblOffset val="100%"/></c:catAx>""" +
            """<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>""" +
            """<c:axPos val="l"/><c:numFmt formatCode="0%" sourceLinked="0"/><c:crossAx val="1"/></c:valAx>""" +
            """</c:plotArea></c:chart></c:chartSpace>""";
        const string doughnutChart =
            $"""<c:chartSpace {strictNamespaces}><c:chart><c:plotArea><c:layout/>""" +
            $"""<c:doughnutChart><c:varyColors val="1"/>{series}<c:firstSliceAng val="0"/><c:holeSize val="50%"/></c:doughnutChart>""" +
            """</c:plotArea></c:chart></c:chartSpace>""";

        var drawing = new StringBuilder(
            $"""<xdr:wsDr xmlns:xdr="http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing" {strictNamespaces}>""");
        for (var chart = 1; chart <= 2; chart++)
        {
            drawing.Append(
                $"""<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{chart * 20}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>""" +
                $"""<xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{chart * 20 + 15}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>""" +
                $"""<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{chart + 1}" name="Chart {chart}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>""" +
                """<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://purl.oclc.org/ooxml/drawingml/chart">""" +
                $"""<c:chart r:id="rId{chart}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>""");
        }
        drawing.Append("</xdr:wsDr>");

        var package = new MemoryStream();
        using (var fixture = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(StrictWorkbook)))
            fixture.CopyTo(package);

        using (var zip = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true))
        {
            Replace("[Content_Types].xml", xml => xml.Replace("</Types>",
                """<Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>""" +
                """<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>""" +
                """<Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>"""));
            Replace("xl/worksheets/sheet2.xml", xml => xml.Replace("</worksheet>", """<drawing r:id="rId1"/></worksheet>"""));
            Write("xl/worksheets/_rels/sheet2.xml.rels", Relationships(
                ("rId1", "drawing", "../drawings/drawing2.xml")));
            Write("xl/drawings/drawing2.xml", drawing.ToString());
            Write("xl/drawings/_rels/drawing2.xml.rels", Relationships(
                ("rId1", "chart", "../charts/chart1.xml"), ("rId2", "chart", "../charts/chart2.xml")));
            Write("xl/charts/chart1.xml", barChart);
            Write("xl/charts/chart2.xml", doughnutChart);

            void Replace(string name, Func<string, string> edit)
            {
                var entry = zip.GetEntry(name)!;
                string xml;
                using (var reader = new StreamReader(entry.Open()))
                    xml = reader.ReadToEnd();
                entry.Delete();
                Write(name, edit(xml));
            }

            void Write(string name, string xml)
            {
                using var writer = new StreamWriter(zip.CreateEntry(name).Open());
                writer.Write(xml);
            }
        }

        package.Position = 0;
        return package;

        static string Relationships(params (string Id, string Type, string Target)[] relationships) =>
            """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">""" +
            string.Concat(relationships.Select(r =>
                $"""<Relationship Id="{r.Id}" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/{r.Type}" Target="{r.Target}"/>""")) +
            "</Relationships>";
    }

    private static async Task AssertSheet1Cells(IXLWorksheet ws)
    {
        await Assert.That(ws.CellsUsed(XLCellsUsedOptions.Contents).Count()).IsEqualTo(8);

        await Assert.That(ws.Cell("A1").Value).IsEqualTo(1);
        await Assert.That(ws.Cell("B1").Value).IsEqualTo("String1");
        await Assert.That(ws.Cell("A2").Value).IsEqualTo(2);
        await Assert.That(ws.Cell("B2").Value).IsEqualTo("String2");

        await Assert.That(ws.Cell("C1").FormulaA1).IsEqualTo("2*3");
        await Assert.That(ws.Cell("C1").CachedValue).IsEqualTo(6);
        await Assert.That(ws.Cell("C2").FormulaA1).IsEqualTo("2+3");
        await Assert.That(ws.Cell("C3").FormulaA1).IsEqualTo("2-3");
        await Assert.That(ws.Cell("C4").FormulaA1).IsEqualTo("C1+C2");
        await Assert.That(ws.Cell("C4").CachedValue).IsEqualTo(11);
    }

    private static XElement[] SavedCells(MemoryStream package, string partName) =>
        SavedPart(package, partName).Descendants(TransitionalMain + "c").ToArray();

    private static XDocument SavedPart(MemoryStream package, string partName)
    {
        package.Position = 0;
        using var zip = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        using var part = zip.GetEntry(partName)!.Open();
        return XDocument.Load(part);
    }
}
