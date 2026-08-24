using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.Drawings;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.DrawingML;

/// <summary>
/// Pins the exact XML XLibur writes through the two pieces of DrawingML machinery spec 16 extracts:
/// the shape-property setters that patch a loaded chart, and the anchor construction behind a saved
/// picture. Tasks 2 and 3 move that code without changing a byte of what it emits; any diff here is
/// a finding to investigate, never noise to re-baseline without a written explanation.
/// </summary>
/// <remarks>
/// <para>
/// This corpus is deliberately separate from <c>ChartGoldenCorpusTests</c>, which spec 22 captured
/// over the path that builds a <em>new</em> chart. Nothing pinned the <em>patch</em> path, and that
/// is where <c>SetFill</c>, <c>SetOutline</c> and the ordered insertion do their real work: replacing
/// a fill that is a choice group, mutating an <c>a:ln</c> in place so its unmodeled children survive,
/// and grafting an element into the middle of a schema sequence. Nothing pinned the drawing part at
/// all.
/// </para>
/// <para>
/// The fixtures are embedded resources. To widen the corpus, add a case below, run once with
/// <c>XLIBUR_WRITE_DRAWINGML_GOLDEN=1</c> to write the file, then rebuild so it is embedded. A
/// missing fixture fails rather than writing itself, so a run can never assert against its own
/// output.
/// </para>
/// </remarks>
public class DrawingMLGoldenCorpusTests
{
    private static readonly GoldenCorpus Corpus =
        new("Excel/DrawingML/Golden", "XLIBUR_WRITE_DRAWINGML_GOLDEN");

    // ── The chart patch path ────────────────────────────────────────────

    /// <summary>
    /// A chart part shaped the way Excel writes one, carrying between its series every starting
    /// state the shape-property setters have to cope with: a solid fill beside an outline that has
    /// children XLibur does not model, a gradient fill that is a different member of the same choice
    /// group, a series with no <c>c:spPr</c> at all, and a marker with shape properties of its own.
    /// </summary>
    private const string PatchSourceChartXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <c:chart>
            <c:plotArea>
              <c:layout/>
              <c:barChart>
                <c:barDir val="col"/>
                <c:grouping val="clustered"/>
                <c:varyColors val="0"/>
                <c:ser>
                  <c:idx val="0"/>
                  <c:order val="0"/>
                  <c:tx><c:v>Solid</c:v></c:tx>
                  <c:spPr>
                    <a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill>
                    <a:ln w="19050" cap="rnd"><a:solidFill><a:srgbClr val="203864"/></a:solidFill><a:round/></a:ln>
                    <a:effectLst/>
                  </c:spPr>
                  <c:invertIfNegative val="0"/>
                  <c:cat><c:strRef><c:f>Data!$A$1:$A$2</c:f></c:strRef></c:cat>
                  <c:val><c:numRef><c:f>Data!$B$1:$B$2</c:f></c:numRef></c:val>
                </c:ser>
                <c:ser>
                  <c:idx val="1"/>
                  <c:order val="1"/>
                  <c:tx><c:v>Gradient</c:v></c:tx>
                  <c:spPr>
                    <a:gradFill>
                      <a:gsLst>
                        <a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>
                        <a:gs pos="100000"><a:srgbClr val="4472C4"/></a:gs>
                      </a:gsLst>
                      <a:lin ang="5400000"/>
                    </a:gradFill>
                    <a:effectLst/>
                  </c:spPr>
                  <c:cat><c:strRef><c:f>Data!$A$1:$A$2</c:f></c:strRef></c:cat>
                  <c:val><c:numRef><c:f>Data!$B$1:$B$2</c:f></c:numRef></c:val>
                </c:ser>
                <c:ser>
                  <c:idx val="2"/>
                  <c:order val="2"/>
                  <c:tx><c:v>Bare</c:v></c:tx>
                  <c:cat><c:strRef><c:f>Data!$A$1:$A$2</c:f></c:strRef></c:cat>
                  <c:val><c:numRef><c:f>Data!$C$1:$C$2</c:f></c:numRef></c:val>
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
                  <c:idx val="3"/>
                  <c:order val="3"/>
                  <c:tx><c:v>Line</c:v></c:tx>
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
            <c:plotVisOnly val="1"/>
          </c:chart>
        </c:chartSpace>
        """;

    [Test]
    [Arguments("patch-fill-replaced")]
    [Arguments("patch-fill-cleared")]
    [Arguments("patch-fill-themed")]
    [Arguments("patch-outline-widened")]
    [Arguments("patch-outline-recoloured")]
    [Arguments("patch-gradient-replaced-by-solid")]
    [Arguments("patch-outline-added")]
    [Arguments("patch-shape-properties-created")]
    [Arguments("patch-marker-formatted")]
    public async Task Patched_chart_part_matches_the_golden_fixture(string name)
    {
        await AssertMatchesGolden(name, CapturePatchedChartXml(name));
    }

    /// <summary>
    /// One edit per fixture. Between them they reach every branch of the shape-property writing that
    /// spec 16 task 3 extracts: replacing a fill, clearing one, writing a theme colour, mutating an
    /// existing outline, adding one that was not there, replacing a different member of the fill
    /// choice group, building <c>c:spPr</c> from nothing and positioning it inside <c>c:ser</c>, and
    /// the marker's own shape properties.
    /// </summary>
    private static void PatchFixture(string name, IXLChart chart)
    {
        var solid = chart.Series.ElementAt(0);
        var gradient = chart.Series.ElementAt(1);
        var bare = chart.Series.ElementAt(2);

        switch (name)
        {
            case "patch-fill-replaced":
                solid.FillColor = XLColor.FromHtml("#00B050");
                break;

            case "patch-fill-cleared":
                solid.FillColor = null;
                break;

            case "patch-fill-themed":
                solid.FillColor = XLColor.FromTheme(XLThemeColor.Accent4);
                break;

            case "patch-outline-widened":
                // The a:ln is edited in place, so cap="rnd", a:round and the colour it came with
                // have to survive an edit that names none of them.
                solid.LineWidthPt = 3;
                break;

            case "patch-outline-recoloured":
                solid.LineColor = XLColor.FromTheme(XLThemeColor.Accent4);
                break;

            case "patch-gradient-replaced-by-solid":
                // a:gradFill and a:solidFill are members of one choice group: the gradient has to go,
                // not sit beside the new fill.
                gradient.FillColor = XLColor.Red;
                break;

            case "patch-outline-added":
                // a:ln belongs after the fill and before a:effectLst, into an a:spPr that already
                // holds both.
                gradient.LineColor = XLColor.FromHtml("#203864");
                gradient.LineWidthPt = 1.5;
                break;

            case "patch-shape-properties-created":
                bare.FillColor = XLColor.Blue;
                break;

            case "patch-marker-formatted":
            {
                var line = chart.SecondarySeries.Single();
                line.MarkerStyle = XLMarkerStyle.Square;
                line.MarkerSize = 10;
                line.MarkerFillColor = XLColor.FromHtml("#FFC000");
                line.Smooth = false;
                break;
            }

            default:
                throw new ArgumentOutOfRangeException(nameof(name), name, "Unknown patch fixture.");
        }
    }

    /// <summary>
    /// Loads <see cref="PatchSourceChartXml"/>, applies one fixture's edit, saves with the OpenXML
    /// validator on, and returns the chart part verbatim.
    /// </summary>
    private static string CapturePatchedChartXml(string name)
    {
        using var original = WorkbookWithChartXml(PatchSourceChartXml);
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(original))
        {
            PatchFixture(name, wb.Worksheet("Data").Charts.First());
            wb.SaveAs(saved, validate: true);
        }

        return FirstPartXml(saved, ChartPartOf);
    }

    /// <summary>Produces a workbook whose single chart part holds the given XML.</summary>
    private static MemoryStream WorkbookWithChartXml(string chartXml)
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

    // ── The picture anchor path ─────────────────────────────────────────

    [Test]
    [Arguments("pictures-three-placements")]
    [Arguments("pictures-default-markers")]
    public async Task Drawing_part_matches_the_golden_fixture(string name)
    {
        await AssertMatchesGolden(name, CaptureDrawingPartXml(name));
    }

    /// <summary>
    /// One sheet per fixture. Between them they reach all three anchor forms and both A1 marker
    /// fallbacks — the ones spec 16 task 2 says move into the factory and become part of its
    /// contract, because a future caller inherits them without being told.
    /// </summary>
    private static void DrawingFixture(string name, IXLWorksheet ws)
    {
        switch (name)
        {
            case "pictures-three-placements":
                // xdr:absoluteAnchor, from pixel coordinates rather than cells.
                AddPicture(ws, "Floating")
                    .WithPlacement(XLPicturePlacement.FreeFloating)
                    .MoveTo(50, 60);

                // xdr:oneCellAnchor: a from marker and an extent.
                AddPicture(ws, "Anchored").MoveTo(ws.Cell("C3"), 12, 8);

                // xdr:twoCellAnchor: a from marker and a to marker.
                AddPicture(ws, "Stretched").MoveTo(ws.Cell("B10"), 3, 4, ws.Cell("F16"), 5, 6);
                break;

            case "pictures-default-markers":
                // A picture nobody moved is MoveAndSize with neither marker set, which is the
                // two-marker A1 fallback: from A1, to A1 offset by the picture's own size.
                AddPicture(ws, "DefaultMoveAndSize");

                // Move with no from marker, which is the one-marker A1 fallback.
                AddPicture(ws, "DefaultMove").WithPlacement(XLPicturePlacement.Move);
                break;

            default:
                throw new ArgumentOutOfRangeException(nameof(name), name, "Unknown drawing fixture.");
        }
    }

    private static string CaptureDrawingPartXml(string name)
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            DrawingFixture(name, ws);
            wb.SaveAs(ms, validate: true);
        }

        return FirstPartXml(ms, doc => doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!);
    }

    private static IXLPicture AddPicture(IXLWorksheet ws, string name)
    {
        // A fresh stream per picture: AddPicture reads from the current position, and the same
        // resource is used for every picture in a fixture.
        using var stream = System.Reflection.Assembly.GetExecutingAssembly()
            .GetManifestResourceStream("XLibur.Tests.Resource.Images.ImageHandling.png")!;
        return ws.AddPicture(stream, name);
    }

    // ── The corpus ──────────────────────────────────────────────────────

    /// <summary>
    /// A golden fixture is only worth anything if the same build produces the same bytes twice. If
    /// this fails, the corpus is not a gate — it is a coin toss — and the source of the variation
    /// has to be found before the other tests here mean anything.
    /// </summary>
    [Test]
    public async Task Capturing_a_fixture_twice_produces_the_same_xml()
    {
        await Assert.That(CapturePatchedChartXml("patch-fill-replaced"))
            .IsEqualTo(CapturePatchedChartXml("patch-fill-replaced"));
        await Assert.That(CaptureDrawingPartXml("pictures-three-placements"))
            .IsEqualTo(CaptureDrawingPartXml("pictures-three-placements"));
    }

    private static async Task AssertMatchesGolden(string name, string capturedXml)
    {
        var actual = GoldenCorpus.Normalise(capturedXml);

        if (Corpus.CanWrite)
            Corpus.Write(name, actual);

        var expected = Corpus.Read(name);
        await Assert.That(expected).IsNotNull()
            .Because($"The corpus holds no fixture for '{name}'. {Corpus.RegenerationHint}");

        await Assert.That(actual).IsEqualTo(expected);
    }

    private static ChartPart ChartPartOf(SpreadsheetDocument doc) =>
        doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!.ChartParts.First();

    private static string FirstPartXml(MemoryStream saved, Func<SpreadsheetDocument, OpenXmlPart> select)
    {
        saved.Position = 0;
        using var doc = SpreadsheetDocument.Open(saved, false);
        using var stream = select(doc).GetStream(FileMode.Open, FileAccess.Read);
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}
