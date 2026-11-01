using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;
using XLibur.Excel.Drawings;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Drawings.Shapes;

/// <summary>
/// A shape the API creates is written the way Excel writes a shape it adds (spec 15, design rule 7,
/// acceptance criterion 8): only the geometry, the id, the name and the text's language differ.
/// </summary>
public class ShapeCreationTests
{
    internal const string Fixture = @"Other\Shapes\ExcelShapes.xlsx";
    internal const string Drawing1 = "xl/drawings/drawing1.xml";

    internal static readonly XNamespace Xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    internal static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    [Test]
    public async Task A_new_text_box_is_written_as_Excel_writes_one()
    {
        var created = SaveNew(ws => ws.Shapes.AddTextBox("Hello", ws.Cell("B2")).WithSize(200, 60));

        await Assert.That(Changes(ExcelShape("Plain TextBox"), created)).IsEqualTo(XmlChangeSet.Expect());
    }

    [Test]
    public async Task A_new_preset_shape_is_written_as_Excel_writes_one()
    {
        var created = SaveNew(ws => ws.Shapes.Add(XLShapeType.Rectangle, ws.Cell("F2")));

        await Assert.That(Changes(ExcelShape("Plain Rectangle"), created)).IsEqualTo(XmlChangeSet.Expect());
    }

    [Test]
    public async Task A_new_shape_is_anchored_by_its_placement()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Shapes.AddTextBox("one", ws.Cell("C3")).WithSize(200, 60);
        ws.Shapes.AddTextBox("two", 30, 40);
        ws.Shapes.Add(XLShapeType.Ellipse, ws.Cell("A1")).MoveTo(ws.Cell("B2"), ws.Cell("D5"));
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        var anchors = XDocument.Parse(saved.ReadPart(Drawing1)).Root!.Elements().ToList();

        await Assert.That(anchors.Select(a => a.Name.LocalName))
            .IsEquivalentTo(new[] { "oneCellAnchor", "absoluteAnchor", "twoCellAnchor" });
        await Assert.That(anchors[0].Element(Xdr + "from")!.Element(Xdr + "col")!.Value).IsEqualTo("2");
        await Assert.That(anchors[0].Element(Xdr + "ext")!.Attribute("cx")!.Value).IsEqualTo("1905000");
        await Assert.That(anchors[1].Element(Xdr + "pos")!.Attribute("x")!.Value).IsEqualTo("285750");
        await Assert.That(anchors[2].Element(Xdr + "to")!.Element(Xdr + "row")!.Value).IsEqualTo("4");
        await Assert.That(anchors.Select(a => a.Descendants(Xdr + "cNvPr").Single().Attribute("id")!.Value))
            .IsEquivalentTo(new[] { "1", "2", "3" });
    }

    /// <summary>The one <c>xdr:sp</c> a new workbook saves after <paramref name="add"/>.</summary>
    private static XElement SaveNew(System.Action<IXLWorksheet> add)
    {
        using var wb = new XLWorkbook();
        add(wb.AddWorksheet());
        using var saved = new MemoryStream();
        wb.SaveAs(saved);
        return XDocument.Parse(saved.ReadPart(Drawing1)).Descendants(Xdr + "sp").Single();
    }

    /// <summary>The <c>xdr:sp</c> named <paramref name="name"/> in the fixture Excel wrote.</summary>
    internal static XElement ExcelShape(string name)
    {
        using var source = TestHelper.OpenResource(Fixture);
        return XDocument.Parse(source.ReadPart(Drawing1)).Descendants(Xdr + "sp")
            .Single(sp => (string?)sp.Descendants(Xdr + "cNvPr").Single().Attribute("name") == name);
    }

    /// <summary>
    /// The change set between two shapes with what is meant to differ taken out of both: the id,
    /// the name and Excel's creation id, the position and size, and the language of the text.
    /// </summary>
    private static string Changes(XElement excel, XElement created) =>
        XmlChangeSet.Between(Normalised(excel), Normalised(created)).ToString();

    private static string Normalised(XElement shape)
    {
        var copy = new XElement(shape);
        var properties = copy.Descendants(Xdr + "cNvPr").Single();
        properties.RemoveAll();
        copy.Descendants(A + "xfrm").Remove();
        foreach (var attribute in copy.Descendants().Attributes("lang").ToList())
            attribute.Remove();

        return copy.ToString();
    }
}
