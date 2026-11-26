using System.IO;
using System.IO.Compression;
using System.Linq;
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
