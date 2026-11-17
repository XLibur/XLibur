using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Loading;

/// <summary>
/// A workbook saved as ISO/IEC 29500 Strict writes every SpreadsheetML part in the Strict main
/// namespace. Its cells have to load, and survive a save, as a Transitional workbook's do (#900).
/// </summary>
public class StrictWorkbookTests
{
    private const string StrictWorkbook = @"TryToLoad\LO\xlsx\universal-content-strict.xlsx";

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

        await Assert.That(CellElementCount(saved, "xl/worksheets/sheet1.xml")).IsEqualTo(12);

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await AssertSheet1Cells(reloaded.Worksheet("Sheet1"));
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

    private static int CellElementCount(MemoryStream package, string partName)
    {
        package.Position = 0;
        using var zip = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        using var part = zip.GetEntry(partName)!.Open();
        return XDocument.Load(part).Descendants().Count(e => e.Name.LocalName == "c");
    }
}
