using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Cells;
// ReSharper disable once InconsistentNaming
public class XLCellFormulaTests
{
    [Test]
    public async Task CellFormulaIsStrippedOfEqualSign()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell(1, 1).FormulaA1 = "=B1";
        await Assert.That(ws.Cell(1, 1).FormulaA1).IsEqualTo("B1");
    }

    [Test]
    public async Task DataTable_MaintainProperties()
    {
        await Assert.That(() => TestHelper.LoadSaveAndCompare(
            @"Other\Formulas\DataTableFormula-Excel-Input.xlsx",
            @"Other\Formulas\DataTableFormula-Output.xlsx")).ThrowsNothing();
    }

    [Test]
    public async Task SetDynamicFormulaA1_WritesXldaprMetadataAndCmAttribute()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").SetDynamicFormulaA1("IMAGE(\"https://example.com/image.png\",,3,200,200)");

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;

        using var doc = SpreadsheetDocument.Open(ms, false);
        var wbPart = doc.WorkbookPart!;

        // Verify metadata.xml part exists with XLDAPR
        var metaPart = wbPart.CellMetadataPart;
        await Assert.That(metaPart).IsNotNull().Because("CellMetadataPart should exist");

        var metadata = metaPart!.Metadata;
        var metadataType = metadata.MetadataTypes!.Elements<MetadataType>().First();
        await Assert.That(metadataType.Name!.Value).IsEqualTo("XLDAPR");

        // Verify futureMetadata block exists
        var futureMetadata = metadata.Elements<FutureMetadata>().First();
        await Assert.That(futureMetadata.Name!.Value).IsEqualTo("XLDAPR");

        // Verify cellMetadata has one record
        var cellMeta = metadata.GetFirstChild<CellMetadata>()!;
        await Assert.That(cellMeta.Count!.Value).IsEqualTo(ExpectedCellValue.From(1));

        // Verify the cell has cm attribute set
        var sheetPart = wbPart.WorksheetParts.First();
        var sheetData = sheetPart.Worksheet!.GetFirstChild<SheetData>()!;
        var cell = sheetData.Descendants<Cell>().First(c => c.CellReference == "A1");
        await Assert.That(cell.CellMetaIndex).IsNotNull().Because("Cell should have cm attribute");
        await Assert.That(cell.CellMetaIndex!.Value).IsEqualTo(1u);

        // Verify the formula text
        await Assert.That(cell.CellFormula!.Text).Contains("IMAGE(");
    }

    [Test]
    public async Task SetDynamicFormulaA1_NormalFormulaDoesNotGetCmAttribute()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaA1 = "SUM(B1:B10)";

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;

        using var doc = SpreadsheetDocument.Open(ms, false);
        var wbPart = doc.WorkbookPart!;

        // No metadata part should be created for regular formulas
        await Assert.That(wbPart.CellMetadataPart).IsNull();

        // Cell should not have cm attribute
        var sheetPart = wbPart.WorksheetParts.First();
        var sheetData = sheetPart.Worksheet!.GetFirstChild<SheetData>()!;
        var cell = sheetData.Descendants<Cell>().First(c => c.CellReference == "A1");
        await Assert.That(cell.CellMetaIndex).IsNull();
    }

    [Test]
    public async Task SetDynamicFormulaA1_MultipleDynamicFormulasShareSameMetadata()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").SetDynamicFormulaA1("UNIQUE(B1:B10)");
        ws.Cell("A2").SetDynamicFormulaA1("SORT(C1:C10)");

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;

        using var doc = SpreadsheetDocument.Open(ms, false);
        var wbPart = doc.WorkbookPart!;

        // Only one XLDAPR metadata entry should exist
        var metadata = wbPart.CellMetadataPart!.Metadata;
        var cellMeta = metadata.GetFirstChild<CellMetadata>()!;
        await Assert.That(cellMeta.Count!.Value).IsEqualTo(ExpectedCellValue.From(1));

        // Both cells should reference the same cm index
        var sheetPart = wbPart.WorksheetParts.First();
        var sheetData = sheetPart.Worksheet!.GetFirstChild<SheetData>()!;
        var cellA1 = sheetData.Descendants<Cell>().First(c => c.CellReference == "A1");
        var cellA2 = sheetData.Descendants<Cell>().First(c => c.CellReference == "A2");
        await Assert.That(cellA1.CellMetaIndex!.Value).IsEqualTo(1u);
        await Assert.That(cellA2.CellMetaIndex!.Value).IsEqualTo(1u);
    }

    [Test]
    public async Task SetDynamicFormulaA1_StripsEqualSign()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").SetDynamicFormulaA1("=FILTER(A1:A10, B1:B10>0)");
        await Assert.That(ws.Cell("A1").FormulaA1).Contains("FILTER(");
        await Assert.That(ws.Cell("A1").HasFormula).IsTrue();
    }

    [Test]
    public async Task SetDynamicFormulaA1_RoundTripsCorrectly()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").SetDynamicFormulaA1("IMAGE(\"https://example.com/image.png\",,3,200,200)");

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;

        // Re-load the saved workbook
        using var wb2 = new XLWorkbook(ms);
        var ws2 = wb2.Worksheet(1);
        await Assert.That(ws2.Cell("A1").HasFormula).IsTrue();
        await Assert.That(ws2.Cell("A1").FormulaA1).Contains("IMAGE(");

        // Save again and verify metadata still present
        using var ms2 = new MemoryStream();
        wb2.SaveAs(ms2);
        ms2.Position = 0;

        using var doc = SpreadsheetDocument.Open(ms2, false);
        var cell = doc.WorkbookPart!.WorksheetParts.First()
            .Worksheet!.GetFirstChild<SheetData>()!
            .Descendants<Cell>().First(c => c.CellReference == "A1");

        // The cell should still have cm attribute (round-tripped via CellMetaIndex)
        await Assert.That(cell.CellMetaIndex).IsNotNull();
    }
}
