using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Cells;

// ReSharper disable once InconsistentNaming
[TestFixture]
public class XLCellFormulaTests
{
    [Test]
    public void CellFormulaIsStrippedOfEqualSign()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell(1, 1).FormulaA1 = "=B1";
        Assert.AreEqual("B1", ws.Cell(1, 1).FormulaA1);
    }

    [Test]
    public void DataTable_MaintainProperties()
    {
        Assert.DoesNotThrow(() => TestHelper.LoadSaveAndCompare(
            @"Other\Formulas\DataTableFormula-Excel-Input.xlsx",
            @"Other\Formulas\DataTableFormula-Output.xlsx"));
    }

    [Test]
    public void SetDynamicFormulaA1_WritesXldaprMetadataAndCmAttribute()
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
        Assert.That(metaPart, Is.Not.Null, "CellMetadataPart should exist");

        var metadata = metaPart!.Metadata;
        var metadataType = metadata.MetadataTypes!.Elements<MetadataType>().First();
        Assert.That(metadataType.Name!.Value, Is.EqualTo("XLDAPR"));

        // Verify futureMetadata block exists
        var futureMetadata = metadata.Elements<FutureMetadata>().First();
        Assert.That(futureMetadata.Name!.Value, Is.EqualTo("XLDAPR"));

        // Verify cellMetadata has one record
        var cellMeta = metadata.GetFirstChild<CellMetadata>()!;
        Assert.That(cellMeta.Count!.Value, Is.EqualTo(1));

        // Verify the cell has cm attribute set
        var sheetPart = wbPart.WorksheetParts.First();
        var sheetData = sheetPart.Worksheet!.GetFirstChild<SheetData>()!;
        var cell = sheetData.Descendants<Cell>().First(c => c.CellReference == "A1");
        Assert.That(cell.CellMetaIndex, Is.Not.Null, "Cell should have cm attribute");
        Assert.That(cell.CellMetaIndex!.Value, Is.EqualTo(1u));

        // Verify the formula text
        Assert.That(cell.CellFormula!.Text, Does.Contain("IMAGE("));
    }

    [Test]
    public void SetDynamicFormulaA1_NormalFormulaDoesNotGetCmAttribute()
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
        Assert.That(wbPart.CellMetadataPart, Is.Null);

        // Cell should not have cm attribute
        var sheetPart = wbPart.WorksheetParts.First();
        var sheetData = sheetPart.Worksheet!.GetFirstChild<SheetData>()!;
        var cell = sheetData.Descendants<Cell>().First(c => c.CellReference == "A1");
        Assert.That(cell.CellMetaIndex, Is.Null);
    }

    [Test]
    public void SetDynamicFormulaA1_MultipleDynamicFormulasShareSameMetadata()
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
        Assert.That(cellMeta.Count!.Value, Is.EqualTo(1));

        // Both cells should reference the same cm index
        var sheetPart = wbPart.WorksheetParts.First();
        var sheetData = sheetPart.Worksheet!.GetFirstChild<SheetData>()!;
        var cellA1 = sheetData.Descendants<Cell>().First(c => c.CellReference == "A1");
        var cellA2 = sheetData.Descendants<Cell>().First(c => c.CellReference == "A2");
        Assert.That(cellA1.CellMetaIndex!.Value, Is.EqualTo(1u));
        Assert.That(cellA2.CellMetaIndex!.Value, Is.EqualTo(1u));
    }

    [Test]
    public void SetDynamicFormulaA1_StripsEqualSign()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").SetDynamicFormulaA1("=FILTER(A1:A10, B1:B10>0)");
        Assert.That(ws.Cell("A1").FormulaA1, Does.Contain("FILTER("));
        Assert.That(ws.Cell("A1").HasFormula, Is.True);
    }

    [Test]
    public void SetDynamicFormulaA1_RoundTripsCorrectly()
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
        Assert.That(ws2.Cell("A1").HasFormula, Is.True);
        Assert.That(ws2.Cell("A1").FormulaA1, Does.Contain("IMAGE("));

        // Save again and verify metadata still present
        using var ms2 = new MemoryStream();
        wb2.SaveAs(ms2);
        ms2.Position = 0;

        using var doc = SpreadsheetDocument.Open(ms2, false);
        var cell = doc.WorkbookPart!.WorksheetParts.First()
            .Worksheet!.GetFirstChild<SheetData>()!
            .Descendants<Cell>().First(c => c.CellReference == "A1");

        // The cell should still have cm attribute (round-tripped via CellMetaIndex)
        Assert.That(cell.CellMetaIndex, Is.Not.Null);
    }
}
