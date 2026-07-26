using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;

namespace XLibur.Tests.Excel.PageSetup;

public class PrintAreaTests
{
    [Test]
    [Arguments("A1:B2")]
    [Arguments("A1:B2", "D3:D5")]
    public async Task CanLoadWorksheetWithMultiplePrintAreas(params string[] printAreaRangeAddresses)
    {
        await TestHelper.CreateSaveLoadAssert(
            (_, ws) =>
            {
                foreach (var printAreaRangeAddress in printAreaRangeAddresses)
                    ws.PageSetup.PrintAreas.Add(printAreaRangeAddress);
            },
            async (_, ws) =>
            {
                var actualPrintAddresses = ws.PageSetup.PrintAreas.Select(pa => pa.RangeAddress.ToStringRelative());
                await Assert.That(actualPrintAddresses).IsEquivalentTo(printAreaRangeAddresses, CollectionOrdering.Matching);
            });
    }

    [Test]
    [Arguments("OFFSET(Sheet1!$A$1,0,0,10,5)")]
    [Arguments("OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),3)")]
    public async Task LoadWorkbook_PrintAreaWithFormula_DoesNotThrow(string formula)
    {
        using var ms = new MemoryStream();

        // Create an xlsx directly via OpenXml SDK with a formula-based print area
        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            workbookPart.Workbook = new Workbook(
                new Sheets(
                    new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Sheet1"
                    }),
                new DefinedNames(
                    new DefinedName
                    {
                        Name = "_xlnm.Print_Area",
                        LocalSheetId = 0,
                        Text = formula
                    }));
        }

        ms.Position = 0;

        // Loading should not throw
        await Assert.That(() =>
        {
            using var wb = new XLWorkbook(ms);
        }).ThrowsNothing();
    }

    [Test]
    public async Task LoadAndSave_PrintAreaWithFormula_RoundTrips()
    {
        var formula = "OFFSET(Sheet1!$A$1,0,0,10,5)";
        using var ms = new MemoryStream();

        // Create an xlsx directly via OpenXml SDK with a formula-based print area
        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            workbookPart.Workbook = new Workbook(
                new Sheets(
                    new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Sheet1"
                    }),
                new DefinedNames(
                    new DefinedName
                    {
                        Name = "_xlnm.Print_Area",
                        LocalSheetId = 0,
                        Text = formula
                    }));
        }

        ms.Position = 0;

        // Load and re-save
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(ms))
        {
            wb.SaveAs(saved);
        }

        // Verify the formula-based print area survived the round trip
        saved.Position = 0;
        using (var doc = SpreadsheetDocument.Open(saved, false))
        {
            var definedNames = doc.WorkbookPart!.Workbook.DefinedNames;
            await Assert.That(definedNames).IsNotNull();

            var printArea = definedNames!
                .OfType<DefinedName>()
                .FirstOrDefault(dn => dn.Name == "_xlnm.Print_Area");

            await Assert.That(printArea).IsNotNull().Because("Print area defined name should be preserved after round trip");
            await Assert.That(printArea!.Text).IsEqualTo(formula);
        }
    }
}
