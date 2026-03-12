using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;
using XLibur.Excel;

namespace XLibur.Tests.Excel.PageSetup;

[TestFixture]
public class PrintAreaTests
{
    [Test]
    [TestCase("A1:B2")]
    [TestCase("A1:B2", "D3:D5")]
    public void CanLoadWorksheetWithMultiplePrintAreas(params string[] printAreaRangeAddresses)
    {
        TestHelper.CreateSaveLoadAssert(
            (_, ws) =>
            {
                foreach (var printAreaRangeAddress in printAreaRangeAddresses)
                    ws.PageSetup.PrintAreas.Add(printAreaRangeAddress);
            },
            (_, ws) =>
            {
                var actualPrintAddresses = ws.PageSetup.PrintAreas.Select(pa => pa.RangeAddress.ToStringRelative());
                Assert.That(actualPrintAddresses, Is.EqualTo(printAreaRangeAddresses));
            });
    }

    [Test]
    [TestCase("OFFSET(Sheet1!$A$1,0,0,10,5)")]
    [TestCase("OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),3)")]
    public void LoadWorkbook_PrintAreaWithFormula_DoesNotThrow(string formula)
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
        Assert.DoesNotThrow(() =>
        {
            using var wb = new XLWorkbook(ms);
        });
    }

    [Test]
    public void LoadAndSave_PrintAreaWithFormula_RoundTrips()
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
            Assert.That(definedNames, Is.Not.Null);

            var printArea = definedNames!
                .OfType<DefinedName>()
                .FirstOrDefault(dn => dn.Name == "_xlnm.Print_Area");

            Assert.That(printArea, Is.Not.Null, "Print area defined name should be preserved after round trip");
            Assert.That(printArea!.Text, Is.EqualTo(formula));
        }
    }
}
