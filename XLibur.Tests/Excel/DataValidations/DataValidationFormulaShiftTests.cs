using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.DataValidations;
/// <summary>
/// Covers shifting of cell references *inside* data-validation criteria formulas
/// (formula1/formula2) when rows/columns are inserted or deleted. The validation
/// ranges (sqref) are exercised by <see cref="DataValidationShiftTests"/>; this fixture
/// targets the formula text, which a separate code path
/// (<c>ShiftDataValidationFormula*</c>) re-points.
/// </summary>
public class DataValidationFormulaShiftTests
{
    // ---- Positive cases (formula must change) ----
    [Test]
    public async Task InsertColumn_ShiftsCellReferenceInsideValidationFormula()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var rule = ws.Range("E3:E816").CreateDataValidation();
        rule.Custom("=$D3>0");

        ws.Column(1).InsertColumnsBefore(1);

        await Assert.That(rule.Ranges.Single().RangeAddress.ToString()).IsEqualTo("F3:F816");
        await Assert.That(rule.Value).IsEqualTo("=$E3>0");
    }

    [Test]
    public async Task InsertRow_ShiftsRowReferenceInsideValidationFormula()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var rule = ws.Range("C3:Z3").CreateDataValidation();
        rule.Custom("=D$3>0");

        ws.Row(1).InsertRowsAbove(1);

        await Assert.That(rule.Ranges.Single().RangeAddress.ToString()).IsEqualTo("C4:Z4");
        await Assert.That(rule.Value).IsEqualTo("=D$4>0");
    }

    [Test]
    public async Task DeleteColumn_ShiftsCellReferenceBackInsideValidationFormula()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        // Reference at G3; delete a column to its left -> reference moves back to F3.
        var rule = ws.Range("H3:H10").CreateDataValidation();
        rule.Custom("=$G3>0");

        ws.Column(2).Delete();

        await Assert.That(rule.Ranges.Single().RangeAddress.ToString()).IsEqualTo("G3:G10");
        await Assert.That(rule.Value).IsEqualTo("=$F3>0");
    }

    [Test]
    public async Task DeleteRow_ShiftsRowReferenceBackInsideValidationFormula()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var rule = ws.Range("C8:Z8").CreateDataValidation();
        rule.Custom("=D$8>0");

        ws.Row(2).Delete();

        await Assert.That(rule.Ranges.Single().RangeAddress.ToString()).IsEqualTo("C7:Z7");
        await Assert.That(rule.Value).IsEqualTo("=D$7>0");
    }

    [Test]
    public async Task InsertColumn_ShiftsBothOperandsOfBetweenRule()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var rule = ws.Range("E3:E10").CreateDataValidation();
        rule.WholeNumber.Between("$C3", "$D3");

        ws.Column(1).InsertColumnsBefore(1);

        await Assert.That(rule.MinValue).IsEqualTo("$D3");
        await Assert.That(rule.MaxValue).IsEqualTo("$E3");
    }

    [Test]
    public async Task InsertColumn_ShiftsSameSheetListSourceRange()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var rule = ws.Range("F3:F10").CreateDataValidation();
        rule.List("=$D$3:$D$10");

        ws.Column(1).InsertColumnsBefore(1);

        await Assert.That(rule.Value).IsEqualTo("=$E$3:$E$10");
    }

    [Test]
    public async Task InsertColumn_ShiftsDependentDropdownFormula_LeavesDefinedNamesUntouched()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var rule = ws.Range("E3:E816").CreateDataValidation();
        rule.Custom("=OFFSET(SubCategoryList,MATCH($D3,CategoryList,0)-1,0,COUNTIF(CategoryList,$D3),1)");

        ws.Column(1).InsertColumnsBefore(1);

        await Assert.That(rule.Value).IsEqualTo("=OFFSET(SubCategoryList,MATCH($E3,CategoryList,0)-1,0,COUNTIF(CategoryList,$E3),1)");
    }

    [Test]
    public async Task InsertColumn_ShiftsFormulaForFirstColumnInsert_RangeShifterShortCircuit()
    {
        // The range shifter early-returns for first-column inserts; the formula pass must
        // still run. This is the exact scenario from the original bug report.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var rule = ws.Range("B3:B10").CreateDataValidation();
        rule.Custom("=$A3>0");

        ws.Column(1).InsertColumnsBefore(1);

        await Assert.That(rule.Ranges.Single().RangeAddress.ToString()).IsEqualTo("C3:C10");
        await Assert.That(rule.Value).IsEqualTo("=$B3>0");
    }

    [Test]
    public async Task InsertRow_ShiftsFormulaForFirstRowInsert_RangeShifterShortCircuit()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var rule = ws.Range("C2:Z2").CreateDataValidation();
        rule.Custom("=D$1>0");

        ws.Row(1).InsertRowsAbove(1);

        await Assert.That(rule.Ranges.Single().RangeAddress.ToString()).IsEqualTo("C3:Z3");
        await Assert.That(rule.Value).IsEqualTo("=D$2>0");
    }

    // ---- Negative cases (formula must NOT change) ----

    [Test]
    public async Task InsertColumn_LeavesConstantOperandsUnchanged()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var rule = ws.Range("E3:E10").CreateDataValidation();
        rule.WholeNumber.Between(0, 1);

        ws.Column(1).InsertColumnsBefore(1);

        await Assert.That(rule.MinValue).IsEqualTo("0");
        await Assert.That(rule.MaxValue).IsEqualTo("1");
    }

    [Test]
    public async Task InsertColumn_LeavesQuotedLiteralListUnchanged()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var rule = ws.Range("E3:E10").CreateDataValidation();
        rule.List("\"Yes,No,Maybe\"");

        ws.Column(1).InsertColumnsBefore(1);

        await Assert.That(rule.Value).IsEqualTo("\"Yes,No,Maybe\"");
    }

    [Test]
    public async Task InsertColumn_LeavesCrossSheetListSourceUnchanged_WhenValidationHostSheetMutated()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Other lookup");
        var ws = wb.AddWorksheet("Sheet1");

        var rule = ws.Range("E3:E10").CreateDataValidation();
        rule.List("='Other lookup'!$D$2:$D$9");

        // Mutate the sheet the validation lives on (NOT the referenced sheet);
        // the cross-sheet reference must be untouched.
        ws.Column(1).InsertColumnsBefore(1);

        await Assert.That(rule.Value).IsEqualTo("='Other lookup'!$D$2:$D$9");
    }

    [Test]
    public async Task InsertColumn_ShiftsCrossSheetListSource_WhenReferencedSheetMutated()
    {
        using var wb = new XLWorkbook();
        var lookup = wb.AddWorksheet("Other lookup");
        var ws = wb.AddWorksheet("Sheet1");

        var rule = ws.Range("E3:E10").CreateDataValidation();
        rule.List("='Other lookup'!$D$2:$D$9");

        // Mutating the *referenced* sheet shifts the reference (D -> E).
        lookup.Column(1).InsertColumnsBefore(1);

        await Assert.That(rule.Value).IsEqualTo("='Other lookup'!$E$2:$E$9");
    }

    [Test]
    public async Task InsertColumn_LeavesReferenceBeforeInsertionPointUnchanged()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        // Reference at $B3; insert at column F (well to the right) -> no shift.
        var rule = ws.Range("J3:J10").CreateDataValidation();
        rule.Custom("=$B3>0");

        ws.Column(6).InsertColumnsBefore(1);

        await Assert.That(rule.Value).IsEqualTo("=$B3>0");
    }

    // ---- Round-trip (save -> reopen with OpenXML) ----

    [Test]
    public async Task InsertColumn_ShiftedFormulaSurvivesSaveAndReload()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            var rule = ws.Range("E3:E816").CreateDataValidation();
            rule.Custom("=$D3>0");

            ws.Column(1).InsertColumnsBefore(1);
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var doc = SpreadsheetDocument.Open(saved, false);
        var sheetPart = doc.WorkbookPart!.WorksheetParts.First();
        var dv = sheetPart.Worksheet.Descendants<DataValidation>().Single();

        await Assert.That(dv.SequenceOfReferences!.InnerText).IsEqualTo("F3:F816");
        await Assert.That(dv.Formula1!.InnerText.TrimStart('=')).IsEqualTo("$E3>0");
    }
}
