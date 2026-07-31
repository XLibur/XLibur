using XLibur.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.DataValidations;

public class DataValidationTests
{
    [Test]
    public async Task Validation_Reference_List_Values_From_Separate_Sheet()
    {
        var wb = new XLWorkbook();
        var valuesSheet = wb.Worksheets.Add("ValuesSheet");
        var cell = valuesSheet.Cell("E1");
        cell.SetValue("Value 1");
        cell = cell.CellBelow();
        cell.SetValue("Value 2");
        cell = cell.CellBelow();
        cell.SetValue("Value 3");
        cell = cell.CellBelow();
        cell.SetValue("Value 4");

        var uiSheet = wb.Worksheets.Add("UI Sheet");
        uiSheet.Cell("A1").SetValue("Cell below has validation with references to the 'ValuesSheet'.");
        cell = uiSheet.Cell("A2");
        cell.GetDataValidation().List(valuesSheet.Range("ValuesSheet!$E$1:$E$4"));

        await Assert.That(cell.GetDataValidation().AllowedValues).IsEqualTo(XLAllowedValues.List);
        await Assert.That(cell.GetDataValidation().Value).IsEqualTo("ValuesSheet!$E$1:$E$4");
    }

    [Test]
    public async Task Validation_1()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Data Validation Issue");
        var cell = ws.Cell("E1");
        cell.SetValue("Value 1");
        cell = cell.CellBelow();
        cell.SetValue("Value 2");
        cell = cell.CellBelow();
        cell.SetValue("Value 3");
        cell = cell.CellBelow();
        cell.SetValue("Value 4");

        ws.Cell("A1").SetValue("Cell below has Validation Only.");
        cell = ws.Cell("A2");
        cell.GetDataValidation().List(ws.Range("$E$1:$E$4"));

        ws.Cell("B1").SetValue("Cell below has Validation with a title.");
        cell = ws.Cell("B2");
        cell.GetDataValidation().List(ws.Range("$E$1:$E$4"));
        cell.GetDataValidation().InputTitle = "Title for B2";

        await Assert.That(cell.GetDataValidation().AllowedValues).IsEqualTo(XLAllowedValues.List);
        await Assert.That(cell.GetDataValidation().Value).IsEqualTo("'Data Validation Issue'!$E$1:$E$4");
        await Assert.That(cell.GetDataValidation().InputTitle).IsEqualTo("Title for B2");

        ws.Cell("C1").SetValue("Cell below has Validation with a message.");
        cell = ws.Cell("C2");
        cell.GetDataValidation().List(ws.Range("$E$1:$E$4"));
        cell.GetDataValidation().InputMessage = "Message for C2";

        await Assert.That(cell.GetDataValidation().AllowedValues).IsEqualTo(XLAllowedValues.List);
        await Assert.That(cell.GetDataValidation().Value).IsEqualTo("'Data Validation Issue'!$E$1:$E$4");
        await Assert.That(cell.GetDataValidation().InputMessage).IsEqualTo("Message for C2");

        ws.Cell("D1").SetValue("Cell below has Validation with title and message.");
        cell = ws.Cell("D2");
        cell.GetDataValidation().List(ws.Range("$E$1:$E$4"));
        cell.GetDataValidation().InputTitle = "Title for D2";
        cell.GetDataValidation().InputMessage = "Message for D2";

        await Assert.That(cell.GetDataValidation().AllowedValues).IsEqualTo(XLAllowedValues.List);
        await Assert.That(cell.GetDataValidation().Value).IsEqualTo("'Data Validation Issue'!$E$1:$E$4");
        await Assert.That(cell.GetDataValidation().InputTitle).IsEqualTo("Title for D2");
        await Assert.That(cell.GetDataValidation().InputMessage).IsEqualTo("Message for D2");
    }

    [Test]
    public async Task Validation_2()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell("A1").SetValue("A");
        ws.Cell("B1").CreateDataValidation().Custom("Sheet1!A1");

        var ws2 = wb.AddWorksheet("Sheet2");
        ws2.Cell("A1").SetValue("B");
        ws.Cell("B1").CopyTo(ws2.Cell("B1"));

        await Assert.That(ws2.Cell("B1").GetDataValidation().Value).IsEqualTo("Sheet1!A1");
    }

    [Test]
    public async Task Validation_3()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell("A1").SetValue("A");
        ws.Cell("B1").CreateDataValidation().Custom("A1");
        ws.FirstRow().InsertRowsAbove(1);

        await Assert.That(ws.Cell("B2").GetDataValidation().Value).IsEqualTo("A2");
    }

    [Test]
    public async Task Validation_4()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell("A1").SetValue("A");
        ws.Cell("B1").CreateDataValidation().Custom("A1");
        ws.Cell("B1").CopyTo(ws.Cell("B2"));
        await Assert.That(ws.Cell("B2").GetDataValidation().Value).IsEqualTo("A2");
    }

    [Test]
    public async Task Validation_5()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell("A1").SetValue("A");
        ws.Cell("B1").CreateDataValidation().Custom("A1");
        ws.FirstColumn().InsertColumnsBefore(1);

        await Assert.That(ws.Cell("C1").GetDataValidation().Value).IsEqualTo("B1");
    }

    [Test]
    public async Task Validation_6()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell("A1").SetValue("A");
        ws.Cell("B1").CreateDataValidation().Custom("A1");
        ws.Cell("B1").CopyTo(ws.Cell("C1"));
        await Assert.That(ws.Cell("C1").GetDataValidation().Value).IsEqualTo("B1");
    }

    [Test]
    public async Task Validation_persists_on_Cell_DataValidation()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("People");

        ws.FirstCell().SetValue("Categories")
            .CellBelow().SetValue("A")
            .CellBelow().SetValue("B");

        var table = ws.RangeUsed().CreateTable();

        var dv = table.DataRange.CreateDataValidation();
        dv.ErrorTitle = "Error";

        await Assert.That(table.DataRange.FirstCell().GetDataValidation().ErrorTitle).IsEqualTo("Error");
    }

    [Test]
    public async Task Validation_persists_on_Worksheet_DataValidations()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("People");

        ws.FirstCell().SetValue("Categories")
            .CellBelow().SetValue("A");

        var table = ws.RangeUsed().CreateTable();

        var dv = table.DataRange.CreateDataValidation();
        dv.ErrorTitle = "Error";

        await Assert.That(ws.DataValidations.Single().ErrorTitle).IsEqualTo("Error");
    }

    [Test]
    [Arguments("A1:C3", 5, false, "A1:C3")]
    [Arguments("A1:C3", 2, false, "A1:C4")]
    [Arguments("A1:C3", 1, false, "A2:C4")]
    [Arguments("A1:C3", 5, true, "A1:C3")]
    [Arguments("A1:C3", 2, true, "A1:C4")]
    [Arguments("A1:C3", 1, true, "A2:C4")]
    public async Task DataValidationShiftedOnRowInsert(string initialAddress, int rowNum, bool setValue,
        string expectedAddress)
    {
        // Arrange
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("DataValidation");
        var validation = ws.Range(initialAddress).CreateDataValidation();
        validation.WholeNumber.Between(0, 100);
        if (setValue)
            ws.Range(initialAddress).Value = 50;

        // Act
        ws.Row(rowNum).InsertRowsAbove(1);

        // Assert
        await Assert.That(ws.DataValidations.Count()).IsEqualTo(1);
        await Assert.That(ws.DataValidations.First().Ranges.Count()).IsEqualTo(1);
        await Assert.That(ws.DataValidations.First().Ranges.First().RangeAddress.ToString()).IsEqualTo(expectedAddress);
    }

    [Test]
    [Arguments("A1:C3", 5, false, "A1:C3")]
    [Arguments("A1:C3", 2, false, "A1:D3")]
    [Arguments("A1:C3", 1, false, "B1:D3")]
    [Arguments("A1:C3", 5, true, "A1:C3")]
    [Arguments("A1:C3", 2, true, "A1:D3")]
    [Arguments("A1:C3", 1, true, "B1:D3")]
    public async Task DataValidationShiftedOnColumnInsert(string initialAddress, int columnNum, bool setValue,
        string expectedAddress)
    {
        // Arrange
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("DataValidation");
        var validation = ws.Range(initialAddress).CreateDataValidation();
        validation.WholeNumber.Between(0, 100);
        if (setValue)
            ws.Range(initialAddress).Value = 50;

        // Act
        ws.Column(columnNum).InsertColumnsBefore(1);

        // Assert
        await Assert.That(ws.DataValidations.Count()).IsEqualTo(1);
        await Assert.That(ws.DataValidations.First().Ranges.Count()).IsEqualTo(1);
        await Assert.That(ws.DataValidations.First().Ranges.First().RangeAddress.ToString()).IsEqualTo(expectedAddress);
    }

    [Test]
    public async Task DataValidationClearSplitsRange()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("DataValidation");
        var validation = ws.Range("A1:C3").CreateDataValidation();
        validation.WholeNumber.Between(0, 100);

        // Act
        ws.Cell("B2").Clear(XLClearOptions.DataValidation);

        // Assert
        await Assert.That(ws.Cell("B2").HasDataValidation).IsFalse();
        await Assert.That(ws.Range("A1:C3").Cells().Where(c => c.Address.ToString() != "B2").All(c => c.HasDataValidation)).IsTrue();
    }

    [Test]
    public async Task ClearingDataValidationOnASheetWithoutAnyIsANoOp()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("DataValidation");

        // Clear used to add a rule covering the range and delete it again regardless, which left the
        // collection empty but cost more the bigger the sheet's validation index had grown (#271).
        ws.Range("A1:C3").Clear(XLClearOptions.DataValidation);
        ws.Range("A1:C3").Clear(XLClearOptions.All);

        await Assert.That(ws.DataValidations.Count()).IsEqualTo(0);
        await Assert.That(ws.Cell("B2").HasDataValidation).IsFalse();
    }

    [Test]
    public async Task ClearingDataValidationLeavesRulesOutsideTheRangeAlone()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("DataValidation");
        ws.Range("E1:E3").CreateDataValidation().WholeNumber.Between(0, 100);

        ws.Range("A1:C3").Clear(XLClearOptions.All);

        await Assert.That(ws.DataValidations.Count()).IsEqualTo(1);
        await Assert.That(ws.Range("E1:E3").Cells().All(c => c.HasDataValidation)).IsTrue();
    }

    [Test]
    public async Task CopyToClearsOnlyTheValidationItOverwrites()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("DataValidation");
        ws.Range("A1:C1").Value = "template";
        ws.Range("A5:C5").CreateDataValidation().WholeNumber.Between(0, 100);
        ws.Range("A6:C6").CreateDataValidation().WholeNumber.Between(0, 100);

        ws.Range("A1:C1").CopyTo(ws.Cell("A5"));

        await Assert.That(ws.Range("A5:C5").Cells().Any(c => c.HasDataValidation)).IsFalse();
        await Assert.That(ws.Range("A6:C6").Cells().All(c => c.HasDataValidation)).IsTrue();
        await Assert.That(ws.Cell("A5").GetString()).IsEqualTo("template");
    }

    [Test]
    public async Task NewDataValidationSplitsRange()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("DataValidation");
        var validation = ws.Range("A1:C3").CreateDataValidation();
        validation.WholeNumber.Between(10, 100);

        // Act
        ws.Cell("B2").CreateDataValidation().WholeNumber.Between(-100, -0);

        // Assert
        await Assert.That(ws.Cell("B2").GetDataValidation().MinValue).IsEqualTo("-100");
        await Assert.That(ws.Range("A1:C3").Cells().Where(c => c.Address.ToString() != "B2").All(c => c.HasDataValidation)).IsTrue();
        await Assert.That(ws.Range("A1:C3").Cells().Where(c => c.Address.ToString() != "B2")
            .All(c => c.GetDataValidation().MinValue == "10")).IsTrue();
    }

    [Test]
    public async Task LongListValue_SavedViaExtensionFormat()
    {
        var values = string.Join(",", Enumerable.Range(1, 20)
            .Select(i => Guid.NewGuid().ToString("N")));

        await Assert.That(values.Length).IsGreaterThan(255);

        using var wb = new XLWorkbook();
        var dv = wb.AddWorksheet("Sheet 1").Cell(1, 1).GetDataValidation();
        dv.List(values);

        await Assert.That(dv.Value).IsEqualTo("\"" + values + "\"");

        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        ms.Position = 0;
        using var wb2 = new XLWorkbook(ms);
        var dv2 = wb2.Worksheet(1).Cell(1, 1).GetDataValidation();
        await Assert.That(dv2.Value).IsEqualTo("\"" + values + "\"");
    }

    [Test]
    public async Task CannotCreateDataValidationWithoutRange()
    {
        await Assert.That(() => new XLDataValidation(null)).Throws<ArgumentNullException>();
    }

    [Test]
    public async Task DataValidationHasWorksheetAndRangesWhenCreated()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range("A1:A3");

        var dv = new XLDataValidation(range);

        await Assert.That(dv.Worksheet).IsSameReferenceAs(ws);
        await Assert.That(dv.Ranges.Single()).IsSameReferenceAs(range);
    }

    [Test]
    public async Task CanAddRangeFromSameWorksheet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range1 = ws.Range("A1:A3");
        var range2 = ws.Range("C1:C3");
        var ranges3 = ws.Ranges("D1:D3,F1:F3");
        var dv = new XLDataValidation(range1);

        dv.AddRange(range2);
        dv.AddRanges(ranges3);

        await Assert.That(dv.Ranges.Any(r => r == range1)).IsTrue();
        await Assert.That(dv.Ranges.Any(r => r == range2)).IsTrue();
        await Assert.That(dv.Ranges.Any(r => r == ranges3.First())).IsTrue();
        await Assert.That(dv.Ranges.Any(r => r == ranges3.Last())).IsTrue();
    }

    [Test]
    public async Task CanAddRangeFromAnotherWorksheet()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet();
        var ws2 = wb.AddWorksheet();
        var range1 = ws1.Range("A1:A3");
        var range2 = ws2.Range("C1:C3");
        var dv = new XLDataValidation(range1);

        dv.AddRange(range2);

        await Assert.That(dv.Ranges.Any(r => r != range2 && r.RangeAddress.ToString() == range2.RangeAddress.ToString())).IsTrue();
    }

    [Test]
    public async Task CanClearRanges()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range1 = ws.Range("A1:A3");
        var range2 = ws.Range("C1:C3");
        var ranges3 = ws.Ranges("D1:D3,F1:F3");
        var dv = new XLDataValidation(range1);
        dv.AddRange(range2);
        dv.AddRanges(ranges3);

        dv.ClearRanges();

        await Assert.That(dv.Ranges).IsEmpty();
    }

    [Test]
    public async Task CanRemoveExistingRange()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range1 = ws.Range("A1:A3");
        var range2 = ws.Range("C1:C3");

        var dv = new XLDataValidation(range1);
        dv.AddRange(range2);

        dv.RemoveRange(range1);

        await Assert.That(dv.Ranges.Single()).IsSameReferenceAs(range2);
    }

    [Test]
    public async Task RemovingExistingRangeDoesNoFail()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range1 = ws.Range("A1:A3");
        var range2 = ws.Range("C1:C3");

        var dv = new XLDataValidation(range1);

        dv.RemoveRange(range2);
        dv.RemoveRange(null);

        await Assert.That(dv.Ranges.Single()).IsSameReferenceAs(range1);
    }

    [Test]
    public async Task AddRangeFiresEvent()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range1 = ws.Range("A1:A3");
        var range2 = ws.Range("C1:C3");
        var dv = new XLDataValidation(range1);

        IXLRange addedRange = null;

        dv.RangeAdded += (s, e) => addedRange = e.Range;

        dv.AddRange(range2);

        await Assert.That(addedRange).IsSameReferenceAs(range2);
    }

    [Test]
    [Arguments(XLAllowedValues.List)]
    [Arguments(XLAllowedValues.Custom)]
    [Arguments(XLAllowedValues.AnyValue)]
    public async Task DataValidation_DoesNotWriteOperatorAttribute_ForTypesWithoutOperator(XLAllowedValues allowedValues)
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            var dv = ws.Range("A1:A5").CreateDataValidation();
            switch (allowedValues)
            {
                case XLAllowedValues.List:
                    dv.List("\"Yes,No\"");
                    break;
                case XLAllowedValues.Custom:
                    dv.Custom("A1>0");
                    break;
                case XLAllowedValues.AnyValue:
                    // AnyValue is the default, just set input message to create a validation
                    dv.InputMessage = "Enter any value";
                    break;
                case XLAllowedValues.WholeNumber:
                case XLAllowedValues.Decimal:
                case XLAllowedValues.Date:
                case XLAllowedValues.Time:
                case XLAllowedValues.TextLength:
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(allowedValues), allowedValues, null);
            }

            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var worksheetPart = doc.WorkbookPart.WorksheetParts.First();
        var dataValidation = worksheetPart.Worksheet
            .Descendants<DataValidation>()
            .First();

        await Assert.That(dataValidation.Operator).IsNull();
    }

    [Test]
    public async Task AddRangesFiresMultipleEvents()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range1 = ws.Range("A1:A3");
        var ranges = ws.Ranges("D1:D3,F1:F3");
        var dv = new XLDataValidation(range1);

        var addedRanges = new List<IXLRange>();

        dv.RangeAdded += (s, e) => addedRanges.Add(e.Range);

        dv.AddRanges(ranges);

        await Assert.That(addedRanges.Count).IsEqualTo(2);
    }

    [Test]
    public async Task RemoveRangeFiresEvent()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range1 = ws.Range("A1:A3");
        var range2 = ws.Range("C1:C3");
        var dv = new XLDataValidation(range1);
        dv.AddRange(range2);
        IXLRange removedRange = null;
        dv.RangeRemoved += (s, e) => removedRange = e.Range;

        dv.RemoveRange(range2);

        await Assert.That(removedRange).IsSameReferenceAs(range2);
    }

    [Test]
    public async Task RemoveNonExistingRangeDoesNotFireEvent()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range1 = ws.Range("A1:A3");
        var range2 = ws.Range("C1:C3");
        var dv = new XLDataValidation(range1);

        var fired = false;
        dv.RangeRemoved += (_, _) => fired = true;

        dv.RemoveRange(range2);

        await Assert.That(fired).IsFalse().Because("Expected not to fire event");
    }

    [Test]
    public async Task ClearRangesFiresMultipleEvents()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range1 = ws.Range("A1:A3");
        var range2 = ws.Range("C1:C3");
        var dv = new XLDataValidation(range1);
        dv.AddRange(range2);

        var removedRanges = new List<IXLRange>();

        dv.RangeRemoved += (s, e) => removedRanges.Add(e.Range);

        dv.ClearRanges();

        await Assert.That(removedRanges.Count).IsEqualTo(2);
    }

    [Test]
    [Arguments("$F$2:$F$8", "$F$2:$F$8")]
    [Property("Description", "Cell reference stored verbatim")]
    [Arguments("Sheet1!$A$1:$A$10", "Sheet1!$A$1:$A$10")]
    [Property("Description", "Sheet-qualified reference stored verbatim")]
    [Arguments("\"foobar\"", "\"foobar\"")]
    [Property("Description", "Already quoted string stored verbatim")]
    [Arguments("\"foo,bar,baz\"", "\"foo,bar,baz\"")]
    [Property("Description", "Already quoted CSV stored verbatim")]
    [Arguments("=YesNo", "=YesNo")]
    [Property("Description", "Formula reference stored verbatim")]
    [Arguments("=Sheet1!$A$1:$A$10", "=Sheet1!$A$1:$A$10")]
    [Property("Description", "Formula with sheet reference stored verbatim")]
    [Arguments("foobar", "\"foobar\"")]
    [Property("Description", "Literal string gets quoted")]
    [Arguments("foo,bar,baz", "\"foo,bar,baz\"")]
    [Property("Description", "Literal CSV list gets quoted")]
    [Arguments("123abc", "\"123abc\"")]
    [Property("Description", "String starting with number gets quoted")]
    [Arguments("MyNamedRange", "\"MyNamedRange\"")]
    [Property("Description", "Non-existent named range gets quoted")]
    public async Task List_String_AutoQuotesLiteralStrings(string input, string expectedValue)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var dv = ws.Cell("A1").CreateDataValidation();

        dv.List(input);

        await Assert.That(dv.AllowedValues).IsEqualTo(XLAllowedValues.List);
        await Assert.That(dv.Value).IsEqualTo(expectedValue);
    }

    [Test]
    public async Task List_String_ExistingNamedRangeStoredVerbatim()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // Create a workbook-scoped named range
        wb.DefinedNames.Add("MyNamedRange", ws.Range("E1:E5"));

        var dv = ws.Cell("A1").CreateDataValidation();
        dv.List("MyNamedRange");

        await Assert.That(dv.Value).IsEqualTo("MyNamedRange");
    }

    [Test]
    public async Task List_String_WorksheetScopedNamedRangeStoredVerbatim()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // Create a worksheet-scoped named range
        ws.DefinedNames.Add("LocalRange", ws.Range("E1:E5"));

        var dv = ws.Cell("A1").CreateDataValidation();
        dv.List("LocalRange");

        await Assert.That(dv.Value).IsEqualTo("LocalRange");
    }

    [Test]
    public async Task Issue1711_ListWithPreQuotedString_AndAutoFilter()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // Set up headers
        ws.Cell("A1").Value = "User";
        ws.Cell("B1").Value = "Date";
        ws.Cell("C1").Value = "Error";
        ws.Cell("D1").Value = "Is Issue";

        // Add data rows
        for (var row = 2; row <= 11; row++)
        {
            ws.Cell(row, 4).Value = row <= 6 ? "Is Issue" : "No Issue";
        }

        // Set up AutoFilter on column 4 with "Is Issue" filter
        ws.RangeUsed().SetAutoFilter().Column(4).AddFilter("Is Issue");

        // User passes a pre-quoted string with comma-separated values
        var errorList = new List<string> { "New", "Backdated", "Old", "Other" };
        var errors = $"\"{string.Join(",", errorList)}\"";

        var dv = ws.Range("C2:C11").CreateDataValidation();
        dv.List(errors, true);

        // Pre-quoted string should be stored verbatim
        await Assert.That(dv.Value).IsEqualTo("\"New,Backdated,Old,Other\"");
    }


}
