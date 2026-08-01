using XLibur.Excel;
using XLibur.Excel.Drawings;
using XLibur.Tests.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Saving;

public class SavingTests
{
    [Test]
    public async Task BooleanValueSavesAsZeroOrOne()
    {
        // When a cell evaluates to a boolean value, the text in the XML has to be true/false (lowercase only) or 0/1
        await TestHelper.CreateAndCompare(() =>
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.FirstCell().FormulaA1 = "=TRUE";
            return wb;
        }, @"Other\Formulas\BooleanFormulaValues.xlsx", evaluateFormulae: true);
    }

    [Test]
    public async Task CanSaveEmptyFile()
    {
        using var ms = new MemoryStream();
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1");
        wb.SaveAs(ms);

        await Assert.That(ms.Length).IsGreaterThan(0);
    }

    [Test]
    public async Task CanSuccessfullySaveFileMultipleTimes()
    {
        using var memoryStream = new MemoryStream();
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");

        // Comments might cause duplicate VmlDrawing Id's - ensure it's tested:
        sheet.Cell(1, 1).GetComment().AddText("abc");

        wb.SaveAs(memoryStream, validate: true);

        for (var i = 1; i <= 3; i++)
        {
            sheet.Cell(i, 1).Value = "test" + i;
            wb.SaveAs(memoryStream, validate: true);
        }

        await Assert.That(sheet.Cell(3, 1).Value).IsEqualTo("test3");
        await Assert.That(memoryStream.Length).IsGreaterThan(0);
    }

    [Test]
    public async Task CanEscape_xHHHH_Correctly()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().Value = "Reserve_TT_A_BLOCAGE_CAG_x6904_2";
            wb.SaveAs(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();
            await Assert.That(ws.FirstCell().Value).IsEqualTo("Reserve_TT_A_BLOCAGE_CAG_x6904_2");
        }
    }

    [Test]
    public async Task CanSaveFileMultipleTimesAfterDeletingWorksheet()
    {
        // https://github.com/XLibur/XLibur/issues/435

        using var ms = new MemoryStream();
        using (var book1 = new XLWorkbook())
        {
            book1.AddWorksheet("sheet1");
            book1.AddWorksheet("sheet2");

            book1.SaveAs(ms);
        }
        ms.Position = 0;

        using (var book2 = new XLWorkbook(ms))
        {
            var ws = book2.Worksheet(1);
            await Assert.That(ws.Name).IsEqualTo("sheet1");
            ws.Delete();
            book2.Save();
            book2.Save();
        }
    }

    [Test]
    public async Task CanSaveAndValidateFileInAnotherCulture()
    {
        string[] cultures = ["it", "de-AT"];

        var originalCulture = Thread.CurrentThread.CurrentCulture;
        try
        {
            foreach (var culture in cultures)
            {
                Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo(culture);

                using var wb = new XLWorkbook();
                var memoryStream = new MemoryStream();
                var ws = wb.Worksheets.Add("Sheet1");

                wb.SaveAs(memoryStream, true);

                await Assert.That(memoryStream.Length).IsGreaterThan(0).Because($"Failed for culture {culture}");
            }
        }
        finally
        {
            Thread.CurrentThread.CurrentCulture = originalCulture;
        }
    }

    [Test]
    public async Task CachedValuePreservedEvenWhenEvaluateFlagIsFalse()
    {
        // Cached values are always written when the formula has been evaluated
        // and is not dirty, regardless of EvaluateFormulasBeforeSaving.
        // The flag only controls whether formulas are re-evaluated during save.
        using var ms = new MemoryStream();
        using (var book1 = new XLWorkbook())
        {
            var sheet = book1.AddWorksheet("sheet1");
            sheet.Cell("A1").Value = 123;
            sheet.Cell("A2").FormulaA1 = "A1*10";
            book1.RecalculateAllFormulas();
            var options = new SaveOptions { EvaluateFormulasBeforeSaving = false };

            book1.SaveAs(ms, options);
        }
        ms.Position = 0;

        using (var book2 = new XLWorkbook(ms))
        {
            var ws = book2.Worksheet(1);

            await Assert.That(ws.Cell("A2").CachedValue).IsEqualTo(1230.0);
        }
    }

    [Test]
    public async Task SaveCachedValueWhenFlagIsTrue()
    {
        using var ms = new MemoryStream();
        using (var book1 = new XLWorkbook())
        {
            var sheet = book1.AddWorksheet("sheet1");
            sheet.Cell("A1").Value = 123;
            sheet.Cell("A2").FormulaA1 = "A1*10";
            sheet.Cell("A3").FormulaA1 = "TEXT(A2, \"# ###\")";
            var options = new SaveOptions { EvaluateFormulasBeforeSaving = true };

            book1.SaveAs(ms, options);
        }
        ms.Position = 0;

        using (var book2 = new XLWorkbook(ms))
        {
            var ws = book2.Worksheet(1);

            await Assert.That(ws.Cell("A2").CachedValue).IsEqualTo(1230);
            await Assert.That(ws.Cell("A3").CachedValue).IsEqualTo("1 230");
        }
    }

    [Test]
    public async Task CanSaveAsCopyReadOnlyFile()
    {
        using var original = new TemporaryFile();
        try
        {
            using var copy = new TemporaryFile();
            // Arrange
            using (var wb = new XLWorkbook())
            {
                var sheet = wb.Worksheets.Add("TestSheet");
                wb.SaveAs(original.Path);
            }
            File.SetAttributes(original.Path, FileAttributes.ReadOnly);

            // Act
            using (var wb = new XLWorkbook(original.Path))
            {
                wb.SaveAs(copy.Path);
            }

            // Assert
            await Assert.That(File.Exists(copy.Path)).IsTrue();
            await Assert.That(File.GetAttributes(copy.Path).HasFlag(FileAttributes.ReadOnly)).IsFalse();
        }
        finally
        {
            // Tear down
            File.SetAttributes(original.Path, FileAttributes.Normal);
        }
    }

    [Test]
    public async Task CanSaveAsOverwriteExistingFile()
    {
        using var existing = new TemporaryFile();
        // Arrange
        File.WriteAllText(existing.Path, "");

        // Act
        using (var wb = new XLWorkbook())
        {
            var sheet = wb.Worksheets.Add("TestSheet");
            wb.SaveAs(existing.Path);
        }

        // Assert
        await Assert.That(File.Exists(existing.Path)).IsTrue();
        await Assert.That(new FileInfo(existing.Path).Length).IsGreaterThan(0);
    }

    [Test]
    // Windows-only: FileAttributes.ReadOnly does not prevent writes on Linux/macOS
    [RunOn(TUnit.Core.Enums.OS.Windows)]
    public async Task CannotSaveAsOverwriteExistingReadOnlyFile()
    {
        using var existing = new TemporaryFile();
        try
        {
            // Arrange
            File.WriteAllText(existing.Path, "");
            File.SetAttributes(existing.Path, FileAttributes.ReadOnly);

            // Act
            Action saveAs = () =>
            {
                using var wb = new XLWorkbook();
                var sheet = wb.Worksheets.Add("TestSheet");
                wb.SaveAs(existing.Path);
            };

            // Assert
            await Assert.That(saveAs).Throws<UnauthorizedAccessException>();
        }
        finally
        {
            // Tear down
            File.SetAttributes(existing.Path, FileAttributes.Normal);
        }
    }

    [Test]
    public async Task PageBreaksDontDuplicateAtSaving()
    {
        // https://github.com/XLibur/XLibur/issues/666

        using var ms = new MemoryStream();
        using (var wb1 = new XLWorkbook())
        {
            var ws = wb1.Worksheets.Add("Page Breaks");
            ws.PageSetup.PrintAreas.Add("A1:D5");
            ws.PageSetup.AddHorizontalPageBreak(2);
            ws.PageSetup.AddVerticalPageBreak(2);
            wb1.SaveAs(ms);
            wb1.Save();
        }
        using (var wb2 = new XLWorkbook(ms))
        {
            var ws = wb2.Worksheets.First();

            await Assert.That(ws.PageSetup.ColumnBreaks.Count).IsEqualTo(1);
            await Assert.That(ws.PageSetup.RowBreaks.Count).IsEqualTo(1);
        }
    }

    [Test]
    public async Task CanSaveFileWithPictureAndComment()
    {
        using var ms = new MemoryStream();
        using var wb = new XLWorkbook();
        using var imageStream = System.Reflection.Assembly.GetAssembly(typeof(XLibur.Examples.BasicTable)).GetManifestResourceStream("XLibur.Examples.Resources.SampleImage.jpg");
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("D4").Value = "Hello world.";

        ws.AddPicture(imageStream, "MyPicture")
            .WithPlacement(XLPicturePlacement.FreeFloating)
            .MoveTo(50, 50)
            .WithSize(200, 200);

        ws.Cell("D4").GetComment().SetVisible().AddText("This is a comment");

        wb.SaveAs(ms);

        await Assert.That(ms.Length).IsGreaterThan(0);
        await Assert.That(ws.Pictures.Count).IsEqualTo(1);
        await Assert.That(ws.Cell("D4").HasComment).IsTrue();
    }

    [Test]
    public async Task PreserveChartsWhenSaving()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Charts\PreserveCharts\inputfile.xlsx"));
        using var ms = new MemoryStream();
        await TestHelper.CreateAndCompare(() =>
        {
            var wb = new XLWorkbook(stream);
            wb.SaveAs(ms);
            return wb;
        }, @"Other\Charts\PreserveCharts\outputfile.xlsx");
    }

    [Test]
    public async Task DeletingAllPicturesRemovesDrawingPart()
    {
        await TestHelper.CreateAndCompare(() =>
        {
            var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx"));
            var wb = new XLWorkbook(stream);
            foreach (var ws in wb.Worksheets)
            {
                var pictureNames = ws.Pictures.Select(pic => pic.Name).ToArray();
                foreach (var name in pictureNames)
                    ws.Pictures.Delete(name);
            }

            return wb;
        }, @"Other\Drawings\NoDrawings\outputfile.xlsx");
    }

    [Test]
    [Arguments("xlsx", SpreadsheetDocumentType.Workbook)]
    [Arguments("xlsm", SpreadsheetDocumentType.MacroEnabledWorkbook)]
    [Arguments("xltx", SpreadsheetDocumentType.Template)]
    [Arguments("xltm", SpreadsheetDocumentType.MacroEnabledTemplate)]
    public async Task SavesAsProperSpreadsheetDocumentType(string extension, SpreadsheetDocumentType expectedType)
    {
        using var tf = new TemporaryFile(Path.ChangeExtension(Path.GetTempFileName(), extension));
        using (var wb = new XLWorkbook())
        {
            wb.Worksheets.Add("Sheet1");
            wb.SaveAs(tf.Path);
        }

        using (var package = SpreadsheetDocument.Open(tf.Path, false))
        {
            await Assert.That(package.DocumentType).IsEqualTo(expectedType);
        }
    }

    [Test]
    public async Task CanSaveTemplateAsWorkbook()
    {
        // See #1375
        using var template = new TemporaryFile(Path.ChangeExtension(Path.GetTempFileName(), "xltx"));
        using var workbook = new TemporaryFile();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet();
            wb.SaveAs(template.Path);
        }
        using (var wb = new XLWorkbook(template.Path))
        {
            wb.SaveAs(workbook.Path);
        }
        using (var package = SpreadsheetDocument.Open(workbook.Path, false))
        {
            await Assert.That(package.DocumentType).IsEqualTo(SpreadsheetDocumentType.Workbook);
        }
    }

    [Test]
    public async Task SaveAsWithNoExtensionFails()
    {
        using var tf = new TemporaryFile("FileWithNoExtension");
        using var wb = new XLWorkbook();
        wb.Worksheets.Add("Sheet1");
        Action action = () => wb.SaveAs(tf.Path);

        await Assert.That(action).Throws<ArgumentException>();
    }

    [Test]
    public async Task SaveAsWithUnsupportedExtensionFails()
    {
        using var tf = new TemporaryFile("FileWithBadExtension.bad");
        using var wb = new XLWorkbook();
        wb.Worksheets.Add("Sheet1");
        Action action = () => wb.SaveAs(tf.Path);

        await Assert.That(action).Throws<ArgumentException>();
    }

    [Test]
    public async Task SaveCellValueWithLeadingQuotationMarkCorrectly()
    {
        var formulaValue = "=IF(TRUE, 1, 0)";
        var quotedFormulaValue = '\'' + formulaValue;
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            var cell = ws.FirstCell();
            cell.SetValue(quotedFormulaValue);
            await Assert.That(cell.HasFormula).IsFalse();
            await Assert.That(cell.Value).IsEqualTo(formulaValue);
            await Assert.That(cell.DataType).IsEqualTo(XLDataType.Text);
            await Assert.That(cell.Style.IncludeQuotePrefix).IsTrue();

            wb.SaveAs(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();
            var cell = ws.FirstCell();
            await Assert.That(cell.HasFormula).IsFalse();
            await Assert.That(cell.HasFormula).IsFalse();
            await Assert.That(cell.Value).IsEqualTo(formulaValue);
            await Assert.That(cell.DataType).IsEqualTo(XLDataType.Text);
            await Assert.That(cell.Style.IncludeQuotePrefix).IsTrue();
        }
    }

    [Test]
    public async Task PreserveHeightOfEmptyRowsOnSaving()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.RowHeight = 50;
            ws.Row(2).Height = 0;
            ws.Row(3).Height = 20;
            ws.Row(4).Height = 100;

            ws.CopyTo("Sheet2");
            wb.SaveAs(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            foreach (var sheetName in new[] { "Sheet1", "Sheet2" })
            {
                var ws = wb.Worksheet(sheetName);

                await Assert.That(ws.Row(1).Height).IsEqualTo(50);
                await Assert.That(ws.Row(2).Height).IsEqualTo(0);
                await Assert.That(ws.Row(3).Height).IsEqualTo(20);
                await Assert.That(ws.Row(4).Height).IsEqualTo(100);
            }
        }
    }

    [Test]
    public async Task PreserveWidthOfEmptyColumnsOnSaving()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Column(2).Width = 0;
            ws.Column(3).Width = 20;
            ws.Column(4).Width = 100;

            ws.CopyTo("Sheet2");
            wb.SaveAs(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            foreach (var sheetName in new[] { "Sheet1", "Sheet2" })
            {
                var ws = wb.Worksheet(sheetName);

                await Assert.That(ws.Column(1).Width).IsEqualTo(ws.ColumnWidth);
                await Assert.That(ws.Column(2).Width).IsEqualTo(0);
                await Assert.That(ws.Column(3).Width).IsEqualTo(20);
                await Assert.That(ws.Column(4).Width).IsEqualTo(100);
            }
        }
    }

    [Test]
    public async Task PreserveAlignmentOnSaving()
    {
        using var input = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\HorizontalAlignment.xlsx"));
        using var output = new MemoryStream();
        using (var wb = new XLWorkbook(input))
        {
            wb.SaveAs(output);
        }

        using (var wb = new XLWorkbook(output))
        {
            await Assert.That(wb.Worksheets.First().Cell("B1").Style.Alignment.Horizontal).IsEqualTo(XLAlignmentHorizontalValues.Center);
        }
    }

    [Test]
    public async Task PreserveMultipleColorScalesOnSaving()
    {
        using var output = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var sheet = wb.Worksheets.Add("test");
            sheet.Column(1).AddConditionalFormat().ColorScale().LowestValue(XLColor.Red)
                .HighestValue(XLColor.Green);

            sheet.Column(2).AddConditionalFormat().ColorScale().LowestValue(XLColor.Alizarin)
                .HighestValue(XLColor.Blue);

            wb.SaveAs(output);
        }

        using (var wb = new XLWorkbook(output))
        {
            var sheet = wb.Worksheets.First();
            var cf = sheet.ConditionalFormats
                .OrderBy(x => x.Range.RangeAddress.FirstAddress.ColumnNumber)
                .ToArray();
            await Assert.That(cf.Length).IsEqualTo(2);
            await Assert.That(cf[0].ConditionalFormatType).IsEqualTo(XLConditionalFormatType.ColorScale);
            await Assert.That(cf[0].Colors[1]).IsEqualTo(XLColor.Red);
            await Assert.That(cf[0].ContentTypes[1]).IsEqualTo(XLCFContentType.Minimum);
            await Assert.That(cf[0].Colors[2]).IsEqualTo(XLColor.Green);
            await Assert.That(cf[0].ContentTypes[2]).IsEqualTo(XLCFContentType.Maximum);
            await Assert.That(cf[1].ConditionalFormatType).IsEqualTo(XLConditionalFormatType.ColorScale);
            await Assert.That(cf[1].Colors[1]).IsEqualTo(XLColor.Alizarin);
            await Assert.That(cf[1].ContentTypes[1]).IsEqualTo(XLCFContentType.Minimum);
            await Assert.That(cf[1].Colors[2]).IsEqualTo(XLColor.Blue);
            await Assert.That(cf[1].ContentTypes[2]).IsEqualTo(XLCFContentType.Maximum);
        }
    }

    [Test]
    public async Task RemoveExistingInlineStringsIfRequired()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\InlineStrings\inputfile.xlsx"));
        using var ms = new MemoryStream();
        await TestHelper.CreateAndCompare(() =>
        {
            var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet(1);

            var numericCells = ws.CellsUsed(c => double.TryParse(c.GetString(), out var _));
            var textCells = ws.CellsUsed(c => !double.TryParse(c.GetString(), out var _));

            foreach (var cell in numericCells)
            {
                cell.Clear(XLClearOptions.AllFormats);

                // This lambda builds the workbook and must return IXLWorkbook, so it cannot
                // be async and cannot await an assertion. Throwing fails the test the same way.
                if (!cell.Value.TryConvert(out double val, CultureInfo.CurrentCulture))
                {
                    throw new InvalidOperationException(
                        $"Cell {cell.Address} was selected as numeric but did not convert to a number.");
                }

                cell.Value = val;
            }

            foreach (var cell in textCells)
            {
                cell.ShareString = true;
            }

            wb.SaveAs(ms);

            return wb;
        }, @"Other\InlineStrings\outputfile.xlsx");
    }

    [Test]
    public async Task CanSaveFileWithEmptyFill()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\EmptyFill.xlsx"));
        using var wb = new XLWorkbook(stream);
        using var ms = new MemoryStream();
        await Assert.That(() => wb.SaveAs(ms, false)).ThrowsNothing();
    }

    [Test]
    public async Task CanSaveSingleRowAutoFilter()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\SingleRowAutoFilter.xlsx"));
        using var wb = new XLWorkbook(stream);
        using var ms = new MemoryStream();
        await Assert.That(() => wb.SaveAs(ms, false)).ThrowsNothing();
    }

    [Test]
    public async Task PivotTableWithVeryLongField()
    {
        await TestHelper.CreateAndCompare(() =>
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            var longText = string.Join(" ", Enumerable.Range(0, 40).Select(i => "1234567890"));

            var data = new[]
            {
                new { Col1 = longText, Col2 = 2}
            };

            var table = ws.FirstCell().InsertTable(data);

            var pvtSheet = wb.AddWorksheet("pvt");

            var pvt = table.CreatePivotTable(pvtSheet.FirstCell(), "PivotTable1");
            pvt.RowLabels.Add("Col1");

            return wb;
        }, @"Other\PivotTableReferenceFiles\LongText\outputfile.xlsx");
    }

    [Test]
    public async Task CanSaveFileWithVml_NoComments()
    {
        //See #1285
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\FileWithButton.xlsm"));
        using var wb = new XLWorkbook(stream);
        using var ms = new MemoryStream();
        await Assert.That(() => wb.SaveAs(ms)).ThrowsNothing();
    }

    [Test]
    public async Task CanEnableWorkbookFilterPrivacyAndSaveInWorkbook()
    {
        using var ms = new MemoryStream();

        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet();
            wb.SaveAs(ms, new SaveOptions { FilterPrivacy = true });
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = SpreadsheetDocument.Open(ms, false))
        {
            await Assert.That(wb.WorkbookPart.Workbook.WorkbookProperties.FilterPrivacy!.Value).IsTrue();
        }
    }

    [Test]
    public async Task WorkbookFilterPrivacyIsNotSetByDefault()
    {
        using var ms = new MemoryStream();

        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet();
            wb.SaveAs(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = SpreadsheetDocument.Open(ms, false))
        {
            await Assert.That(wb.WorkbookPart.Workbook.WorkbookProperties.FilterPrivacy).IsNull();
        }
    }

    [Test]
    public async Task WorkbookFilterPrivacyIsReadCorrectly()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\FilterPrivacyEnabledWorkbook.xlsx"));
        using var wb = SpreadsheetDocument.Open(stream, false);
        await Assert.That(wb.WorkbookPart.Workbook.WorkbookProperties.FilterPrivacy!.Value).IsTrue();
    }

    [Test]
    public async Task CanSaveAsWithDataValidationAfterInsertFirstRowsAboveAndInsertFirstColumnsBefore()
    {
        using var wb = new XLWorkbook();
        using var ms = new MemoryStream();
        var ws = wb.AddWorksheet("WithDataValidation");
        ws.Range("B4:B4").CreateDataValidation().WholeNumber.Between(0, 1);

        ws.Row(1).InsertRowsAbove(1);
        var dv = ws.DataValidations.ToArray();
        await Assert.That(dv.Length).IsEqualTo(1);
        await Assert.That(dv[0].Ranges.Single().RangeAddress.ToString()).IsEqualTo("B5:B5");

        await Assert.That(() => wb.SaveAs(ms)).ThrowsNothing();

        ws.Column(1).InsertColumnsBefore(1);
        dv = ws.DataValidations.ToArray();
        await Assert.That(dv.Length).IsEqualTo(1);
        await Assert.That(dv[0].Ranges.Single().RangeAddress.ToString()).IsEqualTo("C5:C5");

        await Assert.That(() => wb.SaveAs(ms)).ThrowsNothing();
    }

    // https://github.com/XLibur/XLibur/issues/1606
    [Test]
    public async Task CanSaveGSheetsFileWithNewComment()
    {
        using var ms = new MemoryStream();
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\GoogleSheets\file1.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();
        ws.Cell(1, 1).CreateComment().AddText("Test");
        await Assert.That(() => wb.SaveAs(ms)).ThrowsNothing();
    }

    [Test]
    public async Task CanSaveFileToDefaultDirectory()
    {
        var filename = $"test-{Guid.NewGuid()}.xlsx";
        try
        {
            using var wb = new XLWorkbook();
            wb.AddWorksheet().FirstCell().SetValue("Hello, world!");
            await Assert.That(() => wb.SaveAs(filename)).ThrowsNothing();
        }
        finally
        {
            File.Delete(filename);
        }
    }

    [Test]
    public async Task CanAddNewPartsInWorkbookWithDuplicateRelIds()
    {
        // Both Sheet1 and drawing have same relIds: rId2
        // We can add a new worksheet even when there are parts with same relId
        await TestHelper.LoadModifyAndCompare(
            @"Other\Parts\MultiplePartsHaveNonUniqueRelId-input.xlsx",
            wb => wb.AddWorksheet(),
            @"Other\Parts\MultiplePartsHaveNonUniqueRelId-output.xlsx");
    }

    [Test]
    public async Task WorksheetWithDrawingCanBeModified()
    {
        // Issue 2080: Drawing was loading the workbook DOM from the worksheet part and
        // the OpenXML SDK was ignoring worksheet changes saved through streaming, but used
        // the eager loaded DOM instead.
        // Shapes are now preserved across load/save (#2377)
        await TestHelper.LoadModifyAndCompare(
            @"Other\Parts\WorksheetWithDrawingCanBeModified-input.xlsx",
            wb =>
            {
                var ws = wb.Worksheets.Single();
                ws.Cell("A1").Value = "B";
            },
            @"Other\Parts\WorksheetWithDrawingCanBeModified-output.xlsx");
    }

    [Test]
    public async Task CorrectlySaveValidationWithSheetReference()
    {
        // When validation with sheet reference loading was first implemented, there was a
        // disconnect between where those validations were being loaded from and where they
        // were being saved to. This led to exceptions being thrown when these validations
        // were loaded/saved multiple times, so this test makes sure that the fix for that
        // issue continues to work by forcing multiple load/save cycles.

        var filename1 = $"test-{Guid.NewGuid()}.xlsx";
        var filename2 = $"test-{Guid.NewGuid()}.xlsx";
        try
        {
            var path = TestHelper.GetResourcePath(@"TryToLoad\ValidationWithSheetReference.xlsx");
            using var stream = TestHelper.GetStreamFromResource(path);

            using var originalWorkbook = new XLWorkbook(stream);
            await Assert.That(() => originalWorkbook.SaveAs(filename1)).ThrowsNothing();

            using var workbook1 = new XLWorkbook(filename1);
            await Assert.That(() => workbook1.SaveAs(filename2)).ThrowsNothing();

            using var workbook2 = new XLWorkbook(filename2);
            var ws = workbook2.Worksheet("UI Sheet");
            var B2 = ws.Cell("B2");
            await Assert.That(B2.GetDataValidation().AllowedValues).IsEqualTo(XLAllowedValues.List);
            await Assert.That(B2.GetDataValidation().Value).IsEqualTo("$E$1:$E$4");
            var A2 = ws.Cell("A2");
            await Assert.That(A2.GetDataValidation().AllowedValues).IsEqualTo(XLAllowedValues.List);
            await Assert.That(A2.GetDataValidation().Value).IsEqualTo("ValuesSheet!$A$1:$A$4");
        }
        finally
        {
            File.Delete(filename1);
            File.Delete(filename2);
        }
    }

    [Test]
    // Windows-only: VML round-trip comparison is platform-dependent: XDocument serialization produces different XML formatting on Linux vs Windows
    [RunOn(TUnit.Core.Enums.OS.Windows)]
    public async Task FormControlsArePreserved()
    {
        // The sheet contains three form controls: two radio buttons and group box.
        // Form controls are rather complex and this test ensures that the saved
        // file still has VML part (that is the source of truth), drawing part
        // (likely a replacement in a decade or two) and three control parts.
        //
        // Also check that custom text of the form controls is preserved (stored in VML).
        await TestHelper.LoadAndAssert(async wb =>
        {
            await Assert.That(wb.Worksheets.Count).IsGreaterThan(0);
        }, @"Other\Shapes\sheet-with-form-controls-input.xlsx");

        await TestHelper.LoadSaveAndCompare(
            @"Other\Shapes\sheet-with-form-controls-input.xlsx",
            @"Other\Shapes\sheet-with-form-controls-output.xlsx");
    }
}
