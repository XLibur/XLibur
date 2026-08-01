using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using XLibur.Excel;
using XLibur.Excel.Drawings;
using XLibur.Tests.Utils;
using XLibur.Extensions;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Loading;
// Tests in this fixture test only the successful loading of existing Excel files,
// i.e. we test that XLibur doesn't choke on a given input file
// These tests DO NOT test that XLibur successfully recognises all the Excel parts or that it can successfully save those parts again.
public class LoadingTests
{
    public static IEnumerable<string> TryToLoad =>
        TestHelper.ListResourceFiles(s =>
            s.Contains(".TryToLoad.") &&
            !s.Contains(".LO."));

    [Test]
    [MethodDataSource(nameof(TryToLoad))]
    public async Task CanSuccessfullyLoadFiles(string file)
    {
        await TestHelper.LoadFile(file);
    }

    [Test]
    public async Task Can_load_and_save_preserves_timelines()
    {
        // Regression test for https://github.com/ClosedXML/ClosedXML/issues/2132
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath("TryToLoad.Timelines_Missing_21232.xlsx"));
        using var wb = new XLWorkbook(stream);

        var ws = wb.AddWorksheet("Sample Sheet");
        ws.Cell("A1").Value = "Hello World!";

        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        ms.Position = 0;
        using var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read);
        var entryNames = zip.Entries.Select(e => e.FullName).ToList();

        await Assert.That(entryNames.Any(n => n.Contains("timelines/timeline1.xml"))).IsTrue().Because("Timeline part is missing");
        await Assert.That(entryNames.Any(n => n.Contains("timelineCaches/timelineCache1.xml"))).IsTrue().Because("Timeline cache part is missing");

        using (var reader = new StreamReader(zip.GetEntry("xl/workbook.xml")!.Open()))
            await Assert.That(reader.ReadToEnd()).Contains("timelineCacheRef").Because("Workbook XML lost timelineCacheRef");

        using (var reader = new StreamReader(zip.GetEntry("xl/worksheets/sheet1.xml")!.Open()))
            await Assert.That(reader.ReadToEnd()).Contains("timelineRef").Because("Worksheet XML lost timelineRef");
    }

    [Test]
    public async Task Can_load_and_save_file_with_external_image_reference()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath("TryToLoad.external_image_reference_2608.xlsx"));
        using var wb = new XLWorkbook(stream);
        using var ms = new MemoryStream();
        await Assert.That(() => wb.SaveAs(ms)).ThrowsNothing();
    }

    [Test]
    [MethodDataSource(nameof(LOFiles))]
    public async Task CanSuccessfullyLoadLOFiles(string file)
    {
        await TestHelper.LoadFile(file);
    }

    public static IEnumerable<string> LOFiles
    {
        get
        {
            // TODO: unpark all files
            var parkedForLater = new[]
            {
                "TryToLoad.LO.xlsx.column-style-autofilter.xlsx",
                "TryToLoad.LO.xlsx.formats.xlsx",
                "TryToLoad.LO.xlsx.pivot_table.shared-group-field.xlsx",
                "TryToLoad.LO.xlsx.pivot_table.shared-nested-dategroup.xlsx",
                "TryToLoad.LO.xlsx.pivottable_bool_field_filter.xlsx",
                "TryToLoad.LO.xlsx.pivottable_date_field_filter.xlsx",
                "TryToLoad.LO.xlsx.pivottable_double_field_filter.xlsx",
                "TryToLoad.LO.xlsx.pivottable_duplicated_member_filter.xlsx",
                "TryToLoad.LO.xlsx.pivottable_rowcolpage_field_filter.xlsx",
                "TryToLoad.LO.xlsx.pivottable_string_field_filter.xlsx",
                "TryToLoad.LO.xlsx.pivottable_tabular_mode.xlsx",
                "TryToLoad.LO.xlsx.pivot_table_first_header_row.xlsx",
                "TryToLoad.LO.xlsx.tdf100709.xlsx",
                "TryToLoad.LO.xlsx.tdf89139_pivot_table.xlsx",
                "TryToLoad.LO.xlsx.universal-content-strict.xlsx",
                "TryToLoad.LO.xlsx.universal-content.xlsx",
                "TryToLoad.LO.xlsx.xf_default_values.xlsx",
                "TryToLoad.LO.xlsm.pass.CVE-2016-0122-1.xlsm",
                "TryToLoad.LO.xlsm.tdf111974.xlsm",
                "TryToLoad.LO.xlsm.vba-user-function.xlsm",
            };

            return TestHelper.ListResourceFiles(s => s.Contains(".LO.") && !parkedForLater.Any(s.Contains));
        }
    }

    [Test]
    public async Task CorrectlyLoadValidationWithSheetReference()
    {
        // Arrange
        var path = TestHelper.GetResourcePath(@"TryToLoad\ValidationWithSheetReference.xlsx");
        using var stream = TestHelper.GetStreamFromResource(path);

        // Act
        using var wb = new XLWorkbook(stream);

        // Assert
        var ws = wb.Worksheet("UI Sheet");
        var b2 = ws.Cell("B2");
        await Assert.That(b2.GetDataValidation().AllowedValues).IsEqualTo(XLAllowedValues.List);
        await Assert.That(b2.GetDataValidation().Value).IsEqualTo("$E$1:$E$4");
        var a2 = ws.Cell("A2");
        await Assert.That(a2.GetDataValidation().AllowedValues).IsEqualTo(XLAllowedValues.List);
        await Assert.That(a2.GetDataValidation().Value).IsEqualTo("ValuesSheet!$A$1:$A$4");
    }

    [Test]
    public async Task CanLoadAndManipulateFileWithEmptyTable()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\EmptyTable.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();
        var table = ws.Tables.First();
        var rangeBefore = table.RangeAddress.ToString();
        table.DataRange.InsertRowsBelow(5);

        await Assert.That(table.RangeAddress.ToString()).IsNotEqualTo(rangeBefore);
    }

    [Test]
    public async Task CanLoadDate1904SystemCorrectly()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\Date1904System.xlsx"));
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook(stream))
        {
            var ws = wb.Worksheets.First();
            var c = ws.Cell("A2");
            await Assert.That(c.DataType).IsEqualTo(XLDataType.DateTime);
            await Assert.That(c.GetDateTime()).IsEqualTo(new DateTime(2017, 10, 27, 21, 0, 0, DateTimeKind.Unspecified));
            wb.SaveAs(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();
            var c = ws.Cell("A2");
            await Assert.That(c.DataType).IsEqualTo(XLDataType.DateTime);
            await Assert.That(c.GetDateTime()).IsEqualTo(new DateTime(2017, 10, 27, 21, 0, 0, DateTimeKind.Unspecified));
            wb.SaveAs(ms);
        }
    }

    [Test]
    public async Task CanLoadAndSaveFileWithMismatchingSheetIdAndRelId()
    {
        // This file's workbook.xml contains:
        // <x:sheet name="Data" sheetId="13" r:id="rId1" />
        // and the mismatch between the sheetId and r:id can create problems.
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\FileWithMismatchSheetIdAndRelId.xlsx"));
        using var wb = new XLWorkbook(stream);
        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        await Assert.That(ms.Length).IsGreaterThan(0);
    }

    [Test]
    public async Task CanLoadBasicPivotTable()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\LoadPivotTables.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheet("PivotTable1");
        var pt = ws.PivotTable("PivotTable1");
        await Assert.That(pt.Name).IsEqualTo("PivotTable1");

        await Assert.That(pt.RowLabels.Count()).IsEqualTo(1);
        await Assert.That(pt.RowLabels.Single().SourceName).IsEqualTo("Name");

        await Assert.That(pt.ColumnLabels.Count()).IsEqualTo(1);
        await Assert.That(pt.ColumnLabels.Single().SourceName).IsEqualTo("Month");

        var pv = pt.Values.Single();
        await Assert.That(pv.CustomName).IsEqualTo("Sum of NumberOfOrders");
        await Assert.That(pv.SourceName).IsEqualTo("NumberOfOrders");
    }

    [Test]
    public async Task CanLoadOrderedPivotTable()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\LoadPivotTables.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheet("OrderedPivotTable");
        var pt = ws.PivotTable("OrderedPivotTable");

        await Assert.That(pt.RowLabels.Single().SortType).IsEqualTo(XLPivotSortType.Ascending);
        await Assert.That(pt.ColumnLabels.Single().SortType).IsEqualTo(XLPivotSortType.Descending);
    }

    [Test]
    public async Task CanLoadPivotTableSubtotals()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\LoadPivotTables.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheet("PivotTableSubtotals");
        var pt = ws.PivotTable("PivotTableSubtotals");

        var subtotals = pt.RowLabels.Get("Group").Subtotals.ToArray();

        await Assert.That(subtotals).IsEquivalentTo([
            XLSubtotalFunction.Automatic,
            XLSubtotalFunction.Average,
            XLSubtotalFunction.Count,
            XLSubtotalFunction.Sum
        ]);
    }

    /// <summary>
    /// Pivot table fields can have a <c>name</c> attribute that renames the field
    /// from its cache source name. When loading, the renamed name is stored as
    /// <c>CustomName</c> while the original cache name is <c>SourceName</c>.
    /// Verify the file can be loaded and saved without errors (#2591).
    /// </summary>
    [Test]
    public async Task CanLoadAndSavePivotTableWithRenamedColumns()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\load_pivottable_renamedcolumns_2591.xlsx"));
        using var wb = new XLWorkbook(stream);
        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        await Assert.That(ms.Length).IsGreaterThan(0);
    }

    [Test]
    [Skip("Pivot table style formats are not implemented, so the border is not read back.")]
    public async Task CanLoadPivotTableWithBorder()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\PivotTableWithBorder.xlsx"));
        using var wb = new XLWorkbook(stream);
        var pt = wb.Worksheet(1).PivotTables.PivotTable("PivotTable1");
        var border = pt.RowLabels.Single().StyleFormats.DataValuesFormat.Style.Border;

        await Assert.That(border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thin);
        await Assert.That(border.TopBorder).IsEqualTo(XLBorderStyleValues.Thin);
        await Assert.That(border.RightBorder).IsEqualTo(XLBorderStyleValues.Thin);
        await Assert.That(border.BottomBorder).IsEqualTo(XLBorderStyleValues.Thin);
    }

    /// <summary>
    /// For non-English locales, the default style ("Normal" in English) can be
    /// another piece of text (e.g. ??????? in Russian).
    /// This test ensures that the default style is correctly detected and
    /// no style conflicts occur on save.
    /// </summary>
    [Test]
    public async Task CanSaveFileWithDefaultStyleNameNotInEnglish()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\FileWithDefaultStyleNameNotInEnglish.xlsx"));
        using var wb = new XLWorkbook(stream);
        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        await Assert.That(ms.Length).IsGreaterThan(0);
    }

    /// <summary>
    /// As per https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.cellvalues(v=office.15).aspx
    /// the 'Date' DataType is available only in files saved with Microsoft Office
    /// In other files, the data type will be saved as numeric
    /// XLibur then deduces the data type by inspecting the number format string
    /// </summary>
    [Test]
    public async Task CanLoadLibreOfficeFileWithDates()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\LibreOfficeFileWithDates.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();
        foreach (var cell in ws.CellsUsed())
        {
            await Assert.That(cell.DataType).IsEqualTo(XLDataType.DateTime);
        }
    }

    [Test]
    public async Task CanLoadFileWithImagesWithCorrectAnchorTypes()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();
        await Assert.That(ws.Pictures.Count).IsEqualTo(2);
        await Assert.That(ws.Pictures.First().Placement).IsEqualTo(XLPicturePlacement.FreeFloating);
        await Assert.That(ws.Pictures.Skip(1).First().Placement).IsEqualTo(XLPicturePlacement.Move);

        var ws2 = wb.Worksheets.Skip(1).First();
        await Assert.That(ws2.Pictures.Count).IsEqualTo(1);
        await Assert.That(ws2.Pictures.First().Placement).IsEqualTo(XLPicturePlacement.MoveAndSize);
    }

    [Test]
    public async Task CanLoadFileWithImagesWithCorrectImageType()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageFormats.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();
        await Assert.That(ws.Pictures.Count).IsEqualTo(1);
        await Assert.That(ws.Pictures.First().Format).IsEqualTo(XLPictureFormat.Jpeg);

        var ws2 = wb.Worksheets.Skip(1).First();
        await Assert.That(ws2.Pictures.Count).IsEqualTo(1);
        await Assert.That(ws2.Pictures.First().Format).IsEqualTo(XLPictureFormat.Png);
    }

    [Test]
    public async Task CanLoadAndDeduceAnchorsFromExcelGeneratedFile()
    {
        // This file was produced by Excel. It contains 3 images, but the latter 2 were copied from the first.
        // There is actually only 1 embedded image if you inspect the file's internals.
        // Additionally, Excel saves all image anchors as TwoCellAnchor, but uses the EditAs attribute to distinguish the types
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\ExcelProducedWorkbookWithImages.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();
        await Assert.That(ws.Pictures.Count).IsEqualTo(3);

        await Assert.That(ws.Picture("Picture 1").Placement).IsEqualTo(XLPicturePlacement.MoveAndSize);
        await Assert.That(ws.Picture("Picture 2").Placement).IsEqualTo(XLPicturePlacement.Move);
        await Assert.That(ws.Picture("Picture 3").Placement).IsEqualTo(XLPicturePlacement.FreeFloating);

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);
    }

    [Test]
    public async Task CanLoadFromTemplate()
    {
        using var tf1 = new TemporaryFile();
        using var tf2 = new TemporaryFile();
        using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\AllShapes.xlsx")))
        using (var wb = new XLWorkbook(stream))
        {
            // Save as temporary file
            wb.SaveAs(tf1.Path);
        }

        var workbook = XLWorkbook.OpenFromTemplate(tf1.Path);
        await Assert.That(workbook.Worksheets.Count != 0).IsTrue();
        await Assert.That(workbook.Save).Throws<InvalidOperationException>();

        workbook.SaveAs(tf2.Path);
    }

    /// <summary>
    /// Excel escapes symbol ' in worksheet title so we have to process this correctly.
    /// </summary>
    [Test]
    public async Task CanOpenWorksheetWithEscapedApostrophe()
    {
        string title = "";

        void OpenWorkbook()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\EscapedApostrophe.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheets.First();
            title = ws.Name;
        }

        await Assert.That(OpenWorkbook).ThrowsNothing();
        await Assert.That(title).IsEqualTo("L'E");
    }

    [Test]
    public async Task CanRoundTripSheetProtectionForObjects()
    {
        using var book = new XLWorkbook();
        var sheet = book.AddWorksheet("TestSheet");
        sheet.Protect()
            .AllowElement(XLSheetProtectionElements.EditObjects | XLSheetProtectionElements.EditScenarios);

        await Assert.That(sheet.Protection.AllowedElements).IsEqualTo(XLSheetProtectionElements.SelectEverything | XLSheetProtectionElements.EditObjects | XLSheetProtectionElements.EditScenarios);

        using var xlStream = new MemoryStream();
        book.SaveAs(xlStream);

        using var persistedBook = new XLWorkbook(xlStream);
        var persistedSheet = persistedBook.Worksheets.Worksheet(1);

        await Assert.That(persistedSheet.Protection.AllowedElements).IsEqualTo(sheet.Protection.AllowedElements);
    }

    [Test]
    [Arguments("A1*10", 1230)]
    [Arguments("A1/10", 12.3)]
    [Arguments("A1&\" cells\"", "123 cells")]
    [Arguments("A1&\"000\"", "123000")]
    [Arguments("ISNUMBER(A1)", true)]
    [Arguments("ISBLANK(A1)", false)]
    [Arguments("DATE(2018,1,28)", 43128)]
    public async Task LoadFormulaCachedValue(string formula, object expectedCachedValue)
    {
        using var ms = new MemoryStream();
        using (var book1 = new XLWorkbook())
        {
            var sheet = book1.AddWorksheet("sheet1");
            sheet.Cell("A1").Value = 123;
            sheet.Cell("A2").FormulaA1 = formula;
            var options = new SaveOptions { EvaluateFormulasBeforeSaving = true };

            book1.SaveAs(ms, options);
        }
        ms.Position = 0;

        using (XLWorkbook book2 = new XLWorkbook(ms))
        {
            var ws = book2.Worksheet(1);
            await Assert.That(ws.Cell("A2").NeedsRecalculation).IsFalse();
            await Assert.That(ws.Cell("A2").CachedValue).IsEqualTo(ExpectedCellValue.From(expectedCachedValue));
        }
    }

    [Test]
    public async Task LoadingOptions()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\Misc\Formulas.xlsx"));
        XLCellValue cachedWithoutRecalculation = default;
        await Assert.That(() =>
        {
            // The cached value from the file is preserved without recalculation.
            using var wb = new XLWorkbook(stream, new LoadOptions { RecalculateAllFormulas = false });
            cachedWithoutRecalculation = wb.Worksheets.Single().Cell("C2").CachedValue;
        }).ThrowsNothing();
        await Assert.That(cachedWithoutRecalculation).IsEqualTo(3.0);

        XLCellValue cachedWithRecalculation = default;
        await Assert.That(() =>
        {
            // Recalculation also produces the correct value.
            using var wb = new XLWorkbook(stream, new LoadOptions { RecalculateAllFormulas = true });
            cachedWithRecalculation = wb.Worksheets.Single().Cell("C2").CachedValue;
        }).ThrowsNothing();
        await Assert.That(cachedWithRecalculation).IsEqualTo(3);

        await Assert.That(new XLWorkbook(stream, new LoadOptions { Dpi = new Point(30, 14) }).DpiX).IsEqualTo(30);
        await Assert.That(new XLWorkbook(stream, new LoadOptions { Dpi = new Point(30, 14) }).DpiY).IsEqualTo(14);
    }

    [Test]
    public async Task CanLoadWorksheetStyle()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\BaseColumnWidth.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheet(1);

        await Assert.That(ws.Style.Font.FontSize).IsEqualTo(8);
        await Assert.That(ws.Style.Font.FontName).IsEqualTo("Arial");
        await Assert.That(ws.Cell("A1").Style.Font.FontSize).IsEqualTo(8);
        await Assert.That(ws.Cell("A1").Style.Font.FontName).IsEqualTo("Arial");
    }

    [Test]
    public async Task CanCorrectLoadWorkbookCellWithStringDataType()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\CellWithStringDataType.xlsx"));
        using var wb = new XLWorkbook(stream);
        var cellToCheck = wb.Worksheet(1).Cell("B2");
        await Assert.That(cellToCheck.DataType).IsEqualTo(XLDataType.Text);
        await Assert.That(cellToCheck.Value).IsEqualTo("String with String Data type");
    }

    [Test]
    public async Task CanCorrectLoadWorkbookCellsWithDateTimeDataTypeOrFormatting()
    {
        const string expected = "03/14/2012 13:30:55";
        await TestHelper.LoadAndAssert(async wb =>
        {
            for (int row = 2; row < 18; row++)
            {
                var cellToCheck = wb.Worksheet(1).Cell(row, 2);
                await Assert.That(cellToCheck.DataType).IsEqualTo(XLDataType.DateTime).Because($"Cell B{row} has incorrect DataType");
                await Assert.That(cellToCheck.Value.ToString(CultureInfo.InvariantCulture)).IsEqualTo(expected).Because($"Cell B{row} value differs");
            }
        }, @"TryToLoad\CellsWithDateTimeDataTypeOrFormatting.xlsx");
    }

    [Test]
    public async Task CanCorrectLoadWorkbookCellsWithTimeSpanDataTypeOrFormatting()
    {
        string[] expected = Enumerable.Range(0, 10).Select(_ => "13:30:55.2").Concat(["0:30:55.2"]).ToArray();
        await TestHelper.LoadAndAssert(async wb =>
        {
            for (int i = 0, row = 2; i < expected.Length; i++, row++)
            {
                var cellToCheck = wb.Worksheet(1).Cell(row, 2);
                await Assert.That(cellToCheck.DataType).IsEqualTo(XLDataType.TimeSpan).Because($"Cell B{row} has incorrect DataType");
                await Assert.That(cellToCheck.Value.ToString(CultureInfo.InvariantCulture)).IsEqualTo(expected[i]).Because($"Cell B{row} value differs");
            }
        }, @"TryToLoad\CellsWithTimeSpanDataTypeOrFormatting.xlsx");
    }

    [Test]
    public async Task CanCorrectLoadWorkbookCellsWithDateTimesWithLocalePrefix()
    {
        await TestHelper.LoadAndAssert(async wb =>
        {
            var ws = wb.Worksheet(1);

            await Assert.That(ws.Cell(1, 1).GetFormattedString()).IsEqualTo("21 January 2019");
            await Assert.That(ws.Cell(2, 1).GetFormattedString()).IsEqualTo("21-Jan-19");
            await Assert.That(ws.Cell(3, 1).GetFormattedString()).IsEqualTo("Monday, 21 January 2019");
            await Assert.That(ws.Cell(4, 1).GetFormattedString()).IsEqualTo("21 Jan 2019");
        }, @"TryToLoad\CellsWithDateTimeWithLocalePrefix.xlsx");
    }

    [Test]
    public async Task CanCorrectLoadWorkbookDefaultColumnWidth()
    {
        using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\Styles\DefaultStyles.xlsx")))
        using (var wb = new XLWorkbook(stream))
        {
            var defaultColumnWidth = wb.ColumnWidth;
            var pixelWidth = XLHelper.NoCToPixels(defaultColumnWidth, wb.Style.Font, wb);
            // Column width depends on font metrics (Calibri on Windows vs Carlito on Linux)
            await Assert.That(defaultColumnWidth).IsEqualTo(8.43).Within(1.5);
            await Assert.That(pixelWidth).IsEqualTo(64).Within(20);
        }

        using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\DefaultColumnWidth.xlsx")))
        using (var wb = new XLWorkbook(stream))
        {
            var defaultColumnWidth = wb.ColumnWidth;
            var pixelWidth = XLHelper.NoCToPixels(defaultColumnWidth, wb.Style.Font, wb);
            // Column width depends on font metrics (Calibri on Windows vs Carlito on Linux)
            await Assert.That(defaultColumnWidth).IsEqualTo(8.5).Within(1.5);
            await Assert.That(pixelWidth).IsEqualTo(56).Within(20);
        }
    }

    [Test]
    public async Task CanCorrectLoadWorksheetBaseColumnWidth()
    {
        // default calibi font case
        // Column widths depend on font metrics (Calibri on Windows vs Carlito on Linux)
        using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\Styles\DefaultStyles.xlsx")))
        using (var wb = new XLWorkbook(stream))
        {
            var ws = wb.Worksheet(1);
            await Assert.That(ws.ColumnWidth).IsEqualTo(8.43).Within(1.5);
            await Assert.That(ws.Column(1).Width).IsEqualTo(8.43).Within(1.5);
        }

        // worksheet has base column width.
        using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\BaseColumnWidth.xlsx")))
        using (var wb = new XLWorkbook(stream))
        {
            var ws = wb.Worksheet(1);
            await Assert.That(ws.ColumnWidth).IsEqualTo(11.17).Within(1.5);
            await Assert.That(ws.Column(1).Width).IsEqualTo(11.17).Within(1.5);
        }
    }

    [Test]
    public async Task CanCorrectLoadWorksheetDefaultColumnWidth()
    {
        // worksheet has default column width.
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\SheetDefaultColumnWidth.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheet(1);
        double pixelWidth = XLHelper.NoCToPixels(ws.Column(1).Width, ws.Style.Font, wb);
        // Column widths depend on font metrics (Calibri on Windows vs Carlito on Linux)
        await Assert.That(ws.ColumnWidth).IsEqualTo(19.75).Within(2.0);
        await Assert.That(pixelWidth).IsEqualTo(163).Within(25);
    }

    [Test]
    public async Task CanLoadFileWithInvalidSelectedRanges()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\SelectedRanges\InvalidSelectedRange.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheet(1);

        await Assert.That(ws.SelectedRanges.Count).IsEqualTo(2);
        await Assert.That(ws.SelectedRanges.First().RangeAddress.ToString()).IsEqualTo("B2:B2");
        await Assert.That(ws.SelectedRanges.Last().RangeAddress.ToString()).IsEqualTo("B2:C2");
    }

    [Test]
    public async Task CanLoadCellsWithoutReferencesCorrectly()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\LO\xlsx\row-index-1-based.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheet(1);

        await Assert.That(ws.Name).IsEqualTo("Page 1");

        var expected = new Dictionary<string, XLCellValue>
        {
            ["A1"] = "Action Plan.Name",
            ["B1"] = "Action Plan.Description",
            ["A2"] = "Jerry",
            ["B2"] = "This is a longer Text.\nSecond line.\nThird line.",
            ["A3"] = Blank.Value,
            ["B3"] = Blank.Value
        };

        foreach (var pair in expected)
            await Assert.That(ws.Cell(pair.Key).Value).IsEqualTo(pair.Value).Because(pair.Key);
    }

    [Test]
    public async Task CorrectlyLoadThemeColors()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\StyleReferenceFiles\ThemeColors\inputfile.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheet(1);

        var c = ws.Cell("A1");
        var themeColor = c.Style.Fill.BackgroundColor.ThemeColor;
        await Assert.That(themeColor).IsEqualTo(XLThemeColor.Accent2);
        await Assert.That(wb.Theme.ResolveThemeColor(themeColor).Color.ToHex()).IsEqualTo("FFED7D31");

        c = ws.Cell("A2");
        themeColor = c.Style.Fill.BackgroundColor.ThemeColor;
        await Assert.That(themeColor).IsEqualTo(XLThemeColor.Accent4);
        await Assert.That(wb.Theme.ResolveThemeColor(themeColor).Color.ToHex()).IsEqualTo("FFFFC000");

        c = ws.Cell("A3");
        themeColor = c.Style.Fill.BackgroundColor.ThemeColor;
        await Assert.That(themeColor).IsEqualTo(XLThemeColor.Accent6);
        await Assert.That(wb.Theme.ResolveThemeColor(themeColor).Color.ToHex()).IsEqualTo("FF70AD47");
    }

    [Test]
    public async Task CorrectlyLoadMergedCellsBorder()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\StyleReferenceFiles\MergedCellsBorder\inputfile.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheet(1);

        var c = ws.Cell("B2");
        await Assert.That(c.Style.Border.TopBorderColor.ColorType).IsEqualTo(XLColorType.Theme);
        await Assert.That(c.Style.Border.TopBorderColor.ThemeColor).IsEqualTo(XLThemeColor.Accent1);
        await Assert.That(c.Style.Border.TopBorderColor.ThemeTint).IsEqualTo(0.39994506668294322d).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task CorrectlyLoadDefaultRowAndColumnStyles()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\StyleReferenceFiles\RowAndColumnStyles\inputfile.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheet(1);

        await Assert.That(ws.Row(1).Style.Font.FontSize).IsEqualTo(8);
        await Assert.That(ws.Row(2).Style.Font.FontSize).IsEqualTo(8);
        await Assert.That(ws.Column("A").Style.Font.FontSize).IsEqualTo(8);
    }

    [Test]
    public async Task EmptyNumberFormatIdTreatedAsGeneral()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\EmptyNumberFormatId.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheet(1);

        await Assert.That(ws.Cell("A2").Style.NumberFormat.NumberFormatId).IsEqualTo(XLPredefinedFormat.General);
    }

    [Test]
    public async Task CanLoadProperties()
    {
        const string author = "TestAuthor";
        const string title = "TestTitle";
        const string subject = "TestSubject";
        const string category = "TestCategory";
        const string keywords = "TestKeywords";
        const string comments = "TestComments";
        const string status = "TestStatus";
        var created = new DateTime(2019, 10, 19, 20, 42, 30, DateTimeKind.Unspecified);
        var modified = new DateTime(2020, 11, 20, 09, 51, 20, DateTimeKind.Unspecified);
        const string lastModifiedBy = "TestLastModifiedBy";
        const string company = "TestCompany";
        const string manager = "TestManager";

        using var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("sheet1");

            wb.Properties.Author = author;
            wb.Properties.Title = title;
            wb.Properties.Subject = subject;
            wb.Properties.Category = category;
            wb.Properties.Keywords = keywords;
            wb.Properties.Comments = comments;
            wb.Properties.Status = status;
            wb.Properties.Created = created;
            wb.Properties.Modified = modified;
            wb.Properties.LastModifiedBy = lastModifiedBy;
            wb.Properties.Company = company;
            wb.Properties.Manager = manager;

            wb.SaveAs(stream, true);
        }

        stream.Position = 0;

        using (var wb = new XLWorkbook(stream))
        {
            await Assert.That(wb.Properties.Author).IsEqualTo(author);
            await Assert.That(wb.Properties.Title).IsEqualTo(title);
            await Assert.That(wb.Properties.Subject).IsEqualTo(subject);
            await Assert.That(wb.Properties.Category).IsEqualTo(category);
            await Assert.That(wb.Properties.Keywords).IsEqualTo(keywords);
            await Assert.That(wb.Properties.Comments).IsEqualTo(comments);
            await Assert.That(wb.Properties.Status).IsEqualTo(status);
            await Assert.That(wb.Properties.Created).IsEqualTo(created);
            await Assert.That(wb.Properties.Modified).IsEqualTo(modified);
            await Assert.That(wb.Properties.LastModifiedBy).IsEqualTo(lastModifiedBy);
            await Assert.That(wb.Properties.Company).IsEqualTo(company);
            await Assert.That(wb.Properties.Manager).IsEqualTo(manager);
        }
    }

    // https://github.com/ClosedXML/ClosedXML/issues/1920
    [Test]
    public async Task CanReadGoogleSheetsCommentText()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\GoogleSheets\GoogleDocExportWithComments.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();

        await Assert.That(ws.Cell("A1").HasComment).IsTrue();
        await Assert.That(ws.Cell("A1").GetComment().Text).IsEqualTo("Toook=12");

        await Assert.That(ws.Cell("A2").HasComment).IsFalse();

        await Assert.That(ws.Cell("A4").HasComment).IsTrue();
        await Assert.That(ws.Cell("A4").GetComment().Text).IsEqualTo("assas");

        await Assert.That(ws.Cell("A7").HasComment).IsTrue();
        await Assert.That(ws.Cell("A7").GetComment().Text).IsEqualTo("12123123" + Environment.NewLine);

        // Verify round-trip: save and reload
        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);
        ms.Position = 0;

        using var wb2 = new XLWorkbook(ms);
        var ws2 = wb2.Worksheets.First();
        await Assert.That(ws2.Cell("A1").GetComment().Text).IsEqualTo("Toook=12");
        await Assert.That(ws2.Cell("A4").GetComment().Text).IsEqualTo("assas");
    }

    [Test]
    public async Task CanLoadEmptyStyles()
    {
        // Stylesheet part exists, but no style collection elements are present
        await TestHelper.LoadAndAssert(async wb =>
        {
            using var ms = new MemoryStream();
            wb.SaveAs(ms, true);
        }, @"TryToLoad\EmptyStyles.xlsx");
    }

    [Test]
    public async Task CanLoadInvalidColors()
    {
        // The styles.xml contains two invalid colors: '0' and 'FED+'. Both
        // should be loaded and no exception thrown. The colors are
        // converted using an Excel algorithm.
        await TestHelper.LoadAndAssert(async wb =>
        {
            var ws = wb.Worksheets.Single();
            await Assert.That(ws.Cell("A1").Style.Font.FontColor).IsEqualTo(XLColor.FromArgb(0xFF000000));
            await Assert.That(ws.Cell("A2").Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromArgb(0xFF000FED));
        }, @"TryToLoad\InvalidColors.xlsx");
    }

    [Test]
    public async Task WontCrashOnSheetsWithoutRelId()
    {
        // Some non-Excel producers create workbooks where workbookPart declares
        // sheet with empty r:id, but with name and sheetId. Content of such sheets
        // isn't loaded even if relationship part declares implicit relationship to
        // the worksheets, because workbook has explicit relationships with worksheet
        // part (ISO29500 12.3.23).
        //
        // If excel finds sheet in workbook without r:id, it adds empty sheet with
        // the specified name and so does XLibur.
        await TestHelper.LoadAndAssert(async wb =>
        {
            await Assert.That(wb.Worksheets.Count).IsEqualTo(3);

            // First sheet has r:id, so it keeps content
            await Assert.That(wb.Worksheet("Sheet1").Cell("A1").Value).IsEqualTo("Sheet1");

            // Second sheet doesn't have r:id, so it is empty after load.
            await Assert.That(wb.Worksheet("Sheet without relId").Cell("A1").Value).IsEqualTo(Blank.Value);

            // Third sheet doesn't have r:id and it contains pivot table that is not loaded.
            var ptSheet = wb.Worksheet("Pivot Sheet without relId");
            await Assert.That(ptSheet.Cell("A1").Value).IsEqualTo(Blank.Value);
            await Assert.That(ptSheet.PivotTables.Any()).IsFalse();
        }, @"TryToLoad\SheetsWithoutRelId.xlsx");
    }

    [Test]
    public async Task CanLoadDialogSheet()
    {
        // Workbook can reference multiple different types of sheet, most common is worksheet,
        // but there is also possibility of referencing dialogSheet (basically VBA dialog).
        // dialogSheet is basically obsolete (from Excel 5.0), but still supported. Do not
        // crash when such sheet is encountered. Test file also contains pivot table, because
        // it originally crashed just before pivot table loading.
        await TestHelper.LoadAndAssert(async wb =>
        {
            // Dialog sheet
            await Assert.That(wb.UnsupportedSheets.Count).IsEqualTo(1);

            // Data and pivot sheets
            await Assert.That(wb.Worksheets.Count).IsEqualTo(2);
            await Assert.That(wb.Worksheet("Pivot").PivotTables.Contains("PivotTable1")).IsTrue();
        }, @"TryToLoad\DialogSheet.xlsx");
    }

    // https://github.com/ClosedXML/ClosedXML/issues/2619
    [Test]
    public async Task Can_load_google_sheets_file_with_table_and_autofilter_on_same_range()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\GoogleSheets\2619_exported-broken2.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();
        await Assert.That(ws.Tables.Count()).IsGreaterThanOrEqualTo(1);
    }
}
