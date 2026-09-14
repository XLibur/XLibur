using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using TUnit.Assertions.Enums;
using XLibur.Excel;
using XLibur.Tests.Excel.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Tests.Excel.Worksheets;

/// <summary>
/// Every holder of text that names a sheet, against the three ways a sheet stops being called
/// <c>Data</c>: a rename, <c>IXLWorksheet.Delete()</c>, and <c>IXLWorksheets.Delete(name)</c>. Spec 55
/// task 1 executes the inventory with these.
/// <para>
/// Each test asserts what XLibur does <b>today</b>, wrong answers included, so that the task that
/// fixes one has to change the line that pins it. A comment names the task on every wrong answer.
/// Do not "fix" them here.
/// </para>
/// </summary>
public class SheetLifecycleCharacterizationTests
{
    private const string ChartsheetBook = @"Other\PivotTableReferenceFiles\ChartsheetAndPivotTable.xlsx";

    /// <summary>The three ways a sheet stops being called <c>Data</c>.</summary>
    public enum SheetEvent
    {
        /// <summary><c>IXLWorksheet.Name = "Renamed"</c>.</summary>
        Rename,

        /// <summary><c>IXLWorksheet.Delete()</c>.</summary>
        WorksheetDelete,

        /// <summary><c>IXLWorksheets.Delete("Data")</c>.</summary>
        CollectionDelete,
    }

    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.WorksheetDelete)]
    [Arguments(SheetEvent.CollectionDelete)]
    public async Task Cell_formula_text(SheetEvent sheetEvent)
    {
        using var wb = NewBook(out _, out var other);
        other.Cell("A1").FormulaA1 = "Data!A1*2";

        Apply(wb, sheetEvent);

        await Assert.That(other.Cell("A1").FormulaA1).IsEqualTo(Expect(sheetEvent,
            rename: "Renamed!A1*2",
            worksheetDelete: "#REF!*2", // was "Data!A1*2", naming the deleted sheet (task 4)
            collectionDelete: "#REF!*2"));
    }

    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.WorksheetDelete)]
    [Arguments(SheetEvent.CollectionDelete)]
    public async Task Cell_formula_value(SheetEvent sheetEvent)
    {
        using var wb = NewBook(out _, out var other);
        other.Cell("A1").FormulaA1 = "Data!A1*2";
        await Assert.That(other.Cell("A1").Value).IsEqualTo(10);

        Apply(wb, sheetEvent);

        // The calc engine is purged, and every formula marked dirty, on a rename and on either delete.
        // The collection delete did neither until task 2 made it the one door (D53).
        await Assert.That(other.Cell("A1").NeedsRecalculation).IsEqualTo(Expect(sheetEvent,
            rename: true,
            worksheetDelete: true,
            collectionDelete: true));
        await Assert.That(other.Cell("A1").Value).IsEqualTo(Expect<XLCellValue>(sheetEvent,
            rename: 10,
            worksheetDelete: XLError.CellReference,
            collectionDelete: XLError.CellReference));
    }

    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.WorksheetDelete)]
    [Arguments(SheetEvent.CollectionDelete)]
    public async Task Cell_formula_value_after_save_and_reload(SheetEvent sheetEvent)
    {
        using var ms = new MemoryStream();
        using (var wb = NewBook(out _, out var other))
        {
            other.Cell("A1").FormulaA1 = "Data!A1*2";
            _ = other.Cell("A1").Value;

            Apply(wb, sheetEvent);
            wb.SaveAs(ms);
        }

        using var reloaded = new XLWorkbook(ms);

        await Assert.That(reloaded.Worksheet("Other").Cell("A1").Value).IsEqualTo(Expect<XLCellValue>(sheetEvent,
            rename: 10,
            worksheetDelete: XLError.CellReference,
            collectionDelete: XLError.CellReference)); // was the stale 10, written to the file (D53)
    }

    [Test]
    [Arguments(SheetEvent.WorksheetDelete)]
    [Arguments(SheetEvent.CollectionDelete)]
    public async Task The_deleted_sheet_is_marked_deleted(SheetEvent sheetEvent)
    {
        using var wb = NewBook(out var data, out _);

        Apply(wb, sheetEvent);

        await Assert.That(((XLWorksheet)data).IsDeleted).IsEqualTo(Expect(sheetEvent,
            rename: false,
            worksheetDelete: true,
            collectionDelete: true)); // was false (D53)
    }

    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.WorksheetDelete)]
    [Arguments(SheetEvent.CollectionDelete)]
    public async Task Workbook_scoped_name(SheetEvent sheetEvent)
    {
        using var wb = NewBook(out _, out _);
        wb.DefinedNames.Add("W", "Data!$A$1");

        Apply(wb, sheetEvent);

        await Assert.That(wb.DefinedNames.Single().RefersTo).IsEqualTo(Expect(sheetEvent,
            rename: "Renamed!$A$1",
            worksheetDelete: "#REF!",
            collectionDelete: "#REF!")); // was left naming the deleted sheet (D53)
    }

    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.WorksheetDelete)]
    [Arguments(SheetEvent.CollectionDelete)]
    public async Task Sheet_scoped_name_on_another_sheet(SheetEvent sheetEvent)
    {
        using var wb = NewBook(out _, out var other);
        other.DefinedNames.Add("L", "Data!$A$1");

        Apply(wb, sheetEvent);

        await Assert.That(other.DefinedNames.Single().RefersTo).IsEqualTo(Expect(sheetEvent,
            rename: "Renamed!$A$1",
            worksheetDelete: "#REF!", // was "Data!$A$1": only workbook scope was walked (D54, task 3)
            collectionDelete: "#REF!")); // was "Data!$A$1" too (D54, task 3)
    }

    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.WorksheetDelete)]
    [Arguments(SheetEvent.CollectionDelete)]
    public async Task Validation_list_source(SheetEvent sheetEvent)
    {
        using var wb = NewBook(out _, out var other);
        other.Range("B1:B1").CreateDataValidation().List("=Data!$A$1:$A$3", true);

        Apply(wb, sheetEvent);

        // Not handled by any of the three (task 5, with spec 44). Written verbatim, as a reference
        // to another sheet, so it dangles in the saved file.
        await Assert.That(other.DataValidations.Single().MinValue).IsEqualTo("=Data!$A$1:$A$3");
        await Assert.That(SavedSheetXml(wb, "Other")).Contains("Data!$A$1:$A$3");
    }

    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.WorksheetDelete)]
    [Arguments(SheetEvent.CollectionDelete)]
    public async Task Conditional_format_formula(SheetEvent sheetEvent)
    {
        using var wb = NewBook(out _, out var other);
        other.Range("C1:C1").AddConditionalFormat().WhenIsTrue("=Data!$A$1>0")
            .Fill.SetBackgroundColor(XLColor.Red);

        Apply(wb, sheetEvent);

        // Not handled by any of the three (task 5).
        await Assert.That(other.ConditionalFormats.Single().Values[1].Value).IsEqualTo("Data!$A$1>0");
    }

    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.WorksheetDelete)]
    [Arguments(SheetEvent.CollectionDelete)]
    public async Task Chart_series_references(SheetEvent sheetEvent)
    {
        using var wb = NewBook(out _, out var other);
        var chart = other.Charts.Add(XLChartType.ColumnClustered);
        chart.Series.Add("S", "Data!$A$1:$A$3", "Data!$B$1:$B$3");

        Apply(wb, sheetEvent);

        // Not handled by any of the three (task 6), and written as they stand, so they dangle.
        var series = chart.Series.Single();
        await Assert.That(series.ValueReferences).IsEqualTo("Data!$A$1:$A$3");
        await Assert.That(series.CategoryReferences).IsEqualTo("Data!$B$1:$B$3");

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        var chartXml = ChartGoldenCorpus.FirstChartPartXml(ms);
        await Assert.That(chartXml).Contains("<c:f>Data!$A$1:$A$3</c:f>");
        await Assert.That(chartXml).Contains("<c:f>Data!$B$1:$B$3</c:f>");
    }

    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.WorksheetDelete)]
    [Arguments(SheetEvent.CollectionDelete)]
    public async Task Pivot_cache_source(SheetEvent sheetEvent)
    {
        using var wb = NewBook(out var data, out var other);
        data.Cell("D1").Value = "Name";
        data.Cell("E1").Value = "Amount";
        data.Cell("D2").Value = "a";
        data.Cell("E2").Value = 1;
        data.Cell("D3").Value = "b";
        data.Cell("E3").Value = 2;
        other.PivotTables.Add("pt", other.Cell("H1"), data.Range("D1:E3")).RowLabels.Add("Name");

        Apply(wb, sheetEvent);

        // Not handled by any of the three (task 6): the source no longer resolves, and the saved
        // cache still names the old sheet.
        await Assert.That(wb.PivotCaches.Single().SourceRange).IsNull();

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        await Assert.That(SavedPivotSourceSheet(ms)).IsEqualTo("Data");
    }

    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.WorksheetDelete)]
    [Arguments(SheetEvent.CollectionDelete)]
    public async Task Internal_hyperlink(SheetEvent sheetEvent)
    {
        using var wb = NewBook(out _, out var other);
        other.Cell("D1").SetHyperlink(new XLHyperlink("Data!A1"));

        Apply(wb, sheetEvent);

        // Not handled by any of the three (task 5): sheet-qualified text is kept as it was set.
        await Assert.That(other.Cell("D1").GetHyperlink().InternalAddress).IsEqualTo("Data!A1");
    }

    [Test]
    public async Task Print_area_formula_on_rename()
    {
        using var wb = NewBook(out var data, out _);
        var printAreas = (XLPrintAreas)data.PageSetup.PrintAreas;
        printAreas.FormulaReference = "OFFSET(Data!$A$1,0,0,3,2)";

        Apply(wb, SheetEvent.Rename);

        // Not handled (task 5). A delete takes the print area with its own sheet.
        await Assert.That(printAreas.FormulaReference).IsEqualTo("OFFSET(Data!$A$1,0,0,3,2)");
    }

    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.WorksheetDelete)]
    [Arguments(SheetEvent.CollectionDelete)]
    public async Task Sparkline_source(SheetEvent sheetEvent)
    {
        using var wb = NewBook(out _, out var other);
        other.SparklineGroups.Add("E1", "Data!A1:A3");

        Apply(wb, sheetEvent);

        // A sparkline holds a live range, turned into text at save by exactly this call.
        var expected = Expect(sheetEvent,
            rename: "Renamed!A1:A3",
            worksheetDelete: "#REF!A1:A3",
            collectionDelete: "#REF!A1:A3"); // was the deleted sheet's name, IsDeleted unset (D53)
        var written = other.SparklineGroups.SelectMany(g => g).Single()
            .SourceData.RangeAddress.ToString(XLReferenceStyle.A1, true);
        await Assert.That(written).IsEqualTo(expected);
        await Assert.That(SavedSheetXml(wb, "Other")).Contains($"<xm:f>{expected}</xm:f>");
    }

    [Test]
    [Arguments(SheetEvent.WorksheetDelete)]
    [Arguments(SheetEvent.CollectionDelete)]
    public async Task Deleting_a_sheet_moves_an_unsupported_sheet_behind_it(SheetEvent sheetEvent)
    {
        using var wb = OpenChartsheetBook();
        var chartsheet = wb.UnsupportedSheets.Single();
        var victim = wb.Worksheets.OrderBy(w => w.Position).First();
        var expected = chartsheet.Position > victim.Position ? chartsheet.Position - 1 : chartsheet.Position;

        if (sheetEvent == SheetEvent.WorksheetDelete)
            victim.Delete();
        else
            wb.Worksheets.Delete(victim.Name);

        await Assert.That(chartsheet.Position).IsEqualTo(expected);
    }

    /// <summary>
    /// The file has worksheets <c>Data</c> and <c>Pivot</c> and a chartsheet <c>Chart</c>, all
    /// authored in Excel.
    /// <para>
    /// This test was <c>Add_takes_the_name_of_an_unsupported_sheet</c> and asserted the wrong answer:
    /// the new worksheet took the chartsheet's name, so the saved file declared it twice. Spec 55
    /// task 7 refuses the name.
    /// </para>
    /// </summary>
    [Test]
    public async Task Add_refuses_the_name_of_an_unsupported_sheet()
    {
        using var wb = OpenChartsheetBook();

        await Assert.That(() => wb.Worksheets.Add("Chart")).Throws<ArgumentException>();
        await Assert.That(wb.Worksheets.Contains("Chart")).IsFalse();
    }

    /// <summary>
    /// This test was <c>Rename_takes_the_name_of_an_unsupported_sheet</c> and asserted the wrong
    /// answer. Spec 55 task 7 refuses the name, and the sheet keeps its own.
    /// </summary>
    [Test]
    public async Task Rename_refuses_the_name_of_an_unsupported_sheet()
    {
        using var wb = OpenChartsheetBook();
        var data = wb.Worksheet("Data");

        await Assert.That(() => data.Name = "Chart").Throws<ArgumentException>();
        await Assert.That(data.Name).IsEqualTo("Data");
        await Assert.That(wb.Worksheet("Data")).IsSameReferenceAs(data);
    }

    /// <summary>
    /// This test was <c>The_position_setter_does_not_move_an_unsupported_sheet</c> and asserted the
    /// wrong answer: the moved worksheet and the chartsheet shared a position. Spec 55 task 7 moves
    /// the chartsheet as the add and the delete already did.
    /// </summary>
    [Test]
    public async Task The_position_setter_moves_an_unsupported_sheet_too()
    {
        using var wb = OpenChartsheetBook();
        var chartsheet = wb.UnsupportedSheets.Single();
        var moved = wb.Worksheets.OrderByDescending(w => Math.Abs(w.Position - chartsheet.Position)).First();

        moved.Position = chartsheet.Position;

        var positions = wb.Worksheets.Select(w => w.Position).Append(chartsheet.Position).OrderBy(p => p).ToList();
        await Assert.That(positions).IsEquivalentTo(new[] { 1, 2, 3 }, CollectionOrdering.Matching);
    }

    private static XLWorkbook NewBook(out IXLWorksheet data, out IXLWorksheet other)
    {
        var wb = new XLWorkbook();
        data = wb.AddWorksheet("Data");
        other = wb.AddWorksheet("Other");
        data.Cell("A1").Value = 5;
        data.Cell("A2").Value = 2;
        data.Cell("A3").Value = 3;
        data.Cell("B1").Value = "x";
        data.Cell("B2").Value = "y";
        data.Cell("B3").Value = "z";
        return wb;
    }

    private static XLWorkbook OpenChartsheetBook()
        => new(TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(ChartsheetBook)));

    private static void Apply(XLWorkbook wb, SheetEvent sheetEvent)
    {
        switch (sheetEvent)
        {
            case SheetEvent.Rename:
                wb.Worksheet("Data").Name = "Renamed";
                break;
            case SheetEvent.WorksheetDelete:
                wb.Worksheet("Data").Delete();
                break;
            case SheetEvent.CollectionDelete:
                wb.Worksheets.Delete("Data");
                break;
        }
    }

    private static T Expect<T>(SheetEvent sheetEvent, T rename, T worksheetDelete, T collectionDelete)
        => sheetEvent switch
        {
            SheetEvent.Rename => rename,
            SheetEvent.WorksheetDelete => worksheetDelete,
            _ => collectionDelete,
        };

    private static string SavedSheetXml(XLWorkbook wb, string sheetName)
    {
        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var workbookPart = doc.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>().Single(s => s.Name?.Value == sheetName);
        var part = workbookPart.GetPartById(sheet.Id!.Value!);
        using var reader = new StreamReader(part.GetStream(FileMode.Open, FileAccess.Read));
        return reader.ReadToEnd();
    }

    private static string? SavedPivotSourceSheet(MemoryStream saved)
    {
        saved.Position = 0;
        using var doc = SpreadsheetDocument.Open(saved, false);
        return doc.WorkbookPart!.PivotTableCacheDefinitionParts.Single()
            .PivotCacheDefinition!.CacheSource!.WorksheetSource!.Sheet?.Value;
    }
}
