using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;

namespace XLibur.Tests.Excel.AutoFilters;

/// <summary>
/// A worksheet autofilter is modelled as something that hides rows, which is narrower than the
/// <c>filterColumn</c> element it is written to: there is no room in it for an <c>iconFilter</c>,
/// for the button attributes, for <c>extLst</c>, or for the three dozen <c>dynamicFilter</c>
/// types beyond the two averages. A column that has not been changed is written back from the
/// criteria it was loaded with, so none of those are dropped by a load and save.
/// </summary>
public class AutoFilterRoundTripTests
{
    /// <summary>
    /// <c>iconFilter</c> has no equivalent in the row-hiding model at all — there is no cell
    /// predicate for "shows a red light" without resolving the conditional format first.
    /// </summary>
    [Test]
    public async Task IconFilter_SurvivesARoundTrip()
    {
        var filterColumn = new FilterColumn { ColumnId = 0U };
        filterColumn.Append(new IconFilter { IconSet = IconSetValues.ThreeSymbols, IconId = 1U });

        var written = RoundTripFilterColumn(filterColumn).Elements<IconFilter>().Single();

        await Assert.That(written.IconSet!.InnerText).IsEqualTo("3Symbols");
        await Assert.That(written.IconId!.Value).IsEqualTo(1U);
    }

    [Test]
    public async Task ButtonAttributes_SurviveARoundTrip()
    {
        var filterColumn = new FilterColumn { ColumnId = 0U, HiddenButton = true, ShowButton = false };
        filterColumn.Append(new Top10 { Val = 3D });

        var written = RoundTripFilterColumn(filterColumn);

        await Assert.That(written.HiddenButton!.Value).IsTrue();
        await Assert.That(written.ShowButton!.Value).IsFalse();
    }

    /// <remarks>
    /// The column carries no criteria alongside it. <c>extLst</c> is one of the alternatives in
    /// the <c>CT_FilterColumn</c> choice rather than an addition to them, so a column holding
    /// both it and a <c>top10</c> is not something a valid file can contain.
    /// </remarks>
    [Test]
    public async Task FilterColumnExtensionList_SurvivesARoundTrip()
    {
        var extension = new Extension { Uri = "{89765432-10FE-DCBA-9876-543210FEDCBA}" };
        extension.AppendChild(new OpenXmlUnknownElement("x14", "filter",
            "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"));

        var extensionList = new ExtensionList();
        extensionList.Append(extension);

        var filterColumn = new FilterColumn { ColumnId = 0U };
        filterColumn.Append(extensionList);

        var written = RoundTripFilterColumn(filterColumn);
        var writtenExtension = written.GetFirstChild<ExtensionList>()!.Elements<Extension>().Single();

        await Assert.That(writtenExtension.Uri!.Value).IsEqualTo("{89765432-10FE-DCBA-9876-543210FEDCBA}");
    }

    /// <summary>
    /// <c>ST_DynamicFilterType</c> has around forty values, of which the model covers the two
    /// averages. Loading one of the others used to throw, because the token was looked up in a
    /// two-entry map; now it is carried through untouched.
    /// </summary>
    [Test]
    public async Task DynamicFilterTypeXLiburCannotEvaluate_LoadsAndSurvivesARoundTrip()
    {
        var filterColumn = new FilterColumn { ColumnId = 0U };
        filterColumn.Append(new DynamicFilter { Type = DynamicFilterValues.ThisMonth, Val = 45000D });

        var written = RoundTripFilterColumn(filterColumn).Elements<DynamicFilter>().Single();

        await Assert.That(written.Type!.InnerText).IsEqualTo("thisMonth");
        await Assert.That(written.Val!.Value).IsEqualTo(45000D);
    }

    /// <summary>
    /// The two dynamic types the model does cover still drive row hiding, so they are not simply
    /// being passed through as text.
    /// </summary>
    [Test]
    public async Task AboveAverageDynamicFilter_IsStillEvaluated()
    {
        var filterColumn = new FilterColumn { ColumnId = 1U };
        filterColumn.Append(new DynamicFilter { Type = DynamicFilterValues.AboveAverage, Val = 650D });

        using var input = new MemoryStream();
        CreateWorkbookWithAutoFilter(input, filterColumn);
        input.Position = 0;

        using var wb = new XLWorkbook(input);
        var autoFilter = wb.Worksheet("Data").AutoFilter;

        await Assert.That(autoFilter.Column(2).DynamicType).IsEqualTo(XLFilterDynamicType.AboveAverage);

        // Only Cake, at 800, is above the 650 average of the two rows.
        autoFilter.Reapply();
        await Assert.That(autoFilter.VisibleRows.Select(r => r.Cell(1).GetText()).ToList())
            .IsEquivalentTo(new[] { "Name", "Cake" });
    }

    /// <summary>
    /// Writing an untouched column back from its loaded criteria must not outlive the caller
    /// changing it, or an edit would be silently discarded on save.
    /// </summary>
    [Test]
    public async Task ChangingAColumn_WritesTheChangeRatherThanTheLoadedCriteria()
    {
        var filterColumn = new FilterColumn { ColumnId = 0U, HiddenButton = true };
        filterColumn.Append(new IconFilter { IconSet = IconSetValues.ThreeSymbols, IconId = 1U });

        using var input = new MemoryStream();
        CreateWorkbookWithAutoFilter(input, filterColumn);
        input.Position = 0;

        using var wb = new XLWorkbook(input);
        wb.Worksheet("Data").AutoFilter.Column(1).AddFilter("Cake");

        using var output = new MemoryStream();
        wb.SaveAs(output);

        var written = ReadFilterColumns(output).Single();

        await Assert.That(written.Elements<IconFilter>().Any()).IsFalse();
        await Assert.That(written.HiddenButton?.Value ?? false).IsFalse();
        await Assert.That(written.Elements<Filters>().Single().Elements<Filter>().Single().Val!.Value)
            .IsEqualTo("Cake");
    }

    /// <summary>
    /// A filter built through the API, rather than loaded, still writes what it always did.
    /// </summary>
    [Test]
    public async Task FilterBuiltThroughTheApi_IsWrittenFromItsOwnState()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").SetValue("Name");
        ws.Cell("A2").SetValue("Cookies");
        ws.Cell("A3").SetValue("Cake");

        ws.RangeUsed()!.SetAutoFilter().Column(1).AddFilter("Cake");

        using var output = new MemoryStream();
        wb.SaveAs(output);

        var written = ReadFilterColumns(output).Single();

        await Assert.That(written.ColumnId!.Value).IsEqualTo(0U);
        await Assert.That(written.Elements<Filters>().Single().Elements<Filter>().Single().Val!.Value)
            .IsEqualTo("Cake");
    }

    private static FilterColumn RoundTripFilterColumn(FilterColumn filterColumn)
    {
        using var input = new MemoryStream();
        CreateWorkbookWithAutoFilter(input, filterColumn);
        input.Position = 0;

        using var wb = new XLWorkbook(input);
        using var output = new MemoryStream();
        wb.SaveAs(output);

        return ReadFilterColumns(output).Single();
    }

    private static List<FilterColumn> ReadFilterColumns(MemoryStream saved)
    {
        saved.Position = 0;
        using var doc = SpreadsheetDocument.Open(saved, false);
        var worksheet = doc.WorkbookPart!.WorksheetParts.Single().Worksheet;

        return worksheet!.Elements<AutoFilter>().Single().Elements<FilterColumn>().ToList();
    }

    /// <summary>
    /// A one-sheet workbook whose autofilter carries the given column.
    /// </summary>
    private static void CreateWorkbookWithAutoFilter(Stream stream, FilterColumn filterColumn)
    {
        using var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        var workbookPart = doc.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Data" });

        var sheetData = new SheetData();
        sheetData.Append(CreateHeaderRow());
        sheetData.Append(CreateDataRow(2, "Cookies", 500));
        sheetData.Append(CreateDataRow(3, "Cake", 800));

        var autoFilter = new AutoFilter { Reference = "A1:B3" };
        autoFilter.Append(filterColumn);

        // Schema order puts autoFilter after sheetData.
        worksheetPart.Worksheet = new Worksheet(sheetData, autoFilter);
    }

    private static Row CreateHeaderRow()
    {
        var row = new Row { RowIndex = 1 };
        row.Append(CreateTextCell("A1", "Name"));
        row.Append(CreateTextCell("B1", "Revenue"));
        return row;
    }

    private static Row CreateDataRow(uint rowIndex, string name, double revenue)
    {
        var row = new Row { RowIndex = rowIndex };
        row.Append(CreateTextCell($"A{rowIndex}", name));
        row.Append(new Cell
        {
            CellReference = $"B{rowIndex}",
            DataType = CellValues.Number,
            CellValue = new CellValue(revenue),
        });

        return row;
    }

    private static Cell CreateTextCell(string reference, string value)
    {
        return new Cell
        {
            CellReference = reference,
            DataType = CellValues.String,
            CellValue = new CellValue(value),
        };
    }
}
