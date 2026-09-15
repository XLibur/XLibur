using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using DataField = DocumentFormat.OpenXml.Spreadsheet.DataField;
using Item = DocumentFormat.OpenXml.Spreadsheet.Item;
using PivotField = DocumentFormat.OpenXml.Spreadsheet.PivotField;
using PivotTableDefinition = DocumentFormat.OpenXml.Spreadsheet.PivotTableDefinition;
using Sheet = DocumentFormat.OpenXml.Spreadsheet.Sheet;

namespace XLibur.Tests.Excel.PivotTables;

/// <summary>
/// A field can already have items when it goes on the rows, the columns or the filters: it was on an
/// axis before, a base item was set on it, or a loaded file kept them. It keeps them, in their order,
/// because a base item and the saved layout refer to an item by its position, and gains an item only
/// for each value it has none for (#534).
/// </summary>
internal class XLPivotFieldAxisItemsTests
{
    [Test]
    [Arguments("rows")]
    [Arguments("columns")]
    [Arguments("filters")]
    [Property("Description", "#534: AddFieldToAxis added an item for every cache value and every subtotal, so a field taken off an axis and put back had each item twice")]
    public async Task A_field_taken_off_the_rows_and_put_back_has_each_item_once(string axis)
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var pt = CreatePivotTable(wb);
            pt.RowLabels.Add("Name").AddSubtotal(XLSubtotalFunction.Sum);
            pt.Values.Add("Sold").ShowAsPercentageFrom("Name").And("Cake");
            var field = pt.PivotFields[0];
            await Assert.That(Describe(field)).IsEqualTo("x0,x1,x2,default,sum");

            pt.RowLabels.Remove("Name");
            await Assert.That(Describe(field)).IsEqualTo("x0,x1,x2,default,sum")
                .Because("taken off the rows, the field keeps its items, or this proves nothing");

            Fields(pt, axis).Add("Name");

            await Assert.That(Describe(field)).IsEqualTo("x0,x1,x2,default,sum")
                .Because("each value and each subtotal has one item, where it was");
            await Assert.That(pt.Values.Single().BaseItemValue).IsEqualTo("Cake");
            wb.SaveAs(saved);
        }

        await Assert.That(SavedItems(saved, "Data", "pt", 0)).IsEqualTo("x0,x1,x2,default,sum");
        await Assert.That(SavedBaseItem(saved, "Data", "pt")).IsEqualTo(1U);

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("Data").PivotTables.Single().Values.Single().BaseItemValue)
            .IsEqualTo("Cake");
    }

    /// <summary>
    /// Excel saved this file. On <c>OrderedPivotTable</c>, field 1 is on the rows and lists its items
    /// out of the pivot cache's order.
    /// </summary>
    private const string ItemsOutOfCacheOrder = @"TryToLoad\LoadPivotTables.xlsx";

    [Test]
    [Property("Description", "#534: a loaded field taken off the rows and put back had its items twice, the second time in the cache's order")]
    public async Task A_loaded_field_taken_off_the_rows_and_put_back_keeps_its_items_in_their_order()
    {
        using var saved = new MemoryStream();
        string sheetName;
        using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(ItemsOutOfCacheOrder)))
        using (var wb = new XLWorkbook(stream))
        {
            var pt = (XLPivotTable)wb.Worksheets.SelectMany(ws => ws.PivotTables)
                .Single(p => p.Name == "OrderedPivotTable");
            sheetName = pt.Worksheet.Name;
            var sourceName = pt.RowLabels.Single().SourceName;
            var field = pt.PivotFields[1];
            await Assert.That(Describe(field)).IsEqualTo("x2,x0,x3,x1,x4,default")
                .Because("the items must be out of the cache's order, or this proves nothing");

            pt.RowLabels.Remove(sourceName);
            pt.RowLabels.Add(sourceName);

            await Assert.That(Describe(field)).IsEqualTo("x2,x0,x3,x1,x4,default");
            wb.SaveAs(saved);
        }

        await Assert.That(SavedItems(saved, sheetName, "OrderedPivotTable", 1)).IsEqualTo("x2,x0,x3,x1,x4,default");
    }

    [Test]
    [Property("Description", "#534: setting a base item gave a field on no axis an item, and putting the field on an axis then gave it that item again")]
    public async Task A_field_given_a_base_item_on_no_axis_has_each_item_once_on_the_rows()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var pt = CreatePivotTable(wb);
            pt.Values.Add("Sold").ShowAsPercentageFrom("Name").And("Cake");
            var field = pt.PivotFields[0];
            await Assert.That(Describe(field)).IsEqualTo("x1")
                .Because("the base item gives the field on no axis its one item, or this proves nothing");

            pt.RowLabels.Add("Name");

            await Assert.That(Describe(field)).IsEqualTo("x1,x0,x2,default")
                .Because("the item the field had keeps its position, and the missing ones go after it");
            await Assert.That(pt.Values.Single().BaseItemValue).IsEqualTo("Cake");
            wb.SaveAs(saved);
        }

        await Assert.That(SavedItems(saved, "Data", "pt", 0)).IsEqualTo("x1,x0,x2,default");
        await Assert.That(SavedBaseItem(saved, "Data", "pt")).IsEqualTo(0U);

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("Data").PivotTables.Single().Values.Single().BaseItemValue)
            .IsEqualTo("Cake");
    }

    /// <summary>
    /// Excel saved this file. Field 2, <c>verbruik</c>, is a value field on no axis, and Excel kept its
    /// items: one for each of its 19 values, out of the pivot cache's order, and a default item.
    /// </summary>
    private const string FieldOnNoAxisWithItems = @"Other\PivotTable\pivottable_customfield.xlsx";

    [Test]
    [Property("Description", "#534: Excel keeps the items of a field on no axis, and putting such a field on an axis gave it each item twice")]
    public async Task A_loaded_field_on_no_axis_keeps_the_items_Excel_saved_when_it_goes_on_the_rows()
    {
        using var saved = new MemoryStream();
        string before;
        string sheetName;
        string pivotTableName;
        using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(FieldOnNoAxisWithItems)))
        using (var wb = new XLWorkbook(stream))
        {
            var pt = (XLPivotTable)wb.Worksheets.SelectMany(ws => ws.PivotTables).Single();
            sheetName = pt.Worksheet.Name;
            pivotTableName = pt.Name;
            var field = pt.PivotFields[2];
            await Assert.That(field.Axis).IsNull();
            await Assert.That(field.Items.Count).IsEqualTo(20)
                .Because("the field must have Excel's items, or this proves nothing");
            await Assert.That(field.Items[0].ItemIndex).IsEqualTo(11)
                .Because("the items must be out of the cache's order, or this proves nothing");
            before = Describe(field);

            pt.RowLabels.Add("verbruik");

            await Assert.That(Describe(field)).IsEqualTo(before);
            wb.SaveAs(saved);
        }

        await Assert.That(SavedItems(saved, sheetName, pivotTableName, 2)).IsEqualTo(before);
    }

    [Test]
    [Property("Description", "#534: a value the field has no item for is added before its subtotal items, which stay last")]
    public async Task A_value_the_cache_gained_while_the_field_was_on_no_axis_goes_before_the_subtotal_items()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb);
        pt.RowLabels.Add("Name");
        pt.RowLabels.Remove("Name");
        var field = pt.PivotFields[0];
        await Assert.That(Describe(field)).IsEqualTo("x0,x1,x2,default");

        // A second pivot table over the same cache adds a value to it through a base item.
        var other = pt.Worksheet.PivotTables.Add("other", pt.Worksheet.Cell("J1"), pt.PivotCache);
        other.Values.Add("Sold").ShowAsPercentageFrom("Name").And("Scone");

        pt.RowLabels.Add("Name");

        await Assert.That(Describe(field)).IsEqualTo("x0,x1,x2,x3,default");
    }

    private static XLPivotTable CreatePivotTable(XLWorkbook wb)
    {
        var ws = wb.AddWorksheet("Data");
        var range = ws.Cell("A1").InsertData(new object[]
        {
            ("Name", "Sold"),
            ("Pie", 7),
            ("Cake", 10),
            ("Tart", 3),
        });
        return (XLPivotTable)ws.PivotTables.Add("pt", ws.Cell("E1"), range!);
    }

    private static IXLPivotFields Fields(IXLPivotTable pt, string axis) => axis switch
    {
        "rows" => pt.RowLabels,
        "columns" => pt.ColumnLabels,
        "filters" => pt.ReportFilters,
        _ => throw new ArgumentOutOfRangeException(nameof(axis)),
    };

    /// <summary>
    /// The items of a field in order: <c>x</c> and the cache index for a value, otherwise the item type
    /// as the file writes it, such as <c>default</c>.
    /// </summary>
    private static string Describe(XLPivotTableField field)
        => string.Join(",", field.Items.Select(item => item.ItemIndex is { } index
            ? $"x{index}"
            : char.ToLowerInvariant(item.ItemType.ToString()[0]) + item.ItemType.ToString()[1..]));

    /// <summary>The items of one pivot field as the save wrote them, in the form of <see cref="Describe"/>.</summary>
    private static string SavedItems(Stream package, string sheetName, string pivotTableName, int fieldIndex)
        => ReadPivotDefinition(package, sheetName, pivotTableName, definition =>
        {
            var field = definition.PivotFields!.Elements<PivotField>().ElementAt(fieldIndex);
            return string.Join(",", field.Items?.Elements<Item>()
                .Select(item => item.Index is not null ? $"x{item.Index.Value}" : item.ItemType!.InnerText)
                ?? Enumerable.Empty<string>());
        });

    private static uint? SavedBaseItem(Stream package, string sheetName, string pivotTableName)
        => ReadPivotDefinition(package, sheetName, pivotTableName,
            definition => definition.DataFields!.Elements<DataField>().Single().BaseItem?.Value);

    private static T ReadPivotDefinition<T>(Stream package, string sheetName, string pivotTableName,
        Func<PivotTableDefinition, T> read)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Sheets!.Elements<Sheet>().Single(s => s.Name == sheetName);
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        var definition = worksheetPart.PivotTableParts
            .Select(part => part.PivotTableDefinition!)
            .Single(d => d.Name == pivotTableName);
        return read(definition);
    }
}
