using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using DataField = DocumentFormat.OpenXml.Spreadsheet.DataField;
using Item = DocumentFormat.OpenXml.Spreadsheet.Item;
using ItemValues = DocumentFormat.OpenXml.Spreadsheet.ItemValues;
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
            await Assert.That(Describe(field)).IsEqualTo("x0,x1,x2,sum");

            pt.RowLabels.Remove("Name");
            await Assert.That(Describe(field)).IsEqualTo("x0,x1,x2,sum")
                .Because("taken off the rows, the field keeps its items, or this proves nothing");

            Fields(pt, axis).Add("Name");

            await Assert.That(Describe(field)).IsEqualTo("x0,x1,x2,sum")
                .Because("each value and each subtotal has one item, where it was");
            await Assert.That(pt.Values.Single().BaseItemValue).IsEqualTo("Cake");
            wb.SaveAs(saved);
        }

        await Assert.That(SavedItems(saved, "Data", "pt", 0)).IsEqualTo("x0,x1,x2,sum");
        await Assert.That(SavedBaseItem(saved, "Data", "pt")).IsEqualTo(1U);

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("Data").PivotTables.Single().Values.Single().BaseItemValue)
            .IsEqualTo("Cake");
    }

    [Test]
    [Property("Description", "#550: a field built in code starts with the automatic subtotal, so a custom subtotal added to it gave it a default item next to its custom ones")]
    public async Task A_field_with_custom_subtotals_has_no_default_item_on_the_rows()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var pt = CreatePivotTable(wb);
            var name = pt.RowLabels.Add("Name")
                .AddSubtotal(XLSubtotalFunction.Sum)
                .AddSubtotal(XLSubtotalFunction.Average);
            pt.Values.Add("Sold");
            await Assert.That(name.Subtotals).Contains(XLSubtotalFunction.Automatic)
                .Because("the field keeps the automatic subtotal, as Excel's own file does, or this proves nothing");

            await Assert.That(Describe(pt.PivotFields[0])).IsEqualTo("x0,x1,x2,sum,avg");
            wb.SaveAs(saved);
        }

        await Assert.That(SavedItems(saved, "Data", "pt", 0)).IsEqualTo("x0,x1,x2,sum,avg");
        await Assert.That(SavedSubtotals(saved, "Data", "pt", 0)).IsEqualTo("sumSubtotal=1,avgSubtotal=1")
            .Because("Excel writes no defaultSubtotal on a field with custom subtotals");
    }

    [Test]
    [Property("Description", "#550: the automatic subtotal set again while a custom subtotal is set is ignored, so it gives the field no default item")]
    public async Task The_automatic_subtotal_added_after_a_custom_one_gives_no_default_item()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb);
        pt.RowLabels.Add("Name")
            .SetSubtotal(XLSubtotalFunction.Automatic, false)
            .AddSubtotal(XLSubtotalFunction.Sum)
            .AddSubtotal(XLSubtotalFunction.Automatic);

        await Assert.That(Describe(pt.PivotFields[0])).IsEqualTo("x0,x1,x2,sum");
    }

    [Test]
    [Property("Description", "#550: with its last custom subtotal removed, the field's automatic subtotal applies again, with its default item")]
    public async Task A_field_whose_last_custom_subtotal_is_removed_has_its_default_item_again()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb);
        var name = pt.RowLabels.Add("Name").AddSubtotal(XLSubtotalFunction.Sum);

        name.SetSubtotal(XLSubtotalFunction.Sum, false);

        await Assert.That(Describe(pt.PivotFields[0])).IsEqualTo("x0,x1,x2,default");
    }

    [Test]
    [Property("Description", "#550: a field whose only subtotal is the automatic one still has its default item")]
    public async Task A_field_with_only_the_automatic_subtotal_has_a_default_item()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var pt = CreatePivotTable(wb);
            pt.RowLabels.Add("Name");
            pt.Values.Add("Sold");

            await Assert.That(Describe(pt.PivotFields[0])).IsEqualTo("x0,x1,x2,default");
            wb.SaveAs(saved);
        }

        await Assert.That(SavedItems(saved, "Data", "pt", 0)).IsEqualTo("x0,x1,x2,default");
        await Assert.That(SavedSubtotals(saved, "Data", "pt", 0)).IsEqualTo("");
    }

    /// <summary>
    /// Excel saved this file. On <c>PivotTableSubtotals</c>, field 0 is on the rows with the custom
    /// subtotals sum, count and average, and Excel wrote no <c>defaultSubtotal</c> for it and no
    /// default item. Field 1 has the automatic subtotal, and field 4 none.
    /// </summary>
    private const string CustomSubtotals = @"TryToLoad\LoadPivotTables.xlsx";

    private const string CustomSubtotalsSheet = "PivotTableSubtotals";

    [Test]
    [Arguments("rows")]
    [Arguments("columns")]
    [Arguments("filters")]
    [Property("Description", "#550: Excel leaves defaultSubtotal out on a field with custom subtotals, so the load gave the field the automatic subtotal too, and putting it on an axis added a default item. Excel, moving this field to the filters, the columns, or off the rows and back, writes the items it had and no default item")]
    public async Task Excels_field_with_custom_subtotals_taken_off_the_rows_and_put_on_an_axis_gets_no_default_item(string axis)
    {
        using var original = ReadResource(CustomSubtotals);
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(original))
        {
            var pt = (XLPivotTable)wb.Worksheet(CustomSubtotalsSheet).PivotTables.Single();
            var field = pt.PivotFields[0];
            await Assert.That(Describe(field)).IsEqualTo("x1,x0,sum,countA,avg");
            await Assert.That(field.Subtotals).Contains(XLSubtotalFunction.Automatic)
                .Because("the load reads the absent defaultSubtotal as true, or this proves nothing");

            var sourceName = pt.RowLabels.Get(0).SourceName;
            pt.RowLabels.Remove(sourceName);
            Fields(pt, axis).Add(sourceName);

            await Assert.That(Describe(field)).IsEqualTo("x1,x0,sum,countA,avg");
            wb.SaveAs(saved);
        }

        await Assert.That(SavedItems(saved, CustomSubtotalsSheet, CustomSubtotalsSheet, 0))
            .IsEqualTo(SavedItems(original, CustomSubtotalsSheet, CustomSubtotalsSheet, 0));
        await Assert.That(SavedSubtotals(saved, CustomSubtotalsSheet, CustomSubtotalsSheet, 0))
            .IsEqualTo(SavedSubtotals(original, CustomSubtotalsSheet, CustomSubtotalsSheet, 0));
    }

    [Test]
    [Property("Description", "#550: a load and save of Excel's file writes each field's items and subtotal attributes as Excel wrote them")]
    public async Task Excels_fields_with_custom_automatic_and_no_subtotals_save_as_Excel_wrote_them()
    {
        using var original = ReadResource(CustomSubtotals);
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(original))
            wb.SaveAs(saved);

        await Assert.That(SavedSubtotals(original, CustomSubtotalsSheet, CustomSubtotalsSheet, 0))
            .IsEqualTo("sumSubtotal=1,countASubtotal=1,avgSubtotal=1")
            .Because("Excel must have left defaultSubtotal out on the field, or this proves nothing");
        for (var i = 0; i < 5; ++i)
        {
            await Assert.That(SavedItems(saved, CustomSubtotalsSheet, CustomSubtotalsSheet, i))
                .IsEqualTo(SavedItems(original, CustomSubtotalsSheet, CustomSubtotalsSheet, i))
                .Because($"the items of field {i}");
            await Assert.That(SavedSubtotals(saved, CustomSubtotalsSheet, CustomSubtotalsSheet, i))
                .IsEqualTo(SavedSubtotals(original, CustomSubtotalsSheet, CustomSubtotalsSheet, i))
                .Because($"the subtotal attributes of field {i}");
        }
    }

    [Test]
    [Property("Description", "#550: an older XLibur saved a default item next to a field's custom subtotal items; putting the field on an axis now removes it")]
    public async Task A_default_item_saved_next_to_custom_subtotal_items_goes_when_the_field_is_put_on_an_axis()
    {
        using var original = ReadResource(CustomSubtotals);
        using var withDefaultItem = ReadResource(CustomSubtotals);
        EditPivotDefinition(withDefaultItem, CustomSubtotalsSheet, CustomSubtotalsSheet, definition =>
        {
            var items = definition.PivotFields!.Elements<PivotField>().First().Items!;
            items.AppendChild(new Item { ItemType = ItemValues.Default });
            items.Count = (uint)items.Elements<Item>().Count();
        });

        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(withDefaultItem))
        {
            var pt = (XLPivotTable)wb.Worksheet(CustomSubtotalsSheet).PivotTables.Single();
            var field = pt.PivotFields[0];
            await Assert.That(Describe(field)).IsEqualTo("x1,x0,sum,countA,avg,default")
                .Because("the field must have a default item next to its custom ones, or this proves nothing");

            var sourceName = pt.RowLabels.Get(0).SourceName;
            pt.RowLabels.Remove(sourceName);
            pt.RowLabels.Add(sourceName);

            await Assert.That(Describe(field)).IsEqualTo("x1,x0,sum,countA,avg");
            wb.SaveAs(saved);
        }

        await Assert.That(SavedItems(saved, CustomSubtotalsSheet, CustomSubtotalsSheet, 0))
            .IsEqualTo(SavedItems(original, CustomSubtotalsSheet, CustomSubtotalsSheet, 0));
    }

    /// <summary>Change the definition of one pivot table in a package, in place, and rewind the package.</summary>
    private static void EditPivotDefinition(Stream package, string sheetName, string pivotTableName,
        Action<PivotTableDefinition> edit)
    {
        package.Position = 0;
        using (var document = SpreadsheetDocument.Open(package, true))
        {
            var definition = FindPivotDefinition(document, sheetName, pivotTableName);
            edit(definition);
            definition.Save();
        }

        package.Position = 0;
    }

    /// <summary>A copy of a resource, at its start.</summary>
    private static MemoryStream ReadResource(string path)
    {
        var copy = new MemoryStream();
        using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(path)))
            stream.CopyTo(copy);

        copy.Position = 0;
        return copy;
    }

    /// <summary>
    /// The subtotal attributes of one pivot field as the file has them, in the file's order, such as
    /// <c>sumSubtotal=1,avgSubtotal=1</c>.
    /// </summary>
    private static string SavedSubtotals(Stream package, string sheetName, string pivotTableName, int fieldIndex)
        => ReadPivotDefinition(package, sheetName, pivotTableName, definition =>
        {
            var field = definition.PivotFields!.Elements<PivotField>().ElementAt(fieldIndex);
            return string.Join(",", field.GetAttributes()
                .Where(attribute => attribute.LocalName.EndsWith("Subtotal", StringComparison.Ordinal))
                .Select(attribute => $"{attribute.LocalName}={attribute.Value}"));
        });

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
        return read(FindPivotDefinition(document, sheetName, pivotTableName));
    }

    private static PivotTableDefinition FindPivotDefinition(SpreadsheetDocument document, string sheetName,
        string pivotTableName)
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Sheets!.Elements<Sheet>().Single(s => s.Name == sheetName);
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        return worksheetPart.PivotTableParts
            .Select(part => part.PivotTableDefinition!)
            .Single(d => d.Name == pivotTableName);
    }
}
