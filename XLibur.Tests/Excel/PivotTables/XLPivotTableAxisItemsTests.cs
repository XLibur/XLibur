using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;

namespace XLibur.Tests.Excel.PivotTables;

/// <summary>
/// A loaded file can keep the rendered <c>rowItems</c>/<c>colItems</c> of a pivot table's row or
/// column axis (<see cref="XLPivotTableAxis.Items"/>). An item names a value of each field on the
/// axis by position, so once the set of fields on the axis changes, every item on it is stale (#594).
/// </summary>
internal class XLPivotTableAxisItemsTests
{
    /// <summary>
    /// Excel saved this file. The row axis holds field <c>F1</c> alone, and its <c>rowItems</c>
    /// render 5 rows for it. The column axis holds the values sentinel alone.
    /// </summary>
    private const string PivotWithStyles = @"Other\Lion\PivotTables\PivotWithStyles.xlsx";

    [Test]
    [Property("Description", "#594: removing the only field from the row axis left the loaded rowItems behind, so the saved file named a field position the row axis no longer had")]
    public async Task Removing_the_only_field_from_the_rows_clears_the_row_axis_items()
    {
        using var saved = new MemoryStream();
        string sheetName;
        string pivotTableName;
        using (var wb = new XLWorkbook(TestHelper.GetStreamFromResource(
                   TestHelper.GetResourcePath(PivotWithStyles))))
        {
            var pt = (XLPivotTable)wb.Worksheets.SelectMany(ws => ws.PivotTables).First();
            sheetName = pt.Worksheet.Name;
            pivotTableName = pt.Name;
            var sourceName = pt.RowLabels.Single().SourceName;
            await Assert.That(pt.RowAxis.Items.Count).IsGreaterThan(0)
                .Because("the file must hold the rendered row items, or this test proves nothing");

            pt.RowLabels.Remove(sourceName);

            await Assert.That(pt.RowAxis.Items).IsEmpty()
                .Because("the row axis has no field left, so it has nothing left to render");
            wb.SaveAs(saved);
        }

        var xml = PivotTableXml(saved, sheetName, pivotTableName);
        await Assert.That(xml).DoesNotContain("<rowItems")
            .Because("the row axis holds no field now, so it has nothing left to render");
        await Assert.That(xml).Contains("<colItems")
            .Because("the column axis kept its field, so its own items must be left alone");

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var loaded = (XLPivotTable)reloaded.Worksheets.SelectMany(ws => ws.PivotTables).First();
        await Assert.That(loaded.RowAxis.Items).IsEmpty();
    }

    /// <summary>
    /// Excel saved this file. On <c>PivotTableSubtotals</c>, the row axis holds two fields, <c>Group</c>
    /// and <c>Name</c>, and its <c>rowItems</c> render 14 rows for them. The column axis holds one
    /// field, <c>Month</c>, on its own <c>colItems</c>.
    /// </summary>
    private const string CustomSubtotals = @"TryToLoad\LoadPivotTables.xlsx";

    private const string CustomSubtotalsSheet = "PivotTableSubtotals";

    [Test]
    [Property("Description", "#594: removing one of several fields from an axis is as stale as removing the only one, because every item names a value for each field on the axis")]
    public async Task Removing_one_of_several_fields_from_the_rows_clears_the_row_axis_items()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(TestHelper.GetStreamFromResource(
                   TestHelper.GetResourcePath(CustomSubtotals))))
        {
            var pt = (XLPivotTable)wb.Worksheet(CustomSubtotalsSheet).PivotTables.Single();
            await Assert.That(pt.RowLabels.Select(f => f.SourceName)).IsEquivalentTo(new[] { "Group", "Name" })
                .Because("the file must hold two fields on the rows, or this test proves nothing");
            await Assert.That(pt.RowAxis.Items.Count).IsGreaterThan(0)
                .Because("the file must hold the rendered row items, or this test proves nothing");
            var columnItemCount = pt.ColumnAxis.Items.Count;
            await Assert.That(columnItemCount).IsGreaterThan(0)
                .Because("the file must hold the rendered column items too, or this test proves nothing");

            pt.RowLabels.Remove("Group");

            await Assert.That(pt.RowAxis.Items).IsEmpty()
                .Because("a field is gone from the axis, so every item, naming a value for each field left on it, is stale");
            await Assert.That(pt.RowLabels.Single().SourceName).IsEqualTo("Name")
                .Because("the field that stays must still be on the axis");
            await Assert.That(pt.ColumnAxis.Items.Count).IsEqualTo(columnItemCount)
                .Because("the column axis kept its field, so its own items must be left alone");
            wb.SaveAs(saved);
        }

        var xml = PivotTableXml(saved, CustomSubtotalsSheet, CustomSubtotalsSheet);
        await Assert.That(xml).DoesNotContain("<rowItems")
            .Because("the row axis has one field left but no rendered rows for it, so it has nothing to write");
        await Assert.That(xml).Contains("<colItems")
            .Because("the column axis kept its field, so its own items must be left alone");
        await Assert.That(xml).Contains("<rowFields")
            .Because("the row axis kept a field, so it still needs a rowFields element");
    }

    [Test]
    [Property("Description", "#594: adding a field to an axis that already has loaded items would otherwise leave items with fewer x entries than the axis has fields")]
    public async Task Adding_a_field_to_the_rows_clears_row_axis_items_already_there()
    {
        // A table built in code writes no rowItems of its own (Excel lays the axis out itself), so
        // the row axis is put in the state a loaded file would give it -- one field, with rendered
        // items -- by injecting the items a save never writes on its own.
        using var saved = new MemoryStream();
        string sheetName;
        string pivotTableName;
        using (var wb = new XLWorkbook())
        {
            var pt = CreatePivotTable(wb);
            sheetName = pt.Worksheet.Name;
            pivotTableName = pt.Name;
            pt.RowLabels.Add("Name");
            pt.Values.Add("Sold");
            wb.SaveAs(saved);
        }

        InjectRowItems(saved, sheetName, pivotTableName);

        saved.Position = 0;
        using var wb2 = new XLWorkbook(saved);
        var loaded = (XLPivotTable)wb2.Worksheets.SelectMany(ws => ws.PivotTables).Single();
        await Assert.That(loaded.RowAxis.Items.Count).IsEqualTo(3)
            .Because("the injected row items must have loaded, or this test proves nothing");

        loaded.RowLabels.Add("Region");

        await Assert.That(loaded.RowAxis.Items).IsEmpty()
            .Because("a field joined the axis, so the old items, one x entry short, must not survive alongside it");
        await Assert.That(loaded.RowLabels.Select(f => f.SourceName)).IsEquivalentTo(new[] { "Name", "Region" });
    }

    private static XLPivotTable CreatePivotTable(XLWorkbook wb)
    {
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("Name", "Region", "Sold"),
            ("Pie", "North", 7),
            ("Cake", "South", 10),
        });
        var ptSheet = wb.AddWorksheet();
        return (XLPivotTable)ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);
    }

    /// <summary>
    /// Add a <c>rowItems</c> element rendering the row axis' one field, the way a loaded file keeps
    /// them and a table built in code never writes them.
    /// </summary>
    private static void InjectRowItems(Stream package, string sheetName, string pivotTableName)
    {
        package.Position = 0;
        using (var document = SpreadsheetDocument.Open(package, true))
        {
            var workbookPart = document.WorkbookPart!;
            var sheet = workbookPart.Workbook!.Sheets!.Elements<Sheet>().Single(s => s.Name == sheetName);
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
            var definition = worksheetPart.PivotTableParts
                .Select(part => part.PivotTableDefinition!)
                .Single(d => d.Name == pivotTableName);

            var rowItems = new RowItems { Count = 3U };
            rowItems.Append(new RowItem(new MemberPropertyIndex { Val = 0 }));
            rowItems.Append(new RowItem(new MemberPropertyIndex { Val = 1 }));
            rowItems.Append(new RowItem(new MemberPropertyIndex()) { ItemType = ItemValues.Default });

            definition.RowFields!.InsertAfterSelf(rowItems);
            definition.Save();
        }

        package.Position = 0;
    }

    /// <summary>The raw XML of one pivot table's definition, exactly as the save wrote it to the package.</summary>
    private static string PivotTableXml(Stream package, string sheetName, string pivotTableName)
    {
        string partUri;
        package.Position = 0;
        using (var document = SpreadsheetDocument.Open(package, false))
        {
            var workbookPart = document.WorkbookPart!;
            var sheet = workbookPart.Workbook!.Sheets!.Elements<Sheet>().Single(s => s.Name == sheetName);
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
            var pivotTablePart = worksheetPart.PivotTableParts
                .Single(part => part.PivotTableDefinition!.Name == pivotTableName);
            partUri = pivotTablePart.Uri.OriginalString.TrimStart('/');
        }

        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        using var entry = archive.GetEntry(partUri)!.Open();
        using var reader = new StreamReader(entry);
        return reader.ReadToEnd();
    }
}
