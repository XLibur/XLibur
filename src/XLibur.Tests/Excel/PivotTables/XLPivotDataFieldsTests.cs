using System;
using System.IO;
using System.Linq;
using XLibur.Excel;
using XLibur.Excel.PivotTables.Areas;
using System.Threading.Tasks;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.PivotTables;

/// <summary>
/// Test methods of interface <see cref="IXLPivotValues"/> implemented through <see cref="XLPivotDataFields"/> class.
/// </summary>
internal class XLPivotDataFieldsTests
{
    #region IXLPivotValues methods

    #region Add

    [Test]
    public async Task Add_source_name_must_be_from_pivot_cache_field_names()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("Name", "Price"),
            ("Cake", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);

        var ex = await Assert.That(() => pt.Values.Add("Wrong field name")).Throws<ArgumentOutOfRangeException>();

        await Assert.That(ex).IsNotNull();
        await Assert.That(ex!.Message).StartsWith("Field 'Wrong field name' is not in the fields of a pivot cache. Should be one of 'Name','Price'.");
    }

    #endregion

    #region Clear

    [Test]
    public async Task Clear_removes_all_data_fields_from_pivot_table()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("Name", "Price", "Qty"),
            ("Cake", 10, 5),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);
        pt.Values.Add("Price");
        pt.Values.Add("Qty");

        await Assert.That(pt.Values.Count()).IsEqualTo(2);

        pt.Values.Clear();

        await Assert.That(pt.Values.Count()).IsEqualTo(0);
        await Assert.That(pt.Values.Contains("Price")).IsFalse();
        await Assert.That(pt.Values.Contains("Qty")).IsFalse();
    }

    [Test]
    public async Task Clear_on_empty_values_does_not_throw()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("Name", "Price"),
            ("Cake", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);

        await Assert.That(() => pt.Values.Clear()).ThrowsNothing();
    }

    [Test]
    public async Task Clear_allows_re_adding_same_fields()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("Name", "Price"),
            ("Cake", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);
        pt.Values.Add("Price");

        pt.Values.Clear();
        var reAdded = pt.Values.Add("Price");

        await Assert.That(reAdded).IsNotNull();
        await Assert.That(pt.Values.Count()).IsEqualTo(1);
    }

    #endregion

    #region Remove

    [Test]
    public async Task Remove_removes_specific_data_field()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("Name", "Price", "Qty"),
            ("Cake", 10, 5),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);
        pt.Values.Add("Price");
        pt.Values.Add("Qty");

        pt.Values.Remove("Price");

        await Assert.That(pt.Values.Count()).IsEqualTo(1);
        await Assert.That(pt.Values.Contains("Price")).IsFalse();
        await Assert.That(pt.Values.Contains("Qty")).IsTrue();
    }

    [Test]
    public async Task Remove_nonexistent_field_does_not_throw()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("Name", "Price"),
            ("Cake", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);

        await Assert.That(() => pt.Values.Remove("NonExistent")).ThrowsNothing();
    }

    #endregion

    #endregion

    #region Values sentinel on an axis (#572)

    [Test]
    [Property("Description", "#572: Clear left the 'data' sentinel on the column axis, so a save wrote a colFields entry naming data fields the file no longer had")]
    public async Task Clear_takes_the_values_sentinel_off_the_column_axis()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb);
        pt.Values.Add("Price");
        pt.Values.Add("Qty");

        await Assert.That(SourceNames(pt.ColumnLabels)).IsEqualTo(XLConstants.PivotTable.ValuesSentinalLabel)
            .Because("a second value puts the sentinel on the column axis, or this test proves nothing");

        pt.Values.Clear();

        await Assert.That(SourceNames(pt.ColumnLabels)).IsEmpty();
        await Assert.That(SourceNames(pt.RowLabels)).IsEmpty();
        await Assert.That(pt.DataPosition).IsNull()
            .Because("dataPosition is derived from the axes, so it must go with the sentinel");
    }

    [Test]
    [Property("Description", "#572: the sentinel can sit on the row axis, and Clear must take it off whichever axis holds it")]
    public async Task Clear_takes_the_values_sentinel_off_the_row_axis()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb);

        // Placed on the rows by hand, so the second value finds an axis that already holds it
        // and leaves the columns alone.
        pt.RowLabels.Add(XLConstants.PivotTable.ValuesSentinalLabel);
        pt.Values.Add("Price");
        pt.Values.Add("Qty");

        await Assert.That(SourceNames(pt.RowLabels)).IsEqualTo(XLConstants.PivotTable.ValuesSentinalLabel)
            .Because("the sentinel must be on the rows, not the columns, or this test proves nothing");
        await Assert.That(SourceNames(pt.ColumnLabels)).IsEmpty();

        pt.Values.Clear();

        await Assert.That(SourceNames(pt.RowLabels)).IsEmpty();
        await Assert.That(SourceNames(pt.ColumnLabels)).IsEmpty();
        await Assert.That(pt.DataPosition).IsNull();
    }

    [Test]
    [Property("Description", "#572: Remove left the sentinel behind once the last value was gone")]
    public async Task Remove_takes_the_values_sentinel_off_once_the_last_value_is_gone()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb);
        pt.Values.Add("Price");
        pt.Values.Add("Qty");

        await Assert.That(SourceNames(pt.ColumnLabels)).IsEqualTo(XLConstants.PivotTable.ValuesSentinalLabel)
            .Because("a second value puts the sentinel on the column axis, or this test proves nothing");

        pt.Values.Remove("Price");

        await Assert.That(SourceNames(pt.ColumnLabels)).IsEqualTo(XLConstants.PivotTable.ValuesSentinalLabel)
            .Because("one value is left for the sentinel to name, so it stays where it was put");

        pt.Values.Remove("Qty");

        await Assert.That(SourceNames(pt.ColumnLabels)).IsEmpty();
        await Assert.That(SourceNames(pt.RowLabels)).IsEmpty();
        await Assert.That(pt.DataPosition).IsNull();
    }

    [Test]
    [Property("Description", "#572: a table that still has values must keep the sentinel exactly where it was")]
    public async Task Removing_one_of_three_values_leaves_the_sentinel_where_it_was()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb);
        pt.RowLabels.Add(XLConstants.PivotTable.ValuesSentinalLabel);
        pt.Values.Add("Price");
        pt.Values.Add("Qty");
        pt.Values.Add("Name");

        pt.Values.Remove("Price");

        await Assert.That(SourceNames(pt.RowLabels)).IsEqualTo(XLConstants.PivotTable.ValuesSentinalLabel)
            .Because("two values are left, so the sentinel stays on the axis it was put on");
        await Assert.That(SourceNames(pt.ColumnLabels)).IsEmpty()
            .Because("the sentinel must not be moved to the columns while the rows still hold it");
        await Assert.That(pt.DataPosition).IsEqualTo(0);
    }

    [Test]
    [Property("Description", "#586: a removal must never add a sentinel the caller took off by hand")]
    public async Task Removing_one_of_three_values_does_not_add_a_sentinel_the_caller_took_off()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb);
        pt.Values.Add("Price");
        pt.Values.Add("Qty");
        pt.Values.Add("Name");

        await Assert.That(SourceNames(pt.ColumnLabels)).IsEqualTo(XLConstants.PivotTable.ValuesSentinalLabel)
            .Because("a second value puts the sentinel on the column axis, or this test proves nothing");

        // Take the sentinel off by hand, the way the issue reaches the state: through the axis,
        // not through Values.
        pt.ColumnLabels.Remove(XLConstants.PivotTable.ValuesSentinalLabel);

        await Assert.That(SourceNames(pt.ColumnLabels)).IsEmpty()
            .Because("the sentinel must be off both axes, or this test proves nothing");
        await Assert.That(SourceNames(pt.RowLabels)).IsEmpty();

        pt.Values.Remove("Price");

        await Assert.That(SourceNames(pt.ColumnLabels)).IsEmpty()
            .Because("a removal may take a stale sentinel off, but must never impose one the caller did not ask for");
        await Assert.That(SourceNames(pt.RowLabels)).IsEmpty();
    }

    [Test]
    [Property("Description", "#572: an explicitly placed sentinel is kept while a value is left for it to name")]
    public async Task Clearing_down_to_one_value_keeps_a_sentinel_that_was_placed_by_hand()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb);
        pt.RowLabels.Add(XLConstants.PivotTable.ValuesSentinalLabel);
        pt.Values.Add("Price");

        await Assert.That(SourceNames(pt.RowLabels)).IsEqualTo(XLConstants.PivotTable.ValuesSentinalLabel);

        pt.Values.Add("Qty");
        pt.Values.Remove("Qty");

        await Assert.That(SourceNames(pt.RowLabels)).IsEqualTo(XLConstants.PivotTable.ValuesSentinalLabel)
            .Because("one value is left to name, and nothing asked for the sentinel to be taken off the rows");
    }

    [Test]
    [Property("Description", "#572: the saved file must not name a data field it does not have")]
    public async Task A_save_after_clearing_the_values_writes_no_field_of_minus_two()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var pt = CreatePivotTable(wb);
            pt.RowLabels.Add("Name");
            pt.Values.Add("Price");
            pt.Values.Add("Qty");
            pt.Values.Clear();
            wb.SaveAs(saved);
        }

        var xml = PivotTableXml(saved);

        await Assert.That(xml).DoesNotContain("<dataFields")
            .Because("every value was taken out, so there is no dataFields element");
        await Assert.That(xml).DoesNotContain("x=\"-2\"")
            .Because("a field of -2 names the data fields, and this file has none");
        await Assert.That(xml).Contains("<rowFields")
            .Because("the row field the table still has must survive the clear");
    }

    [Test]
    [Property("Description", "#572: a file saved after clearing the values must load back with no sentinel on either axis")]
    public async Task A_table_whose_values_were_cleared_round_trips()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var pt = CreatePivotTable(wb);
            pt.RowLabels.Add("Name");
            pt.Values.Add("Price");
            pt.Values.Add("Qty");
            pt.Values.Clear();
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var loaded = (XLPivotTable)reloaded.Worksheets.SelectMany(ws => ws.PivotTables).Single();

        await Assert.That(SourceNames(loaded.RowLabels)).IsEqualTo("Name");
        await Assert.That(SourceNames(loaded.ColumnLabels)).IsEmpty();
        await Assert.That(loaded.Values.Count()).IsEqualTo(0);
        await Assert.That(loaded.DataPosition).IsNull();
    }

    [Test]
    [Property("Description", "#572: a loaded table keeps the rendered items of its axis, and those name the data fields by index too")]
    public async Task Clearing_the_values_of_a_loaded_table_leaves_no_items_naming_a_data_field()
    {
        // The file's column axis holds the sentinel alone, and its colItems name the data fields
        // 1 and 2 through the 'i' attribute. Taking the sentinel off the axis is not enough: the
        // items would still name data fields the saved file no longer has.
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(TestHelper.GetStreamFromResource(
                   TestHelper.GetResourcePath(@"Other\Lion\PivotTables\PivotWithStyles.xlsx"))))
        {
            var pt = (XLPivotTable)wb.Worksheets.SelectMany(ws => ws.PivotTables).First();
            await Assert.That(SourceNames(pt.ColumnLabels)).IsEqualTo(XLConstants.PivotTable.ValuesSentinalLabel)
                .Because("the file must hold the sentinel on the column axis, or this test proves nothing");
            await Assert.That(pt.ColumnAxis.Items.Count).IsEqualTo(3)
                .Because("the file must hold the rendered column items, or this test proves nothing");

            pt.Values.Clear();
            wb.SaveAs(saved);
        }

        var xml = PivotTableXml(saved);

        await Assert.That(xml).DoesNotContain("<dataFields");
        await Assert.That(xml).DoesNotContain("x=\"-2\"");
        await Assert.That(xml).DoesNotContain("<colItems")
            .Because("the column axis holds no field now, so it has nothing left to render");
        await Assert.That(xml).Contains("<rowItems")
            .Because("the row axis kept its field, so its own items must be left alone");

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var loaded = (XLPivotTable)reloaded.Worksheets.SelectMany(ws => ws.PivotTables).First();

        await Assert.That(SourceNames(loaded.ColumnLabels)).IsEmpty();
        await Assert.That(loaded.ColumnAxis.Items).IsEmpty();
        await Assert.That(loaded.Values.Count()).IsEqualTo(0);
    }

    private static XLPivotTable CreatePivotTable(XLWorkbook wb)
    {
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("Name", "Price", "Qty"),
            ("Cake", 10, 5),
            ("Pie", 20, 7),
        });
        var ptSheet = wb.AddWorksheet();
        return (XLPivotTable)ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);
    }

    private static string SourceNames(IXLPivotFields fields)
    {
        return string.Join(",", fields.Select(f => f.SourceName));
    }

    private static string PivotTableXml(Stream package) => package.ReadPartUnder("xl/pivotTables/pivotTable");

    #endregion
}
