using System;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.PivotTables;
/// <summary>
/// Test methods of interface <see cref="IXLPivotFields"/> implemented through <see cref="XLPivotTableAxis"/>.
/// </summary>
internal class XLPivotTableAxisTests
{
    #region IXLPivotFields methods

    #region Add

    [Test]
    public async Task Add_field_not_yet_in_table_adds_field_and_shared_items()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Count"),
            (1, 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);
        var internalPt = (XLPivotTable)pt;
        await Assert.That(internalPt.PivotFields[0].Items).IsEmpty();

        var idField = pt.RowLabels.Add("ID", "Item ID").AddSubtotal(XLSubtotalFunction.Automatic);

        await Assert.That(idField.SourceName).IsEqualTo("ID");
        await Assert.That(idField.CustomName).IsEqualTo("Item ID");
        await Assert.That(pt.RowLabels.Single().CustomName).IsEqualTo("Item ID");

        // Adds values and default aggregation func to items of the field
        var fieldItems = internalPt.PivotFields[0].Items;
        await Assert.That(fieldItems.Count).IsEqualTo(2);
        await Assert.That(fieldItems[0].ItemType).IsEqualTo(XLPivotItemType.Data);
        await Assert.That(fieldItems[0].ItemIndex).IsEqualTo(0);
        await Assert.That(fieldItems[1].ItemType).IsEqualTo(XLPivotItemType.Default);
    }

    [Test]
    public async Task Same_field_cant_be_added_twice_to_same_axis()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Count"),
            (1, 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);
        pt.RowLabels.Add("ID", "Item ID");

        var ex = await Assert.That(() => pt.RowLabels.Add("ID", "Item ID")).Throws<InvalidOperationException>()!;
        await Assert.That(ex!.Message).IsEqualTo("Custom name 'Item ID' is already used.");
    }

    [Test]
    public async Task Add_field_must_exist_in_cache()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Count"),
            (1, 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);
        await Assert.That(() => pt.RowLabels.Add("ID", "Item ID")).ThrowsNothing();

        var ex = await Assert.That(() => pt.RowLabels.Add("nonexistent")).Throws<InvalidOperationException>()!;
        await Assert.That(ex!.Message).IsEqualTo("Field 'nonexistent' not found in pivot cache.");
    }

    #endregion

    #region Clear

    [Test]
    public async Task Clear_removes_all_fields_from_axis()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Color", "Count"),
            (1, "Blue", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);
        pt.RowLabels.Add("ID", "Item ID");
        pt.RowLabels.Add("Color", "Custom color");

        pt.RowLabels.Clear();

        await Assert.That(pt.RowLabels).IsEmpty();

        // Clear should also remove custom names and axis, otherwise there are problems loading
        // file with such remains in Excel.
        var internalPt = (XLPivotTable)pt;
        await Assert.That(internalPt.PivotFields[0].Name).IsNull();
        await Assert.That(internalPt.PivotFields[0].Axis).IsNull();
        await Assert.That(internalPt.PivotFields[1].Name).IsNull();
        await Assert.That(internalPt.PivotFields[1].Axis).IsNull();
    }

    #endregion

    #region Contains

    [Test]
    public async Task Contains_checks_whether_field_is_present()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Color", "Count"),
            (1, "Blue", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);
        var idField = pt.RowLabels.Add("ID", "Item ID");
        pt.ColumnLabels.Add("Color");

        await Assert.That(pt.RowLabels.Contains("id")).IsTrue();
        await Assert.That(pt.RowLabels.Contains(idField)).IsTrue();
        await Assert.That(pt.RowLabels.Contains("color")).IsFalse();
        await Assert.That(pt.RowLabels.Contains("nonexistent")).IsFalse();
    }

    #endregion

    #region Get(string sourceName)

    [Test]
    public async Task Get_field_by_source_name()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Color", "Count"),
            (1, "Blue", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);
        pt.RowLabels.Add("ID", "Item ID");
        pt.ColumnLabels.Add("Color");

        await Assert.That(pt.RowLabels.Get("id").SourceName).IsEqualTo("ID");
        var ex = await Assert.That(() => pt.RowLabels.Get("color")).Throws<KeyNotFoundException>()!;
        await Assert.That(ex!.Message).IsEqualTo("Field with source name 'color' not found in AxisRow.");
    }

    #endregion

    #region Get(int)

    [Test]
    public async Task Get_field_by_index()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Color", "Count"),
            (1, "Blue", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);
        pt.RowLabels.Add("ID", "Item ID");
        pt.ColumnLabels.Add("Color");

        await Assert.That(pt.RowLabels.Get(0).SourceName).IsEqualTo("ID");
        await Assert.That(() => pt.RowLabels.Get(-2)).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => pt.RowLabels.Get(1)).Throws<ArgumentOutOfRangeException>();
    }

    #endregion

    #region IndexOf

    [Test]
    public async Task IndexOf_finds_field_in_axis_by_source_name()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Color", "Count"),
            (1, "Blue", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);
        var idField = pt.RowLabels.Add("ID", "Item ID");
        pt.ColumnLabels.Add("Color");

        await Assert.That(pt.RowLabels.IndexOf("ID")).IsEqualTo(0);
        await Assert.That(pt.RowLabels.IndexOf(idField)).IsEqualTo(0);
        await Assert.That(pt.RowLabels.IndexOf("item id")).IsEqualTo(-1);
        await Assert.That(pt.RowLabels.IndexOf("Color")).IsEqualTo(-1);
    }

    #endregion

    #region Remove

    [Test]
    public async Task Remove_removes_field()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Color", "Count"),
            (1, "Blue", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);
        pt.RowLabels.Add("ID");

        pt.RowLabels.Remove("id");
        pt.RowLabels.Remove("ID"); // Doesnt throw on already removed.

        await Assert.That(pt.RowLabels).IsEmpty();
    }

    #endregion

    #endregion

    #region SetSubtotal

    [Test]
    public async Task SetSubtotal_adds_subtotal_when_enabled()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Count"),
            (1, 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);
        var field = pt.RowLabels.Add("ID");

        field.SetSubtotal(XLSubtotalFunction.Sum, true);

        await Assert.That(field.Subtotals).Contains(XLSubtotalFunction.Sum);
    }

    [Test]
    public async Task SetSubtotal_removes_subtotal_when_disabled()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Count"),
            (1, 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);
        var field = pt.RowLabels.Add("ID")
            .AddSubtotal(XLSubtotalFunction.Sum)
            .AddSubtotal(XLSubtotalFunction.Average);

        field.SetSubtotal(XLSubtotalFunction.Sum, false);

        await Assert.That(field.Subtotals).DoesNotContain(XLSubtotalFunction.Sum);
        await Assert.That(field.Subtotals).Contains(XLSubtotalFunction.Average);
    }

    [Test]
    public async Task SetSubtotal_can_remove_automatic_to_clear_subtotals()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Count"),
            (1, 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);
        var field = pt.RowLabels.Add("ID");

        // By default, new field has Automatic subtotal
        await Assert.That(field.Subtotals).Contains(XLSubtotalFunction.Automatic);

        field.SetSubtotal(XLSubtotalFunction.Automatic, false);

        await Assert.That(field.Subtotals).IsEmpty();
    }

    [Test]
    public async Task SetSubtotal_does_not_add_duplicate()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Count"),
            (1, 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);
        var field = pt.RowLabels.Add("ID")
            .SetSubtotal(XLSubtotalFunction.Sum, true)
            .SetSubtotal(XLSubtotalFunction.Sum, true);

        await Assert.That(field.Subtotals.Count(s => s == XLSubtotalFunction.Sum)).IsEqualTo(1);
    }

    [Test]
    public async Task Subtotals_exposes_automatic_when_present()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Count"),
            (1, 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);
        var field = pt.RowLabels.Add("ID");

        // Default field has Automatic
        await Assert.That(field.Subtotals).IsEquivalentTo(new[] { XLSubtotalFunction.Automatic });

        // Adding a custom subtotal still shows Automatic
        field.AddSubtotal(XLSubtotalFunction.Sum);
        await Assert.That(field.Subtotals).Contains(XLSubtotalFunction.Automatic);
        await Assert.That(field.Subtotals).Contains(XLSubtotalFunction.Sum);
    }

    [Test]
    public async Task SetSubtotal_on_filter_field_works()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Color", "Count"),
            (1, "Blue", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range!);
        pt.Values.Add("Count");
        var filterField = pt.ReportFilters.Add("Color");

        filterField.SetSubtotal(XLSubtotalFunction.Sum, true);
        await Assert.That(filterField.Subtotals).Contains(XLSubtotalFunction.Sum);

        filterField.SetSubtotal(XLSubtotalFunction.Sum, false);
        await Assert.That(filterField.Subtotals).DoesNotContain(XLSubtotalFunction.Sum);
    }

    #endregion
}
