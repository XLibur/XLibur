using System;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel;
using NUnit.Framework;

namespace XLibur.Tests.Excel.PivotTables;

/// <summary>
/// Test methods of interface <see cref="IXLPivotFields"/> implemented through <see cref="XLPivotTableAxis"/>.
/// </summary>
[TestFixture]
internal class XLPivotTableAxisTests
{
    #region IXLPivotFields methods

    #region Add

    [Test]
    public void Add_field_not_yet_in_table_adds_field_and_shared_items()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Count"),
            (1, 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        var internalPt = (XLPivotTable)pt;
        Assert.IsEmpty(internalPt.PivotFields[0].Items);

        var idField = pt.RowLabels.Add("ID", "Item ID").AddSubtotal(XLSubtotalFunction.Automatic);

        Assert.AreEqual("ID", idField.SourceName);
        Assert.AreEqual("Item ID", idField.CustomName);
        Assert.AreEqual("Item ID", pt.RowLabels.Single().CustomName);

        // Adds values and default aggregation func to items of the field
        var fieldItems = internalPt.PivotFields[0].Items;
        Assert.AreEqual(2, fieldItems.Count);
        Assert.AreEqual(XLPivotItemType.Data, fieldItems[0].ItemType);
        Assert.AreEqual(0, fieldItems[0].ItemIndex);
        Assert.AreEqual(XLPivotItemType.Default, fieldItems[1].ItemType);
    }

    [Test]
    public void Same_field_cant_be_added_twice_to_same_axis()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Count"),
            (1, 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        pt.RowLabels.Add("ID", "Item ID");

        var ex = Assert.Throws<InvalidOperationException>(() => pt.RowLabels.Add("ID", "Item ID"))!;
        Assert.AreEqual("Custom name 'Item ID' is already used.", ex.Message);
    }

    [Test]
    public void Add_field_must_exist_in_cache()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Count"),
            (1, 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        Assert.DoesNotThrow(() => pt.RowLabels.Add("ID", "Item ID"));

        var ex = Assert.Throws<InvalidOperationException>(() => pt.RowLabels.Add("nonexistent"))!;
        Assert.AreEqual("Field 'nonexistent' not found in pivot cache.", ex.Message);
    }

    #endregion

    #region Clear

    [Test]
    public void Clear_removes_all_fields_from_axis()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Color", "Count"),
            (1, "Blue", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        pt.RowLabels.Add("ID", "Item ID");
        pt.RowLabels.Add("Color", "Custom color");

        pt.RowLabels.Clear();

        Assert.IsEmpty(pt.RowLabels);

        // Clear should also remove custom names and axis, otherwise there are problems loading
        // file with such remains in Excel.
        var internalPt = (XLPivotTable)pt;
        Assert.Null(internalPt.PivotFields[0].Name);
        Assert.Null(internalPt.PivotFields[0].Axis);
        Assert.Null(internalPt.PivotFields[1].Name);
        Assert.Null(internalPt.PivotFields[1].Axis);
    }

    #endregion

    #region Contains

    [Test]
    public void Contains_checks_whether_field_is_present()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Color", "Count"),
            (1, "Blue", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        var idField = pt.RowLabels.Add("ID", "Item ID");
        pt.ColumnLabels.Add("Color");

        Assert.True(pt.RowLabels.Contains("id"));
        Assert.True(pt.RowLabels.Contains(idField));
        Assert.False(pt.RowLabels.Contains("color"));
        Assert.False(pt.RowLabels.Contains("nonexistent"));
    }

    #endregion

    #region Get(string sourceName)

    [Test]
    public void Get_field_by_source_name()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Color", "Count"),
            (1, "Blue", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        pt.RowLabels.Add("ID", "Item ID");
        pt.ColumnLabels.Add("Color");

        Assert.AreEqual("ID", pt.RowLabels.Get("id").SourceName);
        var ex = Assert.Throws<KeyNotFoundException>(() => pt.RowLabels.Get("color"))!;
        Assert.AreEqual("Field with source name 'color' not found in AxisRow.", ex.Message);
    }

    #endregion

    #region Get(int)

    [Test]
    public void Get_field_by_index()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Color", "Count"),
            (1, "Blue", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        pt.RowLabels.Add("ID", "Item ID");
        pt.ColumnLabels.Add("Color");

        Assert.AreEqual("ID", pt.RowLabels.Get(0).SourceName);
        Assert.Throws<ArgumentOutOfRangeException>(() => pt.RowLabels.Get(-2));
        Assert.Throws<ArgumentOutOfRangeException>(() => pt.RowLabels.Get(1));
    }

    #endregion

    #region IndexOf

    [Test]
    public void IndexOf_finds_field_in_axis_by_source_name()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Color", "Count"),
            (1, "Blue", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        var idField = pt.RowLabels.Add("ID", "Item ID");
        pt.ColumnLabels.Add("Color");

        Assert.AreEqual(0, pt.RowLabels.IndexOf("ID"));
        Assert.AreEqual(0, pt.RowLabels.IndexOf(idField));
        Assert.AreEqual(-1, pt.RowLabels.IndexOf("item id"));
        Assert.AreEqual(-1, pt.RowLabels.IndexOf("Color"));
    }

    #endregion

    #region Remove

    [Test]
    public void Remove_removes_field()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Color", "Count"),
            (1, "Blue", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        pt.RowLabels.Add("ID");

        pt.RowLabels.Remove("id");
        pt.RowLabels.Remove("ID"); // Doesnt throw on already removed.

        Assert.IsEmpty(pt.RowLabels);
    }

    #endregion

    #endregion

    #region SetSubtotal

    [Test]
    public void SetSubtotal_adds_subtotal_when_enabled()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Count"),
            (1, 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        var field = pt.RowLabels.Add("ID");

        field.SetSubtotal(XLSubtotalFunction.Sum, true);

        Assert.That(field.Subtotals, Does.Contain(XLSubtotalFunction.Sum));
    }

    [Test]
    public void SetSubtotal_removes_subtotal_when_disabled()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Count"),
            (1, 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        var field = pt.RowLabels.Add("ID")
            .AddSubtotal(XLSubtotalFunction.Sum)
            .AddSubtotal(XLSubtotalFunction.Average);

        field.SetSubtotal(XLSubtotalFunction.Sum, false);

        Assert.That(field.Subtotals, Does.Not.Contain(XLSubtotalFunction.Sum));
        Assert.That(field.Subtotals, Does.Contain(XLSubtotalFunction.Average));
    }

    [Test]
    public void SetSubtotal_can_remove_automatic_to_clear_subtotals()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Count"),
            (1, 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        var field = pt.RowLabels.Add("ID");

        // By default, new field has Automatic subtotal
        Assert.That(field.Subtotals, Does.Contain(XLSubtotalFunction.Automatic));

        field.SetSubtotal(XLSubtotalFunction.Automatic, false);

        Assert.That(field.Subtotals, Is.Empty);
    }

    [Test]
    public void SetSubtotal_does_not_add_duplicate()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Count"),
            (1, 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        var field = pt.RowLabels.Add("ID")
            .SetSubtotal(XLSubtotalFunction.Sum, true)
            .SetSubtotal(XLSubtotalFunction.Sum, true);

        Assert.That(field.Subtotals.Count(s => s == XLSubtotalFunction.Sum), Is.EqualTo(1));
    }

    [Test]
    public void Subtotals_exposes_automatic_when_present()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Count"),
            (1, 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        var field = pt.RowLabels.Add("ID");

        // Default field has Automatic
        Assert.That(field.Subtotals, Is.EquivalentTo(new[] { XLSubtotalFunction.Automatic }));

        // Adding a custom subtotal still shows Automatic
        field.AddSubtotal(XLSubtotalFunction.Sum);
        Assert.That(field.Subtotals, Does.Contain(XLSubtotalFunction.Automatic));
        Assert.That(field.Subtotals, Does.Contain(XLSubtotalFunction.Sum));
    }

    [Test]
    public void SetSubtotal_on_filter_field_works()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("ID", "Color", "Count"),
            (1, "Blue", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        pt.Values.Add("Count");
        var filterField = pt.ReportFilters.Add("Color");

        filterField.SetSubtotal(XLSubtotalFunction.Sum, true);
        Assert.That(filterField.Subtotals, Does.Contain(XLSubtotalFunction.Sum));

        filterField.SetSubtotal(XLSubtotalFunction.Sum, false);
        Assert.That(filterField.Subtotals, Does.Not.Contain(XLSubtotalFunction.Sum));
    }

    #endregion
}
