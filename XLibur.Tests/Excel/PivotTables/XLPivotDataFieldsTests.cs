using System;
using System.Linq;
using XLibur.Excel;
using NUnit.Framework;
using XLibur.Excel.PivotTables.Areas;

namespace XLibur.Tests.Excel.PivotTables;

/// <summary>
/// Test methods of interface <see cref="IXLPivotValues"/> implemented through <see cref="XLPivotDataFields"/> class.
/// </summary>
internal class XLPivotDataFieldsTests
{
    #region IXLPivotValues methods

    #region Add

    [Test]
    public void Add_source_name_must_be_from_pivot_cache_field_names()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("Name", "Price"),
            ("Cake", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);

        var ex = Assert.Throws<ArgumentOutOfRangeException>(() => pt.Values.Add("Wrong field name"));

        Assert.NotNull(ex);
        Assert.That(ex.Message, Does.StartWith("Field 'Wrong field name' is not in the fields of a pivot cache. Should be one of 'Name','Price'."));
    }

    #endregion

    #region Clear

    [Test]
    public void Clear_removes_all_data_fields_from_pivot_table()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("Name", "Price", "Qty"),
            ("Cake", 10, 5),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        pt.Values.Add("Price");
        pt.Values.Add("Qty");

        Assert.That(pt.Values.Count(), Is.EqualTo(2));

        pt.Values.Clear();

        Assert.That(pt.Values.Count(), Is.EqualTo(0));
        Assert.That(pt.Values.Contains("Price"), Is.False);
        Assert.That(pt.Values.Contains("Qty"), Is.False);
    }

    [Test]
    public void Clear_on_empty_values_does_not_throw()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("Name", "Price"),
            ("Cake", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);

        Assert.That(() => pt.Values.Clear(), Throws.Nothing);
    }

    [Test]
    public void Clear_allows_re_adding_same_fields()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("Name", "Price"),
            ("Cake", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        pt.Values.Add("Price");

        pt.Values.Clear();
        var reAdded = pt.Values.Add("Price");

        Assert.That(reAdded, Is.Not.Null);
        Assert.That(pt.Values.Count(), Is.EqualTo(1));
    }

    #endregion

    #region Remove

    [Test]
    public void Remove_removes_specific_data_field()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("Name", "Price", "Qty"),
            ("Cake", 10, 5),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        pt.Values.Add("Price");
        pt.Values.Add("Qty");

        pt.Values.Remove("Price");

        Assert.That(pt.Values.Count(), Is.EqualTo(1));
        Assert.That(pt.Values.Contains("Price"), Is.False);
        Assert.That(pt.Values.Contains("Qty"), Is.True);
    }

    [Test]
    public void Remove_nonexistent_field_does_not_throw()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("Name", "Price"),
            ("Cake", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);

        Assert.That(() => pt.Values.Remove("NonExistent"), Throws.Nothing);
    }

    #endregion

    #endregion
}
