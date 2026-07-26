using System;
using System.Linq;
using XLibur.Excel;
using XLibur.Excel.PivotTables.Areas;
using System.Threading.Tasks;

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
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);

        var ex = await Assert.That(() => pt.Values.Add("Wrong field name")).Throws<ArgumentOutOfRangeException>();

        await Assert.That(ex).IsNotNull();
        await Assert.That(ex.Message).StartsWith("Field 'Wrong field name' is not in the fields of a pivot cache. Should be one of 'Name','Price'.");
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
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
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
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);

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
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
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
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
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
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);

        await Assert.That(() => pt.Values.Remove("NonExistent")).ThrowsNothing();
    }

    #endregion

    #endregion
}
