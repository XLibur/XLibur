using NUnit.Framework;
using XLibur.Excel;
using XLibur.Excel.Coordinates;

namespace XLibur.Tests.Excel.PivotTables;

/// <summary>
/// Tests for classes that implement <c>IXLPivotSource</c>.
/// </summary>
[TestFixture]
internal class XLPivotSourceTests
{
    [Test]
    public void Can_load_and_save_all_source_types()
    {
        // The test files contains all possible pivot cache sources. The output is mangled, but
        // Excel can open it and use refresh on each pivot table. External workbook is in the same
        // directory: PivotTable-AllSources-external-data.xlsx.
        // The pivot table that uses connection has a connection to the external workbook
        // PivotTable-AllSources-external-data.xlsx. The connection uses an absolute path, so it
        // needs to be updated according to real directory. Doesn't affect CI, because connection
        // is not actually used to get data.
        // Scenario doesn't throw on refresh, but it incomplete. The cache source is correct though.
        //
        // Open the workbook and click Pivot Table Analyze - Refresh - Refresh All. It shouldn't
        // report an error.
        TestHelper.LoadAndAssert(wb =>
        {
            Assert.That(wb.Worksheets.Count, Is.GreaterThan(0));
        }, @"Other\PivotTable\Sources\PivotTable-AllSources-input.xlsx");

        TestHelper.LoadSaveAndCompare(
            @"Other\PivotTable\Sources\PivotTable-AllSources-input.xlsx",
            @"Other\PivotTable\Sources\PivotTable-AllSources-output.xlsx");
    }

    #region TryGetSource - named range resolution

    [Test]
    public void TryGetSource_resolves_workbook_scoped_named_range()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = "Name";
        ws.Cell("B1").Value = "Value";
        ws.Cell("A2").Value = "Alpha";
        ws.Cell("B2").Value = 1;

        wb.DefinedNames.Add("PivotData", "Data!$A$1:$B$2");

        var source = new XLPivotSourceReference("PivotData");
        var result = source.TryGetSource(wb, out var sheet, out var sheetArea);

        Assert.That(result, Is.True);
        Assert.That(sheet!.Name, Is.EqualTo("Data"));
        Assert.That(sheetArea, Is.EqualTo(new XLSheetRange(1, 1, 2, 2)));
    }

    [Test]
    public void TryGetSource_prefers_table_over_named_range_with_same_name()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = "Name";
        ws.Cell("A2").Value = "Alpha";

        // Create a table named "Source"
        ws.Range("A1:A2").CreateTable("Source");

        // Create a defined name with the same identifier pointing to a different area
        wb.DefinedNames.Add("Source", "Data!$B$1:$B$5");

        var source = new XLPivotSourceReference("Source");
        var result = source.TryGetSource(wb, out var sheet, out var sheetArea);

        Assert.That(result, Is.True);
        Assert.That(sheet!.Name, Is.EqualTo("Data"));
        // Should resolve to table area (A1:A2), not defined name area (B1:B5)
        Assert.That(sheetArea!.Value.LeftColumn, Is.EqualTo(1));
        Assert.That(sheetArea!.Value.RightColumn, Is.EqualTo(1));
    }

    [Test]
    public void TryGetSource_returns_false_for_nonexistent_name()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Data");

        var source = new XLPivotSourceReference("NoSuchName");
        var result = source.TryGetSource(wb, out var sheet, out var sheetArea);

        Assert.That(result, Is.False);
        Assert.That(sheet, Is.Null);
        Assert.That(sheetArea, Is.Null);
    }

    [Test]
    public void TryGetSource_returns_false_for_named_range_referencing_deleted_sheet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = "Name";
        wb.DefinedNames.Add("PivotData", "Data!$A$1:$B$2");

        // Delete the sheet that the named range points to
        wb.Worksheets.Delete("Data");

        var source = new XLPivotSourceReference("PivotData");
        var result = source.TryGetSource(wb, out var sheet, out var sheetArea);

        Assert.That(result, Is.False);
        Assert.That(sheet, Is.Null);
    }

    [Test]
    public void TryGetSource_resolves_area_based_source()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "Header";

        var area = new XLBookArea("Sheet1", new XLSheetRange(1, 1, 5, 3));
        var source = new XLPivotSourceReference(area);
        var result = source.TryGetSource(wb, out var sheet, out var sheetArea);

        Assert.That(result, Is.True);
        Assert.That(sheet!.Name, Is.EqualTo("Sheet1"));
        Assert.That(sheetArea, Is.EqualTo(new XLSheetRange(1, 1, 5, 3)));
    }

    [Test]
    public void TryGetSource_area_returns_false_for_nonexistent_sheet()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1");

        var area = new XLBookArea("NonExistent", new XLSheetRange(1, 1, 5, 3));
        var source = new XLPivotSourceReference(area);
        var result = source.TryGetSource(wb, out var sheet, out var sheetArea);

        Assert.That(result, Is.False);
        Assert.That(sheet, Is.Null);
        Assert.That(sheetArea, Is.Null);
    }

    #endregion
}
