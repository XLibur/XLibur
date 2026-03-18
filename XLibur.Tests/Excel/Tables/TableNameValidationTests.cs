using System;
using System.Data;
using System.Linq;
using XLibur.Excel;
using NUnit.Framework;
using XLibur.Excel.Tables;

namespace XLibur.Tests.Excel.Tables;

[TestFixture]
public class TableNameValidationTests
{
    [Test]
    public void EmptyName_IsInvalid()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        Assert.That(TableNameValidator.IsValidTableName(string.Empty, ws, out _), Is.False);
    }

    [Test]
    public void WhitespaceName_IsInvalid()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        Assert.That(TableNameValidator.IsValidTableName("   ", ws, out _), Is.False);
    }

    [Test]
    public void NameStartingWithNumber_IsInvalid()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        Assert.That(TableNameValidator.IsValidTableName("1Table", ws, out var message), Is.False);
        Assert.That(message, Does.Contain("does not begin with a letter"));
    }

    [Test]
    public void NameLongerThan255_IsInvalid()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var longName = new string('a', 256);
        Assert.That(TableNameValidator.IsValidTableName(longName, ws, out var message), Is.False);
        Assert.That(message, Does.Contain("more than 255 characters"));
    }

    [Test]
    public void NameWithSpaces_IsInvalid()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        Assert.That(TableNameValidator.IsValidTableName("Spaces in name", ws, out var message), Is.False);
        Assert.That(message, Does.Contain("cannot contain spaces"));
    }

    [TestCase("A1")]
    [TestCase("May2019")]
    [TestCase("R1C2")]
    [TestCase("r3c2")]
    [TestCase("R2C33333")]
    [TestCase("RC")]
    public void CellAddress_IsInvalid(string name)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        Assert.That(TableNameValidator.IsValidTableName(name, ws, out var message), Is.False);
        Assert.That(message, Does.Contain("cell address"));
    }

    [Test]
    public void ValidName_IsAccepted()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        Assert.That(TableNameValidator.IsValidTableName("MyTable", ws, out _), Is.True);
    }

    [Test]
    public void NameWithUnderscore_IsAccepted()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        Assert.That(TableNameValidator.IsValidTableName("_MyTable", ws, out _), Is.True);
    }

    [Test]
    public void DuplicateTableName_OnSameSheet_IsInvalid()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var t1 = ws.FirstCell().InsertTable(Enumerable.Range(1, 3).Select(i => new { Number = i }));
        Assert.That(t1.Name, Is.EqualTo("Table1"));

        var t2 = ws.Cell("C1").InsertTable(Enumerable.Range(1, 3).Select(i => new { Number = i }));
        Assert.That(t2.Name, Is.EqualTo("Table2"));

        var ex = Assert.Throws<ArgumentException>(() => t2.Name = "TABLE1");
        Assert.That(ex!.Message, Does.Contain("already a table named"));
    }

    [Test]
    public void CasingOnlyChange_DoesNotThrow()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var t1 = ws.FirstCell().InsertTable(Enumerable.Range(1, 3).Select(i => new { Number = i }));
        Assert.That(t1.Name, Is.EqualTo("Table1"));
        Assert.DoesNotThrow(() => t1.Name = "TABLE1");
        Assert.That(t1.Name, Is.EqualTo("TABLE1"));
    }

    [Test]
    public void SpaceInName_ViaInsertTable_Throws()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var t1 = ws.FirstCell().InsertTable(Enumerable.Range(1, 3).Select(i => new { Number = i }));
        Assert.Throws<ArgumentException>(() => t1.Name = "Table name with spaces");
    }

    [Test]
    public void CellAddressName_ViaInsertTableDataTable_Throws()
    {
        var dt = new DataTable("sheet1");
        dt.Columns.Add("Patient", typeof(string));
        dt.Rows.Add("David");

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        Assert.Throws<ArgumentException>(() => ws.Cell(1, 1).InsertTable(dt, "A1"));
        Assert.Throws<ArgumentException>(() => ws.Cell(1, 1).InsertTable(dt, "R1C2"));
        Assert.Throws<ArgumentException>(() => ws.Cell(1, 1).InsertTable(dt, "r3c2"));
    }

    [Test]
    public void CellAddressName_ViaInsertTableEnumerable_Throws()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        Assert.Throws<ArgumentException>(() =>
            ws.FirstCell().InsertTable(Enumerable.Range(1, 3).Select(i => new { Number = i }), "A1"));
        Assert.Throws<ArgumentException>(() =>
            ws.FirstCell().InsertTable(Enumerable.Range(1, 3).Select(i => new { Number = i }), "R1C2"));
    }

    [Test]
    public void TableName_ConflictsWithWorkbookDefinedName_IsInvalid()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet();
        _ = wb.AddWorksheet();

        wb.DefinedNames.Add("WorkbookDefinedName", "Sheet1!A1:A10");

        var t1 = ws1.FirstCell().InsertTable(Enumerable.Range(1, 3).Select(i => new { Number = i }));

        var ex = Assert.Throws<ArgumentException>(() => t1.Name = "WorkbookDefinedName");
        Assert.That(ex!.Message, Does.Contain("unique across all defined names"));
    }

    [Test]
    public void TableName_ConflictsWithWorksheetDefinedName_IsInvalid()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet();
        var ws2 = wb.AddWorksheet();

        ws2.DefinedNames.Add("SheetDefinedName", "Sheet2!A1:A10");

        var t1 = ws1.FirstCell().InsertTable(Enumerable.Range(1, 3).Select(i => new { Number = i }));

        var ex = Assert.Throws<ArgumentException>(() => t1.Name = "SheetDefinedName");
        Assert.That(ex!.Message, Does.Contain("unique across all defined names"));
    }
}
