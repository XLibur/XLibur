using System;
using System.Data;
using System.Linq;
using XLibur.Excel;
using XLibur.Excel.Tables;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Tables;

public class TableNameValidationTests
{
    [Test]
    public async Task EmptyName_IsInvalid()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(TableNameValidator.IsValidTableName(string.Empty, ws, out _)).IsFalse();
    }

    [Test]
    public async Task WhitespaceName_IsInvalid()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(TableNameValidator.IsValidTableName("   ", ws, out _)).IsFalse();
    }

    [Test]
    public async Task NameStartingWithNumber_IsInvalid()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(TableNameValidator.IsValidTableName("1Table", ws, out var message)).IsFalse();
        await Assert.That(message).Contains("does not begin with a letter");
    }

    [Test]
    public async Task NameLongerThan255_IsInvalid()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var longName = new string('a', 256);
        await Assert.That(TableNameValidator.IsValidTableName(longName, ws, out var message)).IsFalse();
        await Assert.That(message).Contains("more than 255 characters");
    }

    [Test]
    public async Task NameWithSpaces_IsInvalid()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(TableNameValidator.IsValidTableName("Spaces in name", ws, out var message)).IsFalse();
        await Assert.That(message).Contains("cannot contain spaces");
    }

    [Test]
    [Arguments("A1")]
    [Arguments("May2019")]
    [Arguments("R1C2")]
    [Arguments("r3c2")]
    [Arguments("R2C33333")]
    [Arguments("RC")]
    public async Task CellAddress_IsInvalid(string name)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(TableNameValidator.IsValidTableName(name, ws, out var message)).IsFalse();
        await Assert.That(message).Contains("cell address");
    }

    [Test]
    public async Task ValidName_IsAccepted()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(TableNameValidator.IsValidTableName("MyTable", ws, out _)).IsTrue();
    }

    [Test]
    public async Task NameWithUnderscore_IsAccepted()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(TableNameValidator.IsValidTableName("_MyTable", ws, out _)).IsTrue();
    }

    [Test]
    public async Task DuplicateTableName_OnSameSheet_IsInvalid()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var t1 = ws.FirstCell().InsertTable(Enumerable.Range(1, 3).Select(i => new { Number = i }));
        await Assert.That(t1.Name).IsEqualTo("Table1");

        var t2 = ws.Cell("C1").InsertTable(Enumerable.Range(1, 3).Select(i => new { Number = i }));
        await Assert.That(t2.Name).IsEqualTo("Table2");

        var ex = await Assert.That(() => t2.Name = "TABLE1").Throws<ArgumentException>();
        await Assert.That(ex!.Message).Contains("already a table named");
    }

    [Test]
    public async Task CasingOnlyChange_DoesNotThrow()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var t1 = ws.FirstCell().InsertTable(Enumerable.Range(1, 3).Select(i => new { Number = i }));
        await Assert.That(t1.Name).IsEqualTo("Table1");
        await Assert.That(() => t1.Name = "TABLE1").ThrowsNothing();
        await Assert.That(t1.Name).IsEqualTo("TABLE1");
    }

    [Test]
    public async Task SpaceInName_ViaInsertTable_Throws()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var t1 = ws.FirstCell().InsertTable(Enumerable.Range(1, 3).Select(i => new { Number = i }));
        await Assert.That(() => t1.Name = "Table name with spaces").Throws<ArgumentException>();
    }

    [Test]
    public async Task CellAddressName_ViaInsertTableDataTable_Throws()
    {
        var dt = new DataTable("sheet1");
        dt.Columns.Add("Patient", typeof(string));
        dt.Rows.Add("David");

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        await Assert.That(() => ws.Cell(1, 1).InsertTable(dt, "A1")).Throws<ArgumentException>();
        await Assert.That(() => ws.Cell(1, 1).InsertTable(dt, "R1C2")).Throws<ArgumentException>();
        await Assert.That(() => ws.Cell(1, 1).InsertTable(dt, "r3c2")).Throws<ArgumentException>();
    }

    [Test]
    public async Task CellAddressName_ViaInsertTableEnumerable_Throws()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(() =>
            ws.FirstCell().InsertTable(Enumerable.Range(1, 3).Select(i => new { Number = i }), "A1")).Throws<ArgumentException>();
        await Assert.That(() =>
            ws.FirstCell().InsertTable(Enumerable.Range(1, 3).Select(i => new { Number = i }), "R1C2")).Throws<ArgumentException>();
    }

    [Test]
    public async Task TableName_ConflictsWithWorkbookDefinedName_IsInvalid()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet();
        _ = wb.AddWorksheet();

        wb.DefinedNames.Add("WorkbookDefinedName", "Sheet1!A1:A10");

        var t1 = ws1.FirstCell().InsertTable(Enumerable.Range(1, 3).Select(i => new { Number = i }));

        var ex = await Assert.That(() => t1.Name = "WorkbookDefinedName").Throws<ArgumentException>();
        await Assert.That(ex!.Message).Contains("unique across all defined names");
    }

    [Test]
    public async Task TableName_ConflictsWithWorksheetDefinedName_IsInvalid()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet();
        var ws2 = wb.AddWorksheet();

        ws2.DefinedNames.Add("SheetDefinedName", "Sheet2!A1:A10");

        var t1 = ws1.FirstCell().InsertTable(Enumerable.Range(1, 3).Select(i => new { Number = i }));

        var ex = await Assert.That(() => t1.Name = "SheetDefinedName").Throws<ArgumentException>();
        await Assert.That(ex!.Message).Contains("unique across all defined names");
    }
}
