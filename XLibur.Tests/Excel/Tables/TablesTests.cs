using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Attributes;
using XLibur.Excel;
using XLibur.Excel.Exceptions;
using XLibur.Excel.Tables;
using XLibur.Extensions;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Tables;

public class TablesTests
{
    public class TestObjectWithoutAttributes
    {
        public string Column1 { get; set; }

        public string Column2 { get; set; }
    }

    public class TestObjectWithAttributes
    {
        public int UnOrderedColumn { get; set; }

        [XLColumn(Header = "SecondColumn", Order = 1)]
        public string Column1 { get; set; }

        [XLColumn(Header = "FirstColumn", Order = 0)]
        public string Column2 { get; set; }

        [XLColumn(Header = "SomeFieldNotProperty", Order = 2)]
        public int MyField;
    }

    [Test]
    public async Task CanSaveTableCreatedFromEmptyDataTable()
    {
        var dt = new DataTable("sheet1");
        dt.Columns.Add("col1", typeof(string));
        dt.Columns.Add("col2", typeof(double));

        using var wb = new XLWorkbook();
        wb.AddWorksheet(dt);

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        await Assert.That(ms.Length).IsGreaterThan(0);
        await Assert.That(wb.Worksheets.First().Tables.Count()).IsEqualTo(1);
    }

    [Test]
    public async Task PreventAddingOfEmptyDataTable()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var dt = new DataTable();
        var table = ws.FirstCell().InsertTable(dt);

        await Assert.That(table).IsNull();
    }

    [Test]
    public async Task CanSaveTableCreatedFromSingleRow()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue("Title");
        var table = ws.Range("A1").CreateTable();

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        await Assert.That(ms.Length).IsGreaterThan(0);
        await Assert.That(table.Field(0).Name).IsEqualTo("Title");
    }

    [Test]
    public async Task CreatingATableFromHeadersPushCellsBelow()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue("Title")
            .CellBelow().SetValue("X");
        ws.Range("A1").CreateTable();

        await Assert.That(ws.Cell("A2").Value).IsEqualTo(Blank.Value);
        await Assert.That(ws.Cell("A3").GetText()).IsEqualTo("X");
    }

    [Test]
    public async Task Inserting_Column_Sets_Header()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue("Categories")
            .CellBelow().SetValue("A")
            .CellBelow().SetValue("B")
            .CellBelow().SetValue("C");

        var table = ws.RangeUsed()!.CreateTable();
        table.InsertColumnsAfter(1);
        await Assert.That(table.HeadersRow()!.LastCell().GetText()).IsEqualTo("Column2");
    }

    [Test]
    public async Task DataRange_returns_null_if_empty()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue("Categories")
            .CellBelow().SetValue("A")
            .CellBelow().SetValue("B")
            .CellBelow().SetValue("C");

        var table = ws.RangeUsed()!.CreateTable();

        ws.Rows("2:4").Delete();

        await Assert.That(table.DataRange).IsNull();
    }

    [Test]
    public async Task SavingLoadingTableWithNewLineInHeader()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var columnName = "Line1" + Environment.NewLine + "Line2";
        ws.FirstCell().SetValue(columnName)
            .CellBelow().SetValue("A");
        ws.RangeUsed()!.CreateTable();
        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);
        var wb2 = new XLWorkbook(ms);
        var ws2 = wb2.Worksheet(1);
        var table2 = ws2.Table(0);
        var fieldName = table2.Field(0).Name;
        await Assert.That(fieldName).IsEqualTo("Line1\nLine2");
    }

    [Test]
    public async Task SavingLoadingTableWithNewLineInHeader2()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Test");

        var dt = new DataTable();
        var columnName = "Line1" + Environment.NewLine + "Line2";
        dt.Columns.Add(columnName);

        var dr = dt.NewRow();
        dr[columnName] = "some text";
        dt.Rows.Add(dr);
        ws.Cell(1, 1).InsertTable(dt);

        var table1 = ws.Table(0);
        var fieldName1 = table1.Field(0).Name;
        await Assert.That(fieldName1).IsEqualTo(columnName);

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);
        var wb2 = new XLWorkbook(ms);
        var ws2 = wb2.Worksheet(1);
        var table2 = ws2.Table(0);
        var fieldName2 = table2.Field(0).Name;
        await Assert.That(fieldName2).IsEqualTo("Line1\nLine2");
    }

    [Test]
    public async Task TableCreatedFromEmptyDataTable()
    {
        var dt = new DataTable("sheet1");
        dt.Columns.Add("col1", typeof(string));
        dt.Columns.Add("col2", typeof(double));

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().InsertTable(dt);
        await Assert.That(ws.Tables.First().ColumnCount()).IsEqualTo(2);
    }

    [Test]
    public async Task TableCreatedFromEmptyListOfInt()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().InsertTable(new List<int>());
        await Assert.That(ws.Tables.First().ColumnCount()).IsEqualTo(1);
    }

    [Test]
    public async Task TableCreatedFromEmptyListOfObject()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().InsertTable(new List<TestObjectWithoutAttributes>());
        await Assert.That(ws.Tables.First().ColumnCount()).IsEqualTo(2);
    }

    [Test]
    public async Task TableCreatedFromListOfObjectWithPropertyAttributes()
    {
        var l = new List<TestObjectWithAttributes>
        {
            new() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
            new() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 }
        };

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().InsertTable(l);
        await Assert.That(ws.Tables.First().ColumnCount()).IsEqualTo(4);
        await Assert.That(ws.FirstCell().Value).IsEqualTo("FirstColumn");
        await Assert.That(ws.FirstCell().CellRight().Value).IsEqualTo("SecondColumn");
        await Assert.That(ws.FirstCell().CellRight().CellRight().Value).IsEqualTo("SomeFieldNotProperty");
        await Assert.That(ws.FirstCell().CellRight().CellRight().CellRight().Value).IsEqualTo("UnOrderedColumn");
    }

    [Test]
    public async Task EmptyTableCreatedFromListOfObjectWithPropertyAttributes()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().InsertTable(new List<TestObjectWithAttributes>());
        await Assert.That(ws.Tables.First().ColumnCount()).IsEqualTo(4);
        await Assert.That(ws.FirstCell().Value).IsEqualTo("FirstColumn");
        await Assert.That(ws.FirstCell().CellRight().Value).IsEqualTo("SecondColumn");
        await Assert.That(ws.FirstCell().CellRight().CellRight().Value).IsEqualTo("SomeFieldNotProperty");
        await Assert.That(ws.FirstCell().CellRight().CellRight().CellRight().Value).IsEqualTo("UnOrderedColumn");
    }

    [Test]
    public async Task TableInsertAboveFromData()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue("Value");

        var table = ws.Range("A1:A2").CreateTable();
        table.SetShowTotalsRow()
            .Field(0).TotalsRowFunction = XLTotalsRowFunction.Sum;

        var row = table.DataRange!.FirstRow();
        row!.Field("Value").Value = 3;
        row = table.DataRange.InsertRowsAbove(1).First();
        row.Field("Value").Value = 2;
        row = table.DataRange.InsertRowsAbove(1).First();
        row.Field("Value").Value = 1;

        await Assert.That(ws.Cell(2, 1).GetDouble()).IsEqualTo(1);
        await Assert.That(ws.Cell(3, 1).GetDouble()).IsEqualTo(2);
        await Assert.That(ws.Cell(4, 1).GetDouble()).IsEqualTo(3);
    }

    [Test]
    public async Task TableInsertAboveFromRows()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue("Value");

        var table = ws.Range("A1:A2").CreateTable();
        table.SetShowTotalsRow()
            .Field(0).TotalsRowFunction = XLTotalsRowFunction.Sum;

        var row = table.DataRange!.FirstRow();
        row!.Field("Value").Value = 3;
        row = row.InsertRowsAbove(1).First();
        row.Field("Value").Value = 2;
        row = row.InsertRowsAbove(1).First();
        row.Field("Value").Value = 1;

        await Assert.That(ws.Cell(2, 1).GetDouble()).IsEqualTo(1);
        await Assert.That(ws.Cell(3, 1).GetDouble()).IsEqualTo(2);
        await Assert.That(ws.Cell(4, 1).GetDouble()).IsEqualTo(3);
    }

    [Test]
    public async Task TableInsertBelowFromData()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue("Value");

        var table = ws.Range("A1:A2").CreateTable();
        table.SetShowTotalsRow()
            .Field(0).TotalsRowFunction = XLTotalsRowFunction.Sum;

        var row = table.DataRange!.FirstRow();
        row!.Field("Value").Value = 1;
        row = table.DataRange.InsertRowsBelow(1).First();
        row.Field("Value").Value = 2;
        row = table.DataRange.InsertRowsBelow(1).First();
        row.Field("Value").Value = 3;

        await Assert.That(ws.Cell(2, 1).GetDouble()).IsEqualTo(1);
        await Assert.That(ws.Cell(3, 1).GetDouble()).IsEqualTo(2);
        await Assert.That(ws.Cell(4, 1).GetDouble()).IsEqualTo(3);
    }

    [Test]
    public async Task TableInsertBelowFromRows()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue("Value");

        var table = ws.Range("A1:A2").CreateTable();
        table.SetShowTotalsRow()
            .Field(0).TotalsRowFunction = XLTotalsRowFunction.Sum;

        var row = table.DataRange!.FirstRow();
        row!.Field("Value").Value = 1;
        row = row.InsertRowsBelow(1).First();
        row.Field("Value").Value = 2;
        row = row.InsertRowsBelow(1).First();
        row.Field("Value").Value = 3;

        await Assert.That(ws.Cell(2, 1).GetDouble()).IsEqualTo(1);
        await Assert.That(ws.Cell(3, 1).GetDouble()).IsEqualTo(2);
        await Assert.That(ws.Cell(4, 1).GetDouble()).IsEqualTo(3);
    }

    [Test]
    public async Task TableShowHeader()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue("Categories")
            .CellBelow().SetValue("A")
            .CellBelow().SetValue("B")
            .CellBelow().SetValue("C");

        var table = ws.RangeUsed()!.CreateTable();

        await Assert.That(table.Fields.First().Name).IsEqualTo("Categories");

        table.SetShowHeaderRow(false);

        await Assert.That(table.Fields.First().Name).IsEqualTo("Categories");

        await Assert.That(ws.Cell(1, 1).IsEmpty(XLCellsUsedOptions.All)).IsTrue();
        await Assert.That(table.HeadersRow()).IsEqualTo(null);
        await Assert.That(table.DataRange!.FirstRow()!.Field("Categories").GetText()).IsEqualTo("A");
        await Assert.That(table.DataRange.LastRow()!.Field("Categories").GetText()).IsEqualTo("C");
        await Assert.That(table.DataRange.FirstCell().GetText()).IsEqualTo("A");
        await Assert.That(table.DataRange.LastCell().GetText()).IsEqualTo("C");

        table.SetShowHeaderRow();
        var headerRow = table.HeadersRow();
        await Assert.That(headerRow).IsNotEqualTo(null);
        await Assert.That(headerRow!.Cell(1).GetText()).IsEqualTo("Categories");

        table.SetShowHeaderRow(false);

        ws.FirstCell().SetValue("x");

        table.SetShowHeaderRow();

        await Assert.That(ws.FirstCell().GetText()).IsEqualTo("x");
        await Assert.That(ws.Cell("A2").GetText()).IsEqualTo("Categories");
        await Assert.That(headerRow).IsNotEqualTo(null);
        await Assert.That(table.DataRange.FirstRow()!.Field("Categories").GetText()).IsEqualTo("A");
        await Assert.That(table.DataRange.LastRow()!.Field("Categories").GetText()).IsEqualTo("C");
        await Assert.That(table.DataRange.FirstCell().GetText()).IsEqualTo("A");
        await Assert.That(table.DataRange.LastCell().GetText()).IsEqualTo("C");
    }

    [Test]
    [Arguments("Amount")]
    [Arguments("AMOUNT")]
    [Arguments("amount")]
    public async Task FieldNames_of_XLTable_are_case_insensitive(string fieldName)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var table = ws.Cell("A1").InsertTable([new { Amount = 1 }]);

        var expectedField = table.Field(0);
        await Assert.That(table.Field(fieldName)).IsSameReferenceAs(expectedField);
    }

    [Test]
    public async Task ChangeFieldName()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet");
        ws.Cell("A1").SetValue("FName")
            .CellBelow().SetValue("John");

        ws.Cell("B1").SetValue("LName")
            .CellBelow().SetValue("Doe");

        var tbl = ws.RangeUsed()!.CreateTable();
        var nameBefore = tbl.Field(tbl.Fields.Last().Index).Name;
        tbl.Field(tbl.Fields.Last().Index).Name = "LastName";
        var nameAfter = tbl.Field(tbl.Fields.Last().Index).Name;

        var cellValue = ws.Cell("B1").GetText();

        await Assert.That(nameBefore).IsEqualTo("LName");
        await Assert.That(nameAfter).IsEqualTo("LastName");
        await Assert.That(cellValue).IsEqualTo("LastName");

        tbl.ShowHeaderRow = false;
        tbl.Field(tbl.Fields.Last().Index).Name = "LastNameChanged";
        nameAfter = tbl.Field(tbl.Fields.Last().Index).Name;
        await Assert.That(nameAfter).IsEqualTo("LastNameChanged");

        tbl.SetShowHeaderRow(true);
        nameAfter = (string)tbl.Cell("B1").Value;
        await Assert.That(nameAfter).IsEqualTo("LastNameChanged");

        var field = tbl.Field("LastNameChanged");
        await Assert.That(field.Name).IsEqualTo("LastNameChanged");

        tbl.Cell(1, 1).Value = "FirstName";
        await Assert.That(tbl.Field(0).Name).IsEqualTo("FirstName");
    }

    [Test]
    public async Task CanDeleteTableColumn()
    {
        var l = new List<TestObjectWithAttributes>
        {
            new() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
            new() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 }
        };

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var table = ws.FirstCell().InsertTable(l);

        table.Column("C").Delete();

        await Assert.That(table.Fields.Count()).IsEqualTo(3);

        await Assert.That(table.Fields.First().Name).IsEqualTo("FirstColumn");
        await Assert.That(table.Fields.First().Index).IsEqualTo(0);

        await Assert.That(table.Fields.Last().Name).IsEqualTo("UnOrderedColumn");
        await Assert.That(table.Fields.Last().Index).IsEqualTo(2);
    }

    [Test]
    public async Task TestFieldCellTypes()
    {
        var l = new List<TestObjectWithAttributes>
        {
            new() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
            new() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 }
        };

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var table = ws.Cell("B2").InsertTable(l);

        await Assert.That(table.Fields.Count()).IsEqualTo(4);

        await Assert.That(table.Field(0).HeaderCell!.Address.ToString()).IsEqualTo("B2");
        await Assert.That(table.Field(1).HeaderCell!.Address.ToString()).IsEqualTo("C2");
        await Assert.That(table.Field(2).HeaderCell!.Address.ToString()).IsEqualTo("D2");
        await Assert.That(table.Field(3).HeaderCell!.Address.ToString()).IsEqualTo("E2");

        await Assert.That(table.Field(0).TotalsCell).IsNull();
        await Assert.That(table.Field(1).TotalsCell).IsNull();
        await Assert.That(table.Field(2).TotalsCell).IsNull();
        await Assert.That(table.Field(3).TotalsCell).IsNull();

        table.SetShowTotalsRow();

        await Assert.That(table.Field(0).TotalsCell!.Address.ToString()).IsEqualTo("B5");
        await Assert.That(table.Field(1).TotalsCell!.Address.ToString()).IsEqualTo("C5");
        await Assert.That(table.Field(2).TotalsCell!.Address.ToString()).IsEqualTo("D5");
        await Assert.That(table.Field(3).TotalsCell!.Address.ToString()).IsEqualTo("E5");

        var field = table.Fields.Last();

        await Assert.That(field.Column.RangeAddress.ToString()).IsEqualTo("E2:E5");
        await Assert.That(field.DataCells.First().Address.ToString()).IsEqualTo("E3");
        await Assert.That(field.DataCells.Last().Address.ToString()).IsEqualTo("E4");
    }

    [Test]
    public async Task CanDeleteTable()
    {
        var l = new List<TestObjectWithAttributes>
        {
            new() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
            new() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 }
        };

        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(l);
            wb.SaveAs(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();
            var table = ws.Tables.First();

            ws.Tables.Remove(table.Name);
            await Assert.That(ws.Tables.Count()).IsEqualTo(0);
            wb.Save();
        }
    }

    [Test]
    public async Task TableNameCannotBeValidCellName()
    {
        var dt = new DataTable("sheet1");
        dt.Columns.Add("Patient", typeof(string));
        dt.Rows.Add("David");

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        await Assert.That(() => ws.Cell(1, 1).InsertTable(dt, "May2019")).Throws<ArgumentException>();
        await Assert.That(() => ws.Cell(1, 1).InsertTable(dt, "A1")).Throws<ArgumentException>();
        await Assert.That(() => ws.Cell(1, 1).InsertTable(dt, "R1C2")).Throws<ArgumentException>();
        await Assert.That(() => ws.Cell(1, 1).InsertTable(dt, "r3c2")).Throws<ArgumentException>();
        await Assert.That(() => ws.Cell(1, 1).InsertTable(dt, "R2C33333")).Throws<ArgumentException>();
        await Assert.That(() => ws.Cell(1, 1).InsertTable(dt, "RC")).Throws<ArgumentException>();
    }

    [Test]
    public async Task TableNameSetWhenAddingWorksheetWithDataTable()
    {
        var dt = new DataTable("sheet1");
        dt.Columns.Add("Patient", typeof(string));
        dt.Rows.Add("David");

        using (var wb = new XLWorkbook())
        {
            // Generated table name is used and should not be an issue
            await Assert.That(() => wb.AddWorksheet(dt, "t1")).ThrowsNothing();
        }

        using (var wb = new XLWorkbook())
        {
            // Should pass because t1 is a valid sheet name, and is not used for the tableName
            await Assert.That(() => wb.AddWorksheet(dt, "t1", "table1")).ThrowsNothing();

            await Assert.That(wb.Worksheets.Count).IsEqualTo(1);
            await Assert.That(wb.Worksheet(1).Tables.Count()).IsEqualTo(1);
        }
    }

    [Test]
    public async Task CanDeleteTableField()
    {
        var l = new List<TestObjectWithAttributes>
        {
            new() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
            new() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 }
        };

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var table = ws.Cell("B2").InsertTable(l);

        await Assert.That(table.RangeAddress.ToString()).IsEqualTo("B2:E4");

        table.Field("SomeFieldNotProperty").Delete();

        await Assert.That(table.Fields.Count()).IsEqualTo(3);

        await Assert.That(table.Fields.First().Name).IsEqualTo("FirstColumn");
        await Assert.That(table.Fields.First().Index).IsEqualTo(0);

        await Assert.That(table.Fields.Last().Name).IsEqualTo("UnOrderedColumn");
        await Assert.That(table.Fields.Last().Index).IsEqualTo(2);

        await Assert.That(table.RangeAddress.ToString()).IsEqualTo("B2:D4");
    }

    [Test]
    public async Task CanDeleteTableRows()
    {
        var l = new List<TestObjectWithAttributes>
        {
            new() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
            new() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 },
            new() { Column1 = "e", Column2 = "f", MyField = 6, UnOrderedColumn = 555 },
            new() { Column1 = "g", Column2 = "h", MyField = 7, UnOrderedColumn = 333 }
        };

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var table = ws.Cell("B2").InsertTable(l);

        await Assert.That(table.RangeAddress.ToString()).IsEqualTo("B2:E6");

        table.DataRange!.Rows(3, 4).Delete();

        await Assert.That(table.DataRange.Rows().Count()).IsEqualTo(2);

        await Assert.That(table.DataRange.FirstCell().Value).IsEqualTo("b");
        await Assert.That(table.DataRange.LastCell().Value).IsEqualTo(777);

        await Assert.That(table.RangeAddress.ToString()).IsEqualTo("B2:E4");
    }

    [Test]
    public async Task OverlappingTablesThrowsException()
    {
        var dt = new DataTable("sheet1");
        dt.Columns.Add("col1", typeof(string));
        dt.Columns.Add("col2", typeof(double));

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().InsertTable(dt, true);
        await Assert.That(() => ws.FirstCell().CellRight().InsertTable(dt, true)).Throws<InvalidOperationException>();
    }

    [Test]
    public async Task OverwritingTableHeaders()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var table = ws.Cell("A1").InsertTable(new object[]
        {
            ("Header 1", "Header 2"),
            (1, 2)
        }, true);

        // Overwrite the headers of the table with non-string values
        ws.Cell("A1").InsertData(new object[]
        {
            (XLError.IncompatibleValue, 7)
        });

        // The non-string data inserted to headers were converted to strings and used as a field names.
        await Assert.That(table.Field(0).Name).IsEqualTo("#VALUE!");
        await Assert.That(ws.Cell("A1").Value).IsEqualTo("#VALUE!");
        await Assert.That(table.Field(1).Name).IsEqualTo("7");
        await Assert.That(ws.Cell("B1").Value).IsEqualTo("7");
    }

    [Test]
    public async Task OverwritingTableTotalsRow()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var data1 = Enumerable.Range(1, 10)
            .Select(i =>
                new
                {
                    Index = i,
                    Character = Convert.ToChar(64 + i),
                    String = new string('a', i)
                });

        var table = ws.FirstCell().InsertTable(data1, true)
            .SetShowHeaderRow()
            .SetShowTotalsRow();
        table.Fields.First().TotalsRowFunction = XLTotalsRowFunction.Sum;

        var data2 = Enumerable.Range(1, 20)
            .Select(i =>
                new
                {
                    Index = i,
                    Character = Convert.ToChar(64 + i),
                    String = new string('b', i),
                    Int = 64 + i
                });

        ws.FirstCell().CellBelow().InsertData(data2);

        // Was Fields.ForEach(f => Assert...); ForEach takes an Action, so an awaited
        // assertion needs an explicit loop.
        foreach (var f in table.Fields)
        {
            await Assert.That(f.TotalsRowFunction).IsEqualTo(XLTotalsRowFunction.None);
        }

        await Assert.That(table.Field(0).TotalsRowLabel).IsEqualTo("11");
        await Assert.That(table.Field(1).TotalsRowLabel).IsEqualTo("K");
        await Assert.That(table.Field(2).TotalsRowLabel).IsEqualTo("bbbbbbbbbbb");
    }

    [Test]
    public async Task TableRenameTests()
    {
        var l = new List<TestObjectWithAttributes>
        {
            new() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
            new() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 }
        };

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var table1 = ws.FirstCell().InsertTable(l);
        var table2 = ws.Cell("A10").InsertTable(l);

        await Assert.That(table1.Name).IsEqualTo("Table1");
        await Assert.That(table2.Name).IsEqualTo("Table2");

        table1.Name = "table1";
        await Assert.That(table1.Name).IsEqualTo("table1");

        table1.Name = "_table1";
        await Assert.That(table1.Name).IsEqualTo("_table1");

        table1.Name = "\\table1";
        await Assert.That(table1.Name).IsEqualTo("\\table1");

        await Assert.That(() => table1.Name = "").Throws<ArgumentException>();
        await Assert.That(() => table1.Name = "R").Throws<ArgumentException>();
        await Assert.That(() => table1.Name = "C").Throws<ArgumentException>();
        await Assert.That(() => table1.Name = "r").Throws<ArgumentException>();
        await Assert.That(() => table1.Name = "c").Throws<ArgumentException>();

        await Assert.That(() => table1.Name = "123").Throws<ArgumentException>();
        await Assert.That(() => table1.Name = new string('A', 256)).Throws<ArgumentException>();

        await Assert.That(() => table1.Name = "Table2").Throws<ArgumentException>();
        await Assert.That(() => table1.Name = "TABLE2").Throws<ArgumentException>();
    }

    [Test]
    public async Task CanResizeTable()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var data1 = Enumerable.Range(1, 10)
            .Select(i =>
                new
                {
                    Index = i,
                    Character = Convert.ToChar(64 + i),
                    String = new string('a', i)
                });

        var table = ws.FirstCell().InsertTable(data1, true)
            .SetShowHeaderRow()
            .SetShowTotalsRow();
        table.Fields.First().TotalsRowFunction = XLTotalsRowFunction.Sum;

        var data2 = Enumerable.Range(1, 10)
            .Select(i =>
                new
                {
                    Index = i,
                    Character = Convert.ToChar(64 + i),
                    String = new string('b', i),
                    Integer = 64 + i
                });

        ws.FirstCell().CellBelow().InsertData(data2);
        table.Resize(table.FirstCell().Address, table.AsRange().LastCell().CellRight().Address);

        await Assert.That(table.Fields.Count()).IsEqualTo(4);

        await Assert.That(table.Field(3).Name).IsEqualTo("Column4");

        ws.Cell("D1").Value = "Integer";
        await Assert.That(table.Field(3).Name).IsEqualTo("Integer");
    }

    [Test]
    public async Task TableAsDynamicEnumerable()
    {
        var l = new List<TestObjectWithAttributes>
        {
            new() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
            new() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 }
        };

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var table = ws.FirstCell().InsertTable(l);

        foreach (var d in table.AsDynamicEnumerable())
        {
            await Assert.That(() =>
            {
                _ = d.FirstColumn;
                _ = d.SecondColumn;
                _ = d.UnOrderedColumn;
                _ = d.SomeFieldNotProperty;
            }).ThrowsNothing();
        }
    }

    [Test]
    public async Task TableAsDotNetDataTable()
    {
        var l = new List<TestObjectWithAttributes>
        {
            new() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
            new() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 }
        };

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var table = ws.FirstCell().InsertTable(l).AsNativeDataTable();

        await Assert.That(table.Columns.Count).IsEqualTo(4);
        await Assert.That(table.Columns[0].ColumnName).IsEqualTo("FirstColumn");
        await Assert.That(table.Columns[1].ColumnName).IsEqualTo("SecondColumn");
        await Assert.That(table.Columns[2].ColumnName).IsEqualTo("SomeFieldNotProperty");
        await Assert.That(table.Columns[3].ColumnName).IsEqualTo("UnOrderedColumn");

        await Assert.That(table.Columns[0].DataType).IsEqualTo(typeof(string));
        await Assert.That(table.Columns[1].DataType).IsEqualTo(typeof(string));
        await Assert.That(table.Columns[2].DataType).IsEqualTo(typeof(double));
        await Assert.That(table.Columns[3].DataType).IsEqualTo(typeof(double));

        var dr = table.Rows[0];
        await Assert.That(dr["FirstColumn"]).IsEqualTo("b");
        await Assert.That(dr["SecondColumn"]).IsEqualTo("a");
        await Assert.That(Convert.ToDouble(dr["SomeFieldNotProperty"], CultureInfo.InvariantCulture)).IsEqualTo(4d);
        await Assert.That(Convert.ToDouble(dr["UnOrderedColumn"], CultureInfo.InvariantCulture)).IsEqualTo(999d);

        dr = table.Rows[1];
        await Assert.That(dr["FirstColumn"]).IsEqualTo("d");
        await Assert.That(dr["SecondColumn"]).IsEqualTo("c");
        await Assert.That(Convert.ToDouble(dr["SomeFieldNotProperty"], CultureInfo.InvariantCulture)).IsEqualTo(5d);
        await Assert.That(Convert.ToDouble(dr["UnOrderedColumn"], CultureInfo.InvariantCulture)).IsEqualTo(777d);
    }

    [Test]
    public async Task TestTableCellTypes()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var data1 = Enumerable.Range(1, 10)
            .Select(i =>
                new
                {
                    Index = i,
                    Character = Convert.ToChar(64 + i),
                    String = new string('a', i)
                });

        var table = ws.FirstCell().InsertTable(data1, true)
            .SetShowHeaderRow()
            .SetShowTotalsRow();
        table.Fields.First().TotalsRowFunction = XLTotalsRowFunction.Sum;

        await Assert.That(table.HeadersRow()!.Cell(1).TableCellType()).IsEqualTo(XLTableCellType.Header);
        await Assert.That(table.HeadersRow()!.Cell(1).CellBelow().TableCellType()).IsEqualTo(XLTableCellType.Data);
        await Assert.That(table.TotalsRow()!.Cell(1).TableCellType()).IsEqualTo(XLTableCellType.Total);
        await Assert.That(ws.Cell("Z100").TableCellType()).IsEqualTo(XLTableCellType.None);
    }

    [Test]
    public async Task TotalsFunctionsOfHeadersWithWeirdCharacters()
    {
        var l = new List<TestObjectWithAttributes>
        {
            new() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
            new() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 }
        };

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().InsertTable(l, false);

        // Give the headings weird names (i.e. spaces, hashes, single quotes
        ws.Cell("A1").Value = "ABCD    ";
        ws.Cell("B1").Value = "   #BCD";
        ws.Cell("C1").Value = "   as'df   ";
        ws.Cell("D1").Value = "Normal";

        var table = ws.RangeUsed()!.CreateTable();
        await Assert.That(table).IsNotNull();

        table.ShowTotalsRow = true;
        table.Field(0).TotalsRowFunction = XLTotalsRowFunction.Count;
        table.Field(1).TotalsRowFunction = XLTotalsRowFunction.Count;
        table.Field(2).TotalsRowFunction = XLTotalsRowFunction.Sum;
        table.Field(3).TotalsRowFunction = XLTotalsRowFunction.Sum;

        await Assert.That(table.Field(0).TotalsRowFormulaA1).IsEqualTo("SUBTOTAL(103,Table1[[ABCD    ]])");
        await Assert.That(table.Field(1).TotalsRowFormulaA1).IsEqualTo("SUBTOTAL(103,Table1[[   '#BCD]])");
        await Assert.That(table.Field(2).TotalsRowFormulaA1).IsEqualTo("SUBTOTAL(109,Table1[[   as''df   ]])");
        await Assert.That(table.Field(3).TotalsRowFormulaA1).IsEqualTo("SUBTOTAL(109,[Normal])");
    }

    [Test]
    public async Task TotalsFunctionsOfHeadersWithInteriorSpaces()
    {
        // Regression for issue #2864: column names with an interior space (e.g. dates
        // like "Feb 2023") must be wrapped in an extra pair of brackets in the totals
        // row structured reference. The single-bracket form Table1[Feb 2023] causes
        // Excel to raise a #NAME error and repair the file on open.
        //
        // Headers that combine an interior space with a quoted table-field character
        // (' or #) must be both escaped (' -> '', # -> '#) and double-bracket wrapped,
        // e.g. "Feb '23" -> Table1[[Feb ''23]].
        var l = new List<TestObjectWithAttributes>
        {
            new() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
            new() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 }
        };

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().InsertTable(l, false);

        ws.Cell("A1").Value = "Jan 2023";
        ws.Cell("B1").Value = "Feb 2023";
        ws.Cell("C1").Value = "Mar 2023";
        ws.Cell("D1").Value = "Total";

        // Extra columns whose headers mix an interior space with quoted characters.
        ws.Cell("E1").Value = "Feb '23";
        ws.Cell("E2").Value = 1;
        ws.Cell("E3").Value = 2;
        ws.Cell("F1").Value = "Q1 #1";
        ws.Cell("F2").Value = 3;
        ws.Cell("F3").Value = 4;

        var table = ws.RangeUsed()!.CreateTable();
        table.ShowTotalsRow = true;
        table.Field(0).TotalsRowFunction = XLTotalsRowFunction.Sum;
        table.Field(1).TotalsRowFunction = XLTotalsRowFunction.Average;
        table.Field(2).TotalsRowFunction = XLTotalsRowFunction.Count;
        table.Field(3).TotalsRowFunction = XLTotalsRowFunction.Sum;
        table.Field(4).TotalsRowFunction = XLTotalsRowFunction.Sum;
        table.Field(5).TotalsRowFunction = XLTotalsRowFunction.Count;

        await Assert.That(table.Field(0).TotalsRowFormulaA1).IsEqualTo("SUBTOTAL(109,Table1[[Jan 2023]])");
        await Assert.That(table.Field(1).TotalsRowFormulaA1).IsEqualTo("SUBTOTAL(101,Table1[[Feb 2023]])");
        await Assert.That(table.Field(2).TotalsRowFormulaA1).IsEqualTo("SUBTOTAL(103,Table1[[Mar 2023]])");
        // No space => single-bracket, table name not prepended.
        await Assert.That(table.Field(3).TotalsRowFormulaA1).IsEqualTo("SUBTOTAL(109,[Total])");
        // Interior space + quoted character => escaped and double-bracket wrapped.
        await Assert.That(table.Field(4).TotalsRowFormulaA1).IsEqualTo("SUBTOTAL(109,Table1[[Feb ''23]])");
        await Assert.That(table.Field(5).TotalsRowFormulaA1).IsEqualTo("SUBTOTAL(103,Table1[[Q1 '#1]])");
    }

    [Test]
    public async Task CannotCreateDuplicateTablesOverSameRange()
    {
        var l = new List<TestObjectWithAttributes>
        {
            new() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
            new() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 }
        };

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().InsertTable(l);
        await Assert.That(() => ws.RangeUsed()!.CreateTable()).Throws<InvalidOperationException>();
    }

    [Test]
    public async Task CannotCreateTableOverExistingAutoFilter()
    {
        using var wb = new XLWorkbook();

        var data = Enumerable.Range(1, 10).Select(i => new
        {
            Index = i,
            String = $"String {i}"
        });

        var ws = wb.AddWorksheet();
        ws.FirstCell().InsertTable(data, createTable: false);
        ws.RangeUsed()!.SetAutoFilter().Column(1).AddFilter(5);

        await Assert.That(() => ws.RangeUsed()!.CreateTable()).Throws<InvalidOperationException>();
    }

    [Test]
    public async Task CopyTableSameWorksheet()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet1");

        var table = ws1.Range("A1:C2").AsTable();

        Action action = () => table.CopyTo(ws1);

        await Assert.That(action).Throws<InvalidOperationException>();
    }

    [Test]
    public async Task CanInsertDateTimeOffset()
    {
        var now = DateTimeOffset.Now;

        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet();
        ws1.FirstCell().InsertTable([new { TimeStamp = now }]);

        // C# Supports 7 digits milliseconds, but excel only 3
        const string format = "yyyy-MM-dd HH:mm:ss.fff";

        var actual = ws1.Cell("A2").GetDateTime().ToString(format);
        var expected = now.DateTime.ToString(format);
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    public async Task CopyDetachedTableDifferentWorksheets()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet1");
        ws1.Cell("A1").Value = "Custom column 1";
        ws1.Cell("B1").Value = "Custom column 2";
        ws1.Cell("C1").Value = "Custom column 3";
        ws1.Cell("A2").Value = "Value 1";
        ws1.Cell("B2").Value = 123.45;
        ws1.Cell("C2").Value = new DateTime(2018, 5, 10, 0, 0, 0, DateTimeKind.Unspecified);
        var original = ws1.Range("A1:C2").AsTable("Detached_table");
        var ws2 = wb.Worksheets.Add("Sheet2");

        var copy = original.CopyTo(ws2);

        await Assert.That(ws1.Tables.Count()).IsEqualTo(0); // We did not add it
        await Assert.That(ws2.Tables.Count()).IsEqualTo(1);

        await AssertTablesAreEqual(original, copy);

        await Assert.That(copy.RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("Sheet2!A1:C2");
        await Assert.That(ws2.Cell("A1").Value).IsEqualTo("Custom column 1");
        await Assert.That(ws2.Cell("B1").Value).IsEqualTo("Custom column 2");
        await Assert.That(ws2.Cell("C1").Value).IsEqualTo("Custom column 3");
        await Assert.That(ws2.Cell("A2").Value).IsEqualTo("Value 1");
        await Assert.That((double)ws2.Cell("B2").Value).IsEqualTo(123.45).Within(XLHelper.Epsilon);
        await Assert.That(ws2.Cell("C2").Value).IsEqualTo(new DateTime(2018, 5, 10, 0, 0, 0, DateTimeKind.Unspecified));
    }

    [Test]
    public async Task CopyTableDifferentWorksheets()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet1");
        ws1.Cell("A1").Value = "Custom column 1";
        ws1.Cell("B1").Value = "Custom column 2";
        ws1.Cell("C1").Value = "Custom column 3";
        ws1.Cell("A2").Value = "Value 1";
        ws1.Cell("B2").Value = 123.45;
        ws1.Cell("C2").Value = new DateTime(2018, 5, 10, 0, 0, 0, DateTimeKind.Unspecified);
        var original = ws1.Range("A1:C2").AsTable("Attached_table");
        ws1.Tables.Add(original);
        var ws2 = wb.Worksheets.Add("Sheet2");

        original.CopyTo(ws2);

        await Assert.That(ws1.Tables.Count()).IsEqualTo(1);
        await Assert.That(ws2.Tables.Count()).IsEqualTo(1);

        var copy = ws2.Tables.First();

        await AssertTablesAreEqual(original, copy);

        await Assert.That(copy.RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("Sheet2!A1:C2");
        await Assert.That(ws2.Cell("A1").Value).IsEqualTo("Custom column 1");
        await Assert.That(ws2.Cell("B1").Value).IsEqualTo("Custom column 2");
        await Assert.That(ws2.Cell("C1").Value).IsEqualTo("Custom column 3");
        await Assert.That(ws2.Cell("A2").Value).IsEqualTo("Value 1");
        await Assert.That((double)ws2.Cell("B2").Value).IsEqualTo(123.45).Within(XLHelper.Epsilon);
        await Assert.That(ws2.Cell("C2").Value).IsEqualTo(new DateTime(2018, 5, 10, 0, 0, 0, DateTimeKind.Unspecified));
    }

    [Test]
    public async Task NewTableHasNullRelId()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell("A1").Value = "Custom column 1";
            ws.Cell("B1").Value = "Custom column 2";
            ws.Cell("C1").Value = "Custom column 3";
            ws.Cell("A2").Value = "Value 1";
            ws.Cell("B2").Value = 123.45;
            ws.Cell("C2").Value = new DateTime(2018, 5, 10, 0, 0, 0, DateTimeKind.Unspecified);
            var original = ws.Range("A1:C2").CreateTable("Attached_table");

            await Assert.That(ws.Tables.Count()).IsEqualTo(1);
            await Assert.That((original as XLTable)!.RelId).IsNull();

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.Add("Sheet2");
            var original = wb.Worksheets.First().Tables.First();

            await Assert.That((original as XLTable)!.RelId).IsNotNull();

            var copy = original.CopyTo(ws);

            await Assert.That(ws.Tables.Count()).IsEqualTo(1);
            await Assert.That((copy as XLTable)!.RelId).IsNull();

            await AssertTablesAreEqual(original, copy);

            await Assert.That(copy.RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("Sheet2!A1:C2");
            await Assert.That(ws.Cell("A1").Value).IsEqualTo("Custom column 1");
            await Assert.That(ws.Cell("B1").Value).IsEqualTo("Custom column 2");
            await Assert.That(ws.Cell("C1").Value).IsEqualTo("Custom column 3");
            await Assert.That(ws.Cell("A2").Value).IsEqualTo("Value 1");
            await Assert.That((double)ws.Cell("B2").Value).IsEqualTo(123.45).Within(XLHelper.Epsilon);
            await Assert.That(ws.Cell("C2").Value).IsEqualTo(new DateTime(2018, 5, 10, 0, 0, 0, DateTimeKind.Unspecified));
        }
    }

    [Test]
    public async Task CopyTableWithoutData()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet1");
        ws1.Cell("A1").Value = "Custom column 1";
        ws1.Cell("B1").Value = "Custom column 2";
        ws1.Cell("C1").Value = "Custom column 3";
        ws1.Cell("A2").Value = "Value 1";
        ws1.Cell("B2").Value = 123.45;
        ws1.Cell("C2").Value = new DateTime(2018, 5, 10, 0, 0, 0, DateTimeKind.Unspecified);
        var original = ws1.Range("A1:C2").AsTable("Attached_table");
        ws1.Tables.Add(original);
        var ws2 = wb.Worksheets.Add("Sheet2") as XLWorksheet;

        var copy = (original as XLTable)!.CopyTo(ws2!, false);

        await AssertTablesAreEqual(original, copy);

        await Assert.That(copy.RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("Sheet2!A1:C2");
        await Assert.That(ws2!.Cell("A1")!.Value).IsEqualTo("Custom column 1");
        await Assert.That(ws2.Cell("B1")!.Value).IsEqualTo("Custom column 2");
        await Assert.That(ws2.Cell("C1")!.Value).IsEqualTo("Custom column 3");
        await Assert.That(ws2.Cell("A2")!.Value).IsEqualTo(Blank.Value);
        await Assert.That(ws2.Cell("B2")!.Value).IsEqualTo(Blank.Value);
        await Assert.That(ws2.Cell("C2")!.Value).IsEqualTo(Blank.Value);
    }

    [Test]
    public async Task SavingTableWithNullDataRangeThrowsException()
    {
        using var ms = new MemoryStream();
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var data = Enumerable.Range(1, 10)
            .Select(i => new
            {
                Number = i,
                NumberString = string.Concat("Number", i.ToString())
            });

        var table = ws.FirstCell()
            .InsertTable(data)
            .SetShowTotalsRow();

        table.Fields.Last().TotalsRowFunction = XLTotalsRowFunction.Count;

        table.DataRange!.Rows()
            .OrderByDescending(r => r.RowNumber())
            .ToList()
            .ForEach(r => r.WorksheetRow().Delete());

        await Assert.That(table.DataRange).IsNull();
        await Assert.That(() => wb.SaveAs(ms)).Throws<EmptyTableException>();
    }

    [Test]
    public async Task Save_totals_row_label_cell_with_sst_id_matching_the_label()
    {
        // Issue #2602 test. The totals row  wasn't saved with compact SST ID from file, but with a memory SST that has holes.
        await TestHelper.CreateAndCompare(wb =>
        {
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "Dummy1"; // First inserted text - index=0, reference count = 1
            ws.Cell("A2").Value = "Dummy2"; // Second inserted text - index=1, reference count = 1
            ws.Cell("A3").Value = "Dummy3"; // Third inserted text - index=2, reference count = 1
            ws.Cell("A4").Value = "Text"; // Fourth inserted text - index=3, reference count = 1
            var table = ws.Cell("A5").InsertTable([("Text", 17)]); // Also inserts header Item1 and Item2.
            table.ShowTotalsRow = true;
            table.Field(0).TotalsRowLabel = "Text"; // reference count = 3

            // Remove "Dummy*" text. That way, the "Text", "Item1" and "Item2" will be in index 0..2 that were occupied by Dummy*
            // Ensure that cell in total row label A7 references "Text" SST ID
            ws.Range("A1:A3").Value = Blank.Value;
        }, @"Other\Tables\TotalRowSstId.xlsx");
    }

    [Test]
    public async Task CanCreateTableWithWhiteSpaceColumnHeaders()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.Cell("A1").SetValue("Header");
        ws.Cell("B1").SetValue(new string(' ', 1));
        ws.Cell("C1").SetValue(new string(' ', 2));
        ws.Cell("D1").SetValue(new string(' ', 3));

        var table = ws.Range("A1:E3").CreateTable("Table1");

        await Assert.That(table.Field(0).Name).IsEqualTo("Header");
        await Assert.That(table.Field(1).Name).IsEqualTo(new string(' ', 1));
        await Assert.That(table.Field(2).Name).IsEqualTo(new string(' ', 2));
        await Assert.That(table.Field(3).Name).IsEqualTo(new string(' ', 3));
        await Assert.That(table.Field(4).Name).IsEqualTo("Column5");
    }

    [Test]
    public async Task TableNotFound()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        await Assert.That(() => ws.Table("dummy")).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => wb.Table("dummy")).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task SecondTableOnNewSheetHasUniqueName()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet();
        var t1 = ws1.FirstCell().InsertTable(Enumerable.Range(1, 10).Select(i => new { Number = i }));
        await Assert.That(t1.Name).IsEqualTo("Table1");

        var ws2 = wb.AddWorksheet();
        var t2 = ws2.FirstCell().InsertTable(Enumerable.Range(1, 10).Select(i => new { Number = i }));
        await Assert.That(t2.Name).IsEqualTo("Table2");
    }

    private static async Task AssertTablesAreEqual(IXLTable table1, IXLTable table2)
    {
        await Assert.That(table2.RangeAddress.ToString(XLReferenceStyle.A1, false)).IsEqualTo(table1.RangeAddress.ToString(XLReferenceStyle.A1, false));
        await Assert.That(table2.Fields.Count()).IsEqualTo(table1.Fields.Count());
        for (var j = 0; j < table1.Fields.Count(); j++)
        {
            var originalField = table1.Fields.ElementAt(j);
            var copyField = table2.Fields.ElementAt(j);
            await Assert.That(copyField.Name).IsEqualTo(originalField.Name);
            if (table1.ShowTotalsRow)
            {
                await Assert.That(copyField.TotalsRowFormulaA1).IsEqualTo(originalField.TotalsRowFormulaA1);
                await Assert.That(copyField.TotalsRowFunction).IsEqualTo(originalField.TotalsRowFunction);
            }
        }

        await Assert.That(table2.Name).IsEqualTo(table1.Name);
        await Assert.That(table2.ShowAutoFilter).IsEqualTo(table1.ShowAutoFilter);
        await Assert.That(table2.ShowColumnStripes).IsEqualTo(table1.ShowColumnStripes);
        await Assert.That(table2.ShowHeaderRow).IsEqualTo(table1.ShowHeaderRow);
        await Assert.That(table2.ShowRowStripes).IsEqualTo(table1.ShowRowStripes);
        await Assert.That(table2.ShowTotalsRow).IsEqualTo(table1.ShowTotalsRow);
        await Assert.That((table2.Style as XLStyle)!.Value).IsEqualTo((table1.Style as XLStyle)!.Value);
        await Assert.That(table2.Theme).IsEqualTo(table1.Theme);
    }

    [Test]
    [Arguments(typeof(string))]
    [Arguments(typeof(object))]
    public async Task InsertData_WhenValuesAreDbNull_WritesBlanks(Type columnType)
    {
        // arrange
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var data = new DataTable();
        data.Columns.Add("Name", columnType);
        data.Rows.Add("Mario");
        data.Rows.Add(DBNull.Value); // This should be written as blank
        data.Rows.Add("Carlo");

        // act
        var cell = ws.FirstCell();
        var table = cell.InsertTable(data);

        // assert
        await Assert.That(table!.Cell(1, 1).Value).IsEqualTo("Name");
        await Assert.That(table.Cell(2, 1).Value).IsEqualTo("Mario");
        await Assert.That(table.Cell(3, 1).Value.IsBlank).IsTrue();
        await Assert.That(table.Cell(4, 1).Value).IsEqualTo("Carlo");
    }

    [Test]
    public async Task DataRowCount_returns_data_rows_excluding_header_and_totals()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var data = Enumerable.Range(1, 5).Select(i => new { Name = "Item" + i, Value = i });
        var table = ws.FirstCell().InsertTable(data, "Table1", true);

        // Table has header + 5 data rows
        await Assert.That(table.RowCount()).IsEqualTo(6);
        await Assert.That(table.DataRowCount).IsEqualTo(5);
        await Assert.That(table.DataRange!.RowCount()).IsEqualTo(5);

        // Delete all data rows from worksheet (rows 2..6)
        for (var row = 6; row >= 2; row--)
            ws.Row(row).Delete();

        // Resize table to header-only range
        var headerOnlyRange = ws.Range(1, 1, 1, 2);
        table.Resize(headerOnlyRange);

        // After resize, the table range spans only 1 row (the header)
        await Assert.That(table.RowCount()).IsEqualTo(1);
        // DataRowCount should be 0 and DataRange should be null
        await Assert.That(table.DataRowCount).IsEqualTo(0);
        await Assert.That(table.DataRange).IsNull();
    }

    [Test]
    public async Task DataRange_FirstRowUsed_returns_null_when_no_used_rows()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // Create a table with a header and one empty data row
        ws.Cell("A1").Value = "Col1";
        ws.Cell("B1").Value = "Col2";
        ws.Cell("A2").Value = Blank.Value;
        ws.Cell("B2").Value = Blank.Value;

        var table = ws.Range("A1:B2").CreateTable("TestTable");
        var dataRange = table.DataRange!;

        // FirstRowUsed should return null, not throw NRE
        await Assert.That(dataRange.FirstRowUsed()).IsNull();
        await Assert.That(dataRange.LastRowUsed()).IsNull();
    }

    [Test]
    public async Task DataRange_FirstRowUsed_returns_row_when_data_exists()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        ws.Cell("A1").Value = "Col1";
        ws.Cell("B1").Value = "Col2";
        ws.Cell("A2").Value = "Data1";
        ws.Cell("B2").Value = "Data2";

        var table = ws.Range("A1:B2").CreateTable("TestTable");
        var dataRange = table.DataRange!;

        var firstRow = dataRange.FirstRowUsed();
        await Assert.That(firstRow).IsNotNull();

        var lastRow = dataRange.LastRowUsed();
        await Assert.That(lastRow).IsNotNull();
    }

    [Test]
    public async Task LoadTable_without_TableStyleInfo_sets_no_theme_and_clears_style_flags()
    {
        using var ms = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            workbookPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" }));

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");
            worksheetPart.Worksheet = new Worksheet(new SheetData(
                new Row(
                    new Cell { CellReference = "A1", DataType = CellValues.InlineString, InlineString = new InlineString(new Text("Col1")) },
                    new Cell { CellReference = "B1", DataType = CellValues.InlineString, InlineString = new InlineString(new Text("Col2")) })
                { RowIndex = 1 },
                new Row(
                    new Cell { CellReference = "A2", DataType = CellValues.InlineString, InlineString = new InlineString(new Text("A")) },
                    new Cell { CellReference = "B2", DataType = CellValues.InlineString, InlineString = new InlineString(new Text("B")) })
                { RowIndex = 2 }));

            // Add a TablePart referencing a table with NO TableStyleInfo element.
            var tableDefPart = worksheetPart.AddNewPart<TableDefinitionPart>();
            tableDefPart.Table = new Table(
                new AutoFilter { Reference = "A1:B2" },
                new TableColumns(
                    new TableColumn { Id = 1, Name = "Col1" },
                    new TableColumn { Id = 2, Name = "Col2" })
                { Count = 2 })
            {
                Id = 1,
                Name = "TestTable",
                DisplayName = "TestTable",
                Reference = "A1:B2",
                TotalsRowShown = false
            };
            // Explicitly: no TableStyleInfo child

            var tableParts = new TableParts(
                new TablePart { Id = worksheetPart.GetIdOfPart(tableDefPart) })
            { Count = 1 };
            worksheetPart.Worksheet.Append(tableParts);
        }

        ms.Position = 0;
        using var wb = new XLWorkbook(ms);
        var table = wb.Worksheets.First().Tables.First();

        await Assert.That(table.Theme).IsEqualTo(XLTableTheme.None);
        await Assert.That(table.ShowRowStripes).IsFalse();
        await Assert.That(table.ShowColumnStripes).IsFalse();
        await Assert.That(table.EmphasizeFirstColumn).IsFalse();
        await Assert.That(table.EmphasizeLastColumn).IsFalse();
    }
}
