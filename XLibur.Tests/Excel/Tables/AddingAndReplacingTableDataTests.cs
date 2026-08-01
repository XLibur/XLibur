using XLibur.Attributes;
using XLibur.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using XLibur.Extensions;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Tables;

public class AppendingAndReplacingTableDataTests
{
    public class Person
    {
        public int Age { get; set; }

        [XLColumn(Header = "Last name", Order = 2)]
        public string LastName { get; set; }

        [XLColumn(Header = "First name", Order = 1)]
        public string FirstName { get; set; }

        [XLColumn(Header = "Full name", Order = 0)]
        public string FullName => string.Concat(FirstName, " ", LastName);

        [XLColumn(Order = 3)] public DateTime DateOfBirth { get; set; }

        [XLColumn(Header = "Is active", Order = 4)]
        public bool IsActive;
    }

    private static XLWorkbook PrepareWorkbook()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Tables");

        var persons = new[]
        {
            new Person
            {
                FirstName = "Francois", LastName = "Botha", Age = 39, DateOfBirth = new DateTime(1980, 1, 1, 0, 0, 0, DateTimeKind.Unspecified),
                IsActive = true
            },
            new Person
            {
                FirstName = "Leon", LastName = "Oosthuizen", Age = 40, DateOfBirth = new DateTime(1979, 1, 1, 0, 0, 0, DateTimeKind.Unspecified),
                IsActive = false
            },
            new Person
            {
                FirstName = "Rian", LastName = "Prinsloo", Age = 41, DateOfBirth = new DateTime(1978, 1, 1, 0, 0, 0, DateTimeKind.Unspecified),
                IsActive = false
            }
        };

        ws.FirstCell().CellRight().CellBelow().InsertTable(persons);

        ws.Columns().AdjustToContents();

        return wb;
    }

    private static XLWorkbook PrepareWorkbookWithAdditionalColumns()
    {
        var wb = PrepareWorkbook();
        var ws = wb.Worksheets.First();

        var table = ws.Tables.First();
        table.HeadersRow()!
            .LastCell().CellRight()
            .InsertData(Data, transpose: true);

        table.Resize(ws.Range(table.FirstCell(), table.LastCell().CellRight(4)));

        table.Field("CumulativeAge").DataCells.ForEach(c => c.FormulaA1 = $"SUM($G$3:G{c.WorksheetRow().RowNumber()})");
        table.Field("NameLength").DataCells.ForEach(c => c.FormulaA1 = $"LEN(B{c.WorksheetRow().RowNumber()})");
        table.Field("IsOld").DataCells.ForEach(c => c.FormulaA1 = $"=G{c.WorksheetRow().RowNumber()}>=40");
        table.Field("HardCodedValue").DataCells.Value = "40 is not old!";

        return wb;
    }

    private static Person[] NewData =>
    [
        new()
        {
            FirstName = "Michelle", LastName = "de Beer", Age = 35, DateOfBirth = new DateTime(1983, 1, 1, 0, 0, 0, DateTimeKind.Unspecified),
            IsActive = false
        },
        new()
        {
            FirstName = "Marichen", LastName = "van der Gryp", Age = 30, DateOfBirth = new DateTime(1990, 1, 1, 0, 0, 0, DateTimeKind.Unspecified),
            IsActive = true
        }
    ];

    private static readonly string[] Data = ["CumulativeAge", "NameLength", "IsOld", "HardCodedValue"];

    [Test]
    public async Task AddingEmptyEnumerables()
    {
        using var wb = PrepareWorkbook();
        var ws = wb.Worksheets.First();

        var table = ws.Tables.First();

        IEnumerable<Person> personEnumerable = [];

        await Assert.That(table.AppendData(personEnumerable)).IsNull();

        IEnumerable enumerable = Array.Empty<Person>();
        await Assert.That(table.AppendData(enumerable)).IsNull();
    }

    [Test]
    public async Task ReplaceWithEmptyEnumerables()
    {
        using var wb = PrepareWorkbook();
        var ws = wb.Worksheets.First();

        var table = ws.Tables.First();

        IEnumerable<Person> personEnumerable = [];
        await Assert.That(() => table.ReplaceData(personEnumerable)).Throws<InvalidOperationException>();

        IEnumerable enumerable = Array.Empty<Person>();
        await Assert.That(() => table.ReplaceData(enumerable)).Throws<InvalidOperationException>();
    }

    [Test]
    public async Task CanAppendTypedEnumerable()
    {
        using var ms = new MemoryStream();
        using (var wb = PrepareWorkbook())
        {
            var ws = wb.Worksheets.First();

            var table = ws.Tables.First();

            IEnumerable<Person> personEnumerable = NewData;
            var addedRange = table.AppendData(personEnumerable);

            await Assert.That(addedRange.RangeAddress.ToString()).IsEqualTo("B6:G7");
            ws.Columns().AdjustToContents();

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

            await Assert.That(table.DataRange.RowCount()).IsEqualTo(5);
            await Assert.That(table.DataRange.ColumnCount()).IsEqualTo(6);
        }
    }

    [Test]
    public async Task CanAppendToTableWithTotalsRow()
    {
        using var ms = new MemoryStream();
        using (var wb = PrepareWorkbook())
        {
            var ws = wb.Worksheets.First();

            var table = ws.Tables.First();
            table.SetShowTotalsRow(true);
            table.Fields.Last().TotalsRowFunction = XLTotalsRowFunction.Average;

            IEnumerable<Person> personEnumerable = NewData;
            var addedRange = table.AppendData(personEnumerable);

            await Assert.That(addedRange.RangeAddress.ToString()).IsEqualTo("B6:G7");
            ws.Columns().AdjustToContents();

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

            await Assert.That(table.DataRange.RowCount()).IsEqualTo(5);
            await Assert.That(table.DataRange.ColumnCount()).IsEqualTo(6);
        }
    }

    [Test]
    public async Task CanAppendTypedEnumerableAndPushDownCellsBelowTable()
    {
        using var ms = new MemoryStream();
        const string value = "Some value that will be overwritten";
        IXLAddress address;
        using (var wb = PrepareWorkbook())
        {
            var ws = wb.Worksheets.First();

            var table = ws.Tables.First();

            var cell = table.LastRow().FirstCell().CellRight(2).CellBelow(1);
            address = cell.Address;
            cell.Value = value;

            IEnumerable<Person> personEnumerable = NewData;
            var addedRange = table.AppendData(personEnumerable);

            await Assert.That(addedRange.RangeAddress.ToString()).IsEqualTo("B6:G7");
            ws.Columns().AdjustToContents();

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();

            var table = ws.Tables.First();

            var cell = ws.Cell(address);
            await Assert.That(cell.Value).IsEqualTo("de Beer");
            await Assert.That(table.DataRange.RowCount()).IsEqualTo(5);
            await Assert.That(table.DataRange.ColumnCount()).IsEqualTo(6);

            await Assert.That(cell.CellBelow(NewData.Length).Value).IsEqualTo(value);
        }
    }

    [Test]
    public async Task CanAppendUntypedEnumerable()
    {
        using var ms = new MemoryStream();
        using (var wb = PrepareWorkbook())
        {
            var ws = wb.Worksheets.First();

            var table = ws.Tables.First();

            var list = new ArrayList();
            list.AddRange(NewData);

            var addedRange = table.AppendData(list);

            await Assert.That(addedRange.RangeAddress.ToString()).IsEqualTo("B6:G7");

            ws.Columns().AdjustToContents();

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

            await Assert.That(table.DataRange.RowCount()).IsEqualTo(5);
            await Assert.That(table.DataRange.ColumnCount()).IsEqualTo(6);
        }
    }

    [Test]
    public async Task CanAppendDataTable()
    {
        using var ms = new MemoryStream();
        using (var wb = PrepareWorkbook())
        {
            var ws = wb.Worksheets.First();

            var table = ws.Tables.First();

            IEnumerable<Person> personEnumerable = NewData;

            var ws2 = wb.AddWorksheet("temp");
            var dataTable = ws2.FirstCell().InsertTable(personEnumerable).AsNativeDataTable();

            var addedRange = table.AppendData(dataTable);

            await Assert.That(addedRange.RangeAddress.ToString()).IsEqualTo("B6:G7");
            ws.Columns().AdjustToContents();

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

            await Assert.That(table.DataRange.RowCount()).IsEqualTo(5);
            await Assert.That(table.DataRange.ColumnCount()).IsEqualTo(6);
        }
    }

    [Test]
    public async Task CanReplaceWithTypedEnumerable()
    {
        using var ms = new MemoryStream();
        using (var wb = PrepareWorkbook())
        {
            var ws = wb.Worksheets.First();

            var table = ws.Tables.First();

            IEnumerable<Person> personEnumerable = NewData;
            var replacedRange = table.ReplaceData(personEnumerable);

            await Assert.That(replacedRange.RangeAddress.ToString()).IsEqualTo("B3:G4");
            ws.Columns().AdjustToContents();

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

            await Assert.That(table.DataRange.RowCount()).IsEqualTo(2);
            await Assert.That(table.DataRange.ColumnCount()).IsEqualTo(6);
        }
    }

    [Test]
    public async Task CanReplaceWithUntypedEnumerable()
    {
        using var ms = new MemoryStream();
        using (var wb = PrepareWorkbook())
        {
            var ws = wb.Worksheets.First();

            var table = ws.Tables.First();

            var list = new ArrayList();
            list.AddRange(NewData);

            var replacedRange = table.ReplaceData(list);

            await Assert.That(replacedRange.RangeAddress.ToString()).IsEqualTo("B3:G4");

            ws.Columns().AdjustToContents();

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

            await Assert.That(table.DataRange.RowCount()).IsEqualTo(2);
            await Assert.That(table.DataRange.ColumnCount()).IsEqualTo(6);
        }
    }

    [Test]
    public async Task CanReplaceWithDataTable()
    {
        using var ms = new MemoryStream();
        using (var wb = PrepareWorkbook())
        {
            var ws = wb.Worksheets.First();

            var table = ws.Tables.First();

            IEnumerable<Person> personEnumerable = NewData;

            var ws2 = wb.AddWorksheet("temp");
            var dataTable = ws2.FirstCell().InsertTable(personEnumerable).AsNativeDataTable();

            var replacedRange = table.ReplaceData(dataTable);

            await Assert.That(replacedRange.RangeAddress.ToString()).IsEqualTo("B3:G4");
            ws.Columns().AdjustToContents();

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

            await Assert.That(table.DataRange.RowCount()).IsEqualTo(2);
            await Assert.That(table.DataRange.ColumnCount()).IsEqualTo(6);
        }
    }

    [Test]
    public async Task CanReplaceToTableWithTablesRow1()
    {
        using var ms = new MemoryStream();
        using (var wb = PrepareWorkbook())
        {
            var ws = wb.Worksheets.First();

            var table = ws.Tables.First();
            table.SetShowTotalsRow(true);
            table.Fields.Last().TotalsRowFunction = XLTotalsRowFunction.Average;

            // Will cause table to overflow
            var personEnumerable = NewData.Union(NewData).Union(NewData);
            var replacedRange = table.ReplaceData(personEnumerable);

            await Assert.That(replacedRange.RangeAddress.ToString()).IsEqualTo("B3:G8");
            ws.Columns().AdjustToContents();

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

            await Assert.That(table.DataRange.RowCount()).IsEqualTo(6);
            await Assert.That(table.DataRange.ColumnCount()).IsEqualTo(6);
        }
    }

    [Test]
    public async Task CanReplaceToTableWithTablesRow2()
    {
        using var ms = new MemoryStream();
        using (var wb = PrepareWorkbook())
        {
            var ws = wb.Worksheets.First();

            var table = ws.Tables.First();
            table.SetShowTotalsRow(true);
            table.Fields.Last().TotalsRowFunction = XLTotalsRowFunction.Average;

            // Will cause the table to shrink
            var personEnumerable = NewData.Take(1);
            var replacedRange = table.ReplaceData(personEnumerable);

            await Assert.That(replacedRange.RangeAddress.ToString()).IsEqualTo("B3:G3");
            ws.Columns().AdjustToContents();

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

            await Assert.That(table.DataRange.RowCount()).IsEqualTo(1);
            await Assert.That(table.DataRange.ColumnCount()).IsEqualTo(6);
        }
    }

    [Test]
    public async Task CanReplaceWithUntypedEnumerableAndPropagateExtraColumns()
    {
        using var ms = new MemoryStream();
        using (var wb = PrepareWorkbookWithAdditionalColumns())
        {
            var ws = wb.Worksheets.First();
            var table = ws.Tables.First();

            var list = new ArrayList();
            list.AddRange(NewData);
            list.AddRange(NewData);

            var replacedRange = table.ReplaceData(list, propagateExtraColumns: true);

            await Assert.That(replacedRange.RangeAddress.ToString()).IsEqualTo("B3:G6");

            ws.Columns().AdjustToContents();

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

            await Assert.That(table.DataRange.RowCount()).IsEqualTo(4);
            await Assert.That(table.DataRange.ColumnCount()).IsEqualTo(10);

            await Assert.That(table.Worksheet.Cell("H5").FormulaA1).IsEqualTo("SUM($G$3:G5)");
            await Assert.That(table.Worksheet.Cell("H6").FormulaA1).IsEqualTo("SUM($G$3:G6)");
            await Assert.That(table.Worksheet.Cell("H5").Value).IsEqualTo(100);
            await Assert.That(table.Worksheet.Cell("H6").Value).IsEqualTo(130);

            await Assert.That(table.Worksheet.Cell("I5").FormulaA1).IsEqualTo("LEN(B5)");
            await Assert.That(table.Worksheet.Cell("I6").FormulaA1).IsEqualTo("LEN(B6)");
            await Assert.That(table.Worksheet.Cell("I5").Value).IsEqualTo(16);
            await Assert.That(table.Worksheet.Cell("I6").Value).IsEqualTo(21);

            await Assert.That(table.Worksheet.Cell("J5").FormulaA1).IsEqualTo("G5>=40");
            await Assert.That(table.Worksheet.Cell("J6").FormulaA1).IsEqualTo("G6>=40");
            await Assert.That(table.Worksheet.Cell("J5").Value).IsEqualTo(ExpectedCellValue.From(false));
            await Assert.That(table.Worksheet.Cell("J6").Value).IsEqualTo(ExpectedCellValue.From(false));

            await Assert.That(table.Worksheet.Cell("K5").Value).IsEqualTo("40 is not old!");
            await Assert.That(table.Worksheet.Cell("K6").Value).IsEqualTo("40 is not old!");
        }
    }

    [Test]
    public async Task CanReplaceWithTypedEnumerableAndPropagateExtraColumns()
    {
        using var ms = new MemoryStream();
        using (var wb = PrepareWorkbookWithAdditionalColumns())
        {
            var ws = wb.Worksheets.First();

            var table = ws.Tables.First();

            IEnumerable<Person> personEnumerable = NewData.Concat(NewData).OrderBy(p => p.Age);
            var replacedRange = table.ReplaceData(personEnumerable, propagateExtraColumns: true);

            await Assert.That(replacedRange.RangeAddress.ToString()).IsEqualTo("B3:G6");
            ws.Columns().AdjustToContents();

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

            await Assert.That(table.DataRange.RowCount()).IsEqualTo(4);
            await Assert.That(table.DataRange.ColumnCount()).IsEqualTo(10);

            await Assert.That(table.Worksheet.Cell("H5").FormulaA1).IsEqualTo("SUM($G$3:G5)");
            await Assert.That(table.Worksheet.Cell("H6").FormulaA1).IsEqualTo("SUM($G$3:G6)");
            await Assert.That(table.Worksheet.Cell("H5").Value).IsEqualTo(95);
            await Assert.That(table.Worksheet.Cell("H6").Value).IsEqualTo(130);

            await Assert.That(table.Worksheet.Cell("I5").FormulaA1).IsEqualTo("LEN(B5)");
            await Assert.That(table.Worksheet.Cell("I6").FormulaA1).IsEqualTo("LEN(B6)");
            await Assert.That(table.Worksheet.Cell("I5").Value).IsEqualTo(16);
            await Assert.That(table.Worksheet.Cell("I6").Value).IsEqualTo(16);

            await Assert.That(table.Worksheet.Cell("J5").FormulaA1).IsEqualTo("G5>=40");
            await Assert.That(table.Worksheet.Cell("J6").FormulaA1).IsEqualTo("G6>=40");
            await Assert.That(table.Worksheet.Cell("J5").Value).IsEqualTo(ExpectedCellValue.From(false));
            await Assert.That(table.Worksheet.Cell("J6").Value).IsEqualTo(ExpectedCellValue.From(false));

            await Assert.That(table.Worksheet.Cell("K5").Value).IsEqualTo("40 is not old!");
            await Assert.That(table.Worksheet.Cell("K6").Value).IsEqualTo("40 is not old!");
        }
    }

    [Test]
    [Arguments("ListOfPeople[Age]")] // Defined name formula without a A1 reference
    [Arguments("ListOfPeople!A1")] // Defined name formula with an A1 reference
    public async Task CanReplaceTableDataWhenWorksheetHasDefinedNames(string nameFormula)
    {
        // When table data are replaced, the size of a table is modified. That
        // means rows below it are shifted up/down and defined names should be
        // adjusted.
        // TODO: add assert for name shift when formulas are properly shifted. Originally, it threw even on defined name with A1 reference
        using var ms = new MemoryStream();
        using (var wb = PrepareWorkbook())
        {
            var ws = wb.Worksheets.First();

            ws.DefinedNames.Add("ListOfPeople_Age", nameFormula);

            var table = ws.Tables.First();

            IEnumerable<Person> personEnumerable = NewData;
            var replacedRange = table.ReplaceData(personEnumerable);

            await Assert.That(replacedRange.RangeAddress.ToString()).IsEqualTo("B3:G4");

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

            await Assert.That(table.DataRange.RowCount()).IsEqualTo(2);
            await Assert.That(table.DataRange.ColumnCount()).IsEqualTo(6);
        }
    }

    [Test]
    public async Task CanAppendWithUntypedEnumerableAndPropagateExtraColumns()
    {
        using var ms = new MemoryStream();
        using (var wb = PrepareWorkbookWithAdditionalColumns())
        {
            var ws = wb.Worksheets.First();
            var table = ws.Tables.First();

            var list = new ArrayList();
            list.AddRange(NewData);
            list.AddRange(NewData);

            var appendedRange = table.AppendData(list, propagateExtraColumns: true);

            await Assert.That(appendedRange.RangeAddress.ToString()).IsEqualTo("B6:G9");

            ws.Columns().AdjustToContents();

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

            await Assert.That(table.DataRange.RowCount()).IsEqualTo(7);
            await Assert.That(table.DataRange.ColumnCount()).IsEqualTo(10);

            await Assert.That(table.Worksheet.Cell("H8").FormulaA1).IsEqualTo("SUM($G$3:G8)");
            await Assert.That(table.Worksheet.Cell("H9").FormulaA1).IsEqualTo("SUM($G$3:G9)");
            await Assert.That(table.Worksheet.Cell("H8").Value).IsEqualTo(220);
            await Assert.That(table.Worksheet.Cell("H9").Value).IsEqualTo(250);

            await Assert.That(table.Worksheet.Cell("I8").FormulaA1).IsEqualTo("LEN(B8)");
            await Assert.That(table.Worksheet.Cell("I9").FormulaA1).IsEqualTo("LEN(B9)");
            await Assert.That(table.Worksheet.Cell("I8").Value).IsEqualTo(16);
            await Assert.That(table.Worksheet.Cell("I9").Value).IsEqualTo(21);

            await Assert.That(table.Worksheet.Cell("J8").FormulaA1).IsEqualTo("G8>=40");
            await Assert.That(table.Worksheet.Cell("J9").FormulaA1).IsEqualTo("G9>=40");
            await Assert.That(table.Worksheet.Cell("J8").Value).IsEqualTo(ExpectedCellValue.From(false));
            await Assert.That(table.Worksheet.Cell("J9").Value).IsEqualTo(ExpectedCellValue.From(false));

            await Assert.That(table.Worksheet.Cell("K8").Value).IsEqualTo("40 is not old!");
            await Assert.That(table.Worksheet.Cell("K9").Value).IsEqualTo("40 is not old!");
        }
    }

    [Test]
    public async Task CanAppendTypedEnumerableAndPropagateExtraColumns()
    {
        using var ms = new MemoryStream();
        using (var wb = PrepareWorkbookWithAdditionalColumns())
        {
            var ws = wb.Worksheets.First();

            var table = ws.Tables.First();

            IEnumerable<Person> personEnumerable =
                NewData
                    .Concat(NewData)
                    .Concat(NewData)
                    .OrderBy(p => p.FirstName);

            var addedRange = table.AppendData(personEnumerable);

            await Assert.That(addedRange.RangeAddress.ToString()).IsEqualTo("B6:G11");
            ws.Columns().AdjustToContents();

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

            await Assert.That(table.DataRange.RowCount()).IsEqualTo(9);
            await Assert.That(table.DataRange.ColumnCount()).IsEqualTo(10);

            await Assert.That(table.Worksheet.Cell("H10").FormulaA1).IsEqualTo("SUM($G$3:G10)");
            await Assert.That(table.Worksheet.Cell("H11").FormulaA1).IsEqualTo("SUM($G$3:G11)");
            await Assert.That(table.Worksheet.Cell("H10").Value).IsEqualTo(280);
            await Assert.That(table.Worksheet.Cell("H11").Value).IsEqualTo(315);

            await Assert.That(table.Worksheet.Cell("I10").FormulaA1).IsEqualTo("LEN(B10)");
            await Assert.That(table.Worksheet.Cell("I11").FormulaA1).IsEqualTo("LEN(B11)");
            await Assert.That(table.Worksheet.Cell("I10").Value).IsEqualTo(16);
            await Assert.That(table.Worksheet.Cell("I11").Value).IsEqualTo(16);

            await Assert.That(table.Worksheet.Cell("J10").FormulaA1).IsEqualTo("G10>=40");
            await Assert.That(table.Worksheet.Cell("J11").FormulaA1).IsEqualTo("G11>=40");
            await Assert.That(table.Worksheet.Cell("J10").Value).IsEqualTo(ExpectedCellValue.From(false));
            await Assert.That(table.Worksheet.Cell("J11").Value).IsEqualTo(ExpectedCellValue.From(false));

            await Assert.That(table.Worksheet.Cell("K10").Value).IsEqualTo("40 is not old!");
            await Assert.That(table.Worksheet.Cell("K11").Value).IsEqualTo("40 is not old!");
        }
    }
}
