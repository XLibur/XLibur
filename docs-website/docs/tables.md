---
id: tables
title: Tables
sidebar_label: Tables
description: Insert, create, resize, name, style, and query Excel tables with headers, fields, totals rows, and autofilters.
---

# Tables

An Excel *table* (`IXLTable`) is a named, structured region of a worksheet. Unlike a plain
range, a table knows its own header row, its fields, and its data rows — which means it grows
when you append data, supports structured formula references like `Sales[Amount]`, carries a
built-in visual style, and gets an autofilter for free.

There are two routes to one:

- **`InsertTable`** — write a collection or `DataTable` into the sheet *and* wrap it in a table.
- **`CreateTable`** — turn a range that already holds data into a table.

## Insert Table

`InsertTable` on a cell writes the data starting at that cell and returns the table. It accepts
`IEnumerable<T>`, `DataTable`, and even lists of primitives or arrays:

```csharp
using XLibur.Excel;

var people = new List<Person>
{
    new() { Name = "John", Age = 30, City = "London" },
    new() { Name = "Mary", Age = 15, City = "Leeds" },
    new() { Name = "Luis", Age = 21, City = "Madrid" },
};

using var workbook = new XLWorkbook();
var ws = workbook.Worksheets.Add("People");

var table = ws.Cell("A1").InsertTable(people);
```

Property names become the headers, in declaration order. From a `DataTable`, the column names
and types are used instead:

```csharp
DataTable orders = LoadOrders();
ws.Cell("A1").InsertTable(orders);
```

Simple collections work too — a `List<string>` becomes one column, a `List<int[]>` becomes a
ragged grid:

```csharp
ws.Cell("A1").InsertTable(new List<string> { "House", "Car" });

ws.Cell("C1").InsertTable(new List<int[]>
{
    [1, 2, 3],
    [1],
    [1, 2, 3, 4, 5, 6],
});
```

### Controlling headers and the table object

Two optional parameters change the shape of the result:

```csharp
// With an explicit table name
var t1 = ws.Cell("A1").InsertTable(people, "PeopleTable");

// createTable: false writes the data and styles it as a table
// in memory, but does NOT register a real Excel table in the file
var t2 = ws.Cell("A1").InsertTable(people, createTable: false);

// Name + real Excel table
var t3 = ws.Cell("A1").InsertTable(people, "PeopleTable", createTable: true);
```

### Customising columns with `XLColumn`

The `[XLColumn]` attribute controls the header text, the column order, and whether a member is
written at all:

```csharp
using XLibur.Attributes;

public class Person
{
    [XLColumn(Header = "Full name", Order = 1)]
    public string Name { get; set; } = "";

    [XLColumn(Order = 2)]
    public int Age { get; set; }

    [XLColumn(Header = "Home city", Order = 3)]
    public string City { get; set; } = "";

    [XLColumn(Ignore = true)]
    public string InternalId { get; set; } = "";
}
```

Members without an `Order` sort after those that have one.

## Create Table

When the data is already on the sheet, promote the range. The first row of the range is taken
as the header row:

```csharp
var range = ws.Range("A1:E20");

var table = range.CreateTable();             // auto-named Table1, Table2, …
var named = range.CreateTable("Contacts");   // explicit name
```

`AsTable()` is the in-memory sibling: it gives you the same field-based API for reading and
writing without registering a table in the saved file.

```csharp
var view = ws.Range("A1:E20").AsTable();     // convenience only, not persisted as a table
```

Whole-sheet shortcut:

```csharp
var table = ws.RangeUsed()!.CreateTable("Data");
```

## Finding tables

```csharp
var byName = ws.Table("Contacts");
var byIndex = ws.Table(0);                   // 0-based
var count = ws.Tables.Count();

if (ws.Tables.TryGetTable("Contacts", out var t))
{
    // ...
}

foreach (var table in ws.Tables)
{
    Console.WriteLine($"{table.Name}: {table.DataRowCount} rows");
}

// Across the whole workbook
foreach (var table in workbook.Worksheets.SelectMany(sheet => sheet.Tables))
{
    Console.WriteLine($"{table.Worksheet.Name}!{table.Name}");
}
```

## Name Table

Table names are the handle used by structured references (`Contacts[Income]`), so they are
worth setting deliberately. They must be unique within the worksheet:

```csharp
var table = ws.Range("A1:E20").CreateTable();

table.Name = "Contacts";
Console.WriteLine(table.Name);
```

```csharp
// Structured references from elsewhere on the sheet
ws.Cell("H1").FormulaA1 = "=SUM(Contacts[Income])";
ws.Cell("H2").FormulaA1 = "=COUNTA(Contacts[FName])";
```

:::warning
Renaming a table does **not** rewrite structural references that already point at the old
name. Set the name when the table is created, or fix up the formulas yourself.
:::

## Resize Table

`Resize` moves the table boundary. Rows and columns that fall outside the new boundary keep
their content but stop being part of the table:

```csharp
var table = ws.Tables.First();

// By address
table.Resize("A1:F30");
table.Resize(1, 1, 30, 6);                        // firstRow, firstCol, lastRow, lastCol

// By cells — here, trim one column and three rows off the end
table.Resize(table.FirstCell(), table.LastCell().CellLeft().CellAbove(3));

// Or grow by one column and one row
table.Resize(table.FirstCell(), table.LastCell().CellRight().CellBelow(1));

// By range
table.Resize(ws.Range("A1:F30"));
```

### Growing a table with data

Usually you do not want to resize by hand — appending data does it for you:

```csharp
// Append rows to the end of the table; the table boundary expands
table.AppendData(newPeople);
table.AppendData(newOrdersDataTable);

// Replace the data rows entirely, resizing to fit
table.ReplaceData(allPeople);
```

`propagateExtraColumns: true` copies the values and formulas of any columns you added beside
the imported data into the new rows:

```csharp
table.AppendData(newPeople, propagateExtraColumns: true);
```

Inserting a column inside the table also widens it:

```csharp
table.Field("String").Column.InsertColumnsAfter(1, expandRange: true);
```

## Theme Table

Tables carry one of Excel's built-in table styles. `XLTableTheme` exposes them as static
fields — 60 named themes plus `None`:

```csharp
table.Theme = XLTableTheme.TableStyleMedium2;
table.Theme = XLTableTheme.TableStyleLight10;
table.Theme = XLTableTheme.TableStyleDark3;
table.Theme = XLTableTheme.None;               // no built-in styling
```

The names follow Excel's own scheme: `TableStyleLight1`–`21`, `TableStyleMedium1`–`28`, and
`TableStyleDark1`–`11`. Light styles are subtle outlines, Medium styles are the banded ones
most reports use, Dark styles are solid-fill.

Four switches control which parts of the theme are drawn:

```csharp
table.ShowRowStripes = true;         // alternate row shading
table.ShowColumnStripes = false;     // alternate column shading
table.EmphasizeFirstColumn = true;   // bold/filled first column
table.EmphasizeLastColumn = true;    // bold/filled last column
```

All the setters have fluent forms, which chain nicely at creation time:

```csharp
var table = ws.Cell("B2").InsertTable(data, "Sales", createTable: true)
    .SetShowHeaderRow()
    .SetShowTotalsRow()
    .SetShowRowStripes()
    .SetEmphasizeFirstColumn();

table.Theme = XLTableTheme.TableStyleMedium9;
```

Theme colours are resolved from the workbook theme, so changing the theme's accent colours
restyles every table at once — see [Theming](./theming.md).

## Headers and Fields

A *field* is a table column: its header cell, its data cells, and its totals cell.

```csharp
// The header row as a range (null when ShowHeaderRow is false)
var headers = table.HeadersRow();

foreach (var cell in headers!.Cells())
{
    Console.WriteLine(cell.GetString());
}

// Fields by name (case-insensitive) or by 0-based index
var income = table.Field("Income");
var first = table.Field(0);

foreach (var field in table.Fields)
{
    Console.WriteLine($"{field.Index}: {field.Name}");
}
```

### Renaming a column

Setting `Name` on a field rewrites the header cell:

```csharp
table.Field("FName").Name = "First name";
```

### Reading data by field name

Within the data range, `row.Field("Name")` gets the cell in that column — no index arithmetic:

```csharp
foreach (var row in table.DataRange!.Rows())
{
    var firstName = row.Field("FName").GetString();
    var lastName = row.Field("LName").GetString();
    var income = row.Field("Income").GetValue<decimal>();

    Console.WriteLine($"{firstName} {lastName}: {income:C}");
}
```

### Field-level access to cells and styling

```csharp
var field = table.Field("Income");

field.Column.Style.NumberFormat.Format = "$ #,##0.00";   // header + data + totals
foreach (var cell in field.DataCells)                     // data rows only
{
    // ...
}

field.HeaderCell!.Style.Font.Italic = true;
```

### Showing and hiding the header

```csharp
table.ShowHeaderRow = false;   // the header row is removed from the sheet
table.ShowHeaderRow = true;    // and restored, with the field names
```

### The totals row

Enable `ShowTotalsRow`, then set a function per field. `TotalsRowLabel` puts free text in a
cell (typically the leftmost one):

```csharp
table.ShowTotalsRow = true;

table.Field(0).TotalsRowLabel = "Totals";
table.Field("Income").TotalsRowFunction = XLTotalsRowFunction.Sum;
table.Field("Age").TotalsRowFunction = XLTotalsRowFunction.Average;
table.Field("Orders").TotalsRowFunction = XLTotalsRowFunction.CountNumbers;
```

Available functions: `None`, `Sum`, `Minimum`, `Maximum`, `Average`, `Count`, `CountNumbers`,
`StandardDeviation`, `Variance`, `Custom`.

For anything the built-in list does not cover, write a formula directly — the function is set
to `Custom` for you:

```csharp
table.Field(0).TotalsRowFormulaA1 = "CONCATENATE(\"Count: \", COUNTA(Contacts[FName]))";
```

The totals row itself is a range like any other:

```csharp
table.TotalsRow()!.Style.Font.Bold = true;
```

## AutoFilter

Tables get an autofilter automatically. `ShowAutoFilter` controls the dropdown arrows, and
`table.AutoFilter` is the same API described in [AutoFilter](./autofilter.md):

```csharp
table.ShowAutoFilter = true;      // arrows visible (default)
table.ShowAutoFilter = false;     // hide them — useful for print-oriented tables

// Filter to two values in the first column
table.AutoFilter.Column(1).AddFilter(3).AddFilter(4);

// Sort by the first column
table.AutoFilter.Sort();
table.AutoFilter.Sort(2, XLSortOrder.Descending);

// Clear all filters
table.AutoFilter.Clear();
```

:::note
Filter column numbers are **relative to the table**, not to the sheet. For a table starting at
`B2`, `Column(1)` is sheet column `B`. The same applies to the `Column("A")` letter overload.
:::

:::note
A worksheet can have at most one *sheet-level* autofilter, but every table on the sheet has its
own. Prefer table autofilters when a sheet holds more than one data block.
:::

## Ranges within a table

```csharp
table.RangeAddress          // the whole table, including header and totals
table.DataRange             // data rows only — null if the table has no data rows
table.HeadersRow()          // header row, or null
table.TotalsRow()           // totals row, or null
table.DataRowCount          // number of data rows
```

## Exporting a table

```csharp
// As a System.Data.DataTable
DataTable dt = table.AsNativeDataTable();

// As dynamic objects, one per row, with properties named after the fields
foreach (var row in table.AsDynamicEnumerable())
{
    Console.WriteLine(row.Income);
}
```

## Copying and removing

```csharp
// Copy a table onto another sheet
var copy = table.CopyTo(workbook.Worksheet("Backup"));

// Clear the contents but keep the table
table.Clear(XLClearOptions.Contents);

// Remove the table definition (the cells stay)
ws.Tables.Remove("Contacts");
```

## A worked example

```csharp
using XLibur.Excel;

using var workbook = new XLWorkbook();
var ws = workbook.Worksheets.Add("Sales");

var sales = Enumerable.Range(1, 10).Select(i => new
{
    Id = i,
    Product = $"Product {(char)('A' + i - 1)}",
    Units = i * 7,
    UnitPrice = 9.99m + i,
});

var table = ws.Cell("B2").InsertTable(sales, "SalesTable", createTable: true)
    .SetShowHeaderRow()
    .SetShowTotalsRow()
    .SetShowRowStripes();

table.Theme = XLTableTheme.TableStyleMedium9;

table.Field("UnitPrice").Column.Style.NumberFormat.Format = "$ #,##0.00";

table.Field(0).TotalsRowLabel = "Totals";
table.Field("Units").TotalsRowFunction = XLTotalsRowFunction.Sum;
table.Field("UnitPrice").TotalsRowFunction = XLTotalsRowFunction.Average;

table.AutoFilter.Column(2).Contains("Product");   // column 2 of the table == "Product"
table.AutoFilter.Sort(3, XLSortOrder.Descending); // sort by "Units"

ws.Columns().AdjustToContents();
workbook.SaveAs("SalesTable.xlsx");
```

## Where to next

- [AutoFilter](./autofilter.md) — the full filter and sort API
- [Pivot Tables](./pivot-tables.md) — tables are the cleanest pivot data source
- [Importing and Exporting Data](./importing-exporting.md) — `InsertData` when you don't want a table
