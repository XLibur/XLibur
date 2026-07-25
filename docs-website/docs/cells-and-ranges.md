---
id: cells-and-ranges
title: Cells and Ranges
sidebar_label: Cells and Ranges
description: Address cells, read and write typed values, work with ranges, merge, insert, delete, and name regions of a worksheet.
---

# Cells and Ranges

Everything you write into a workbook lands in a cell. XLibur gives you three ways to get at
one — by Excel address (`"B4"`), by row/column index, or by navigating from another cell —
and a *range* type for operating on many cells at once.

## Addressing cells

```csharp
var ws = workbook.Worksheet("Data");

var byAddress = ws.Cell("B4");
var byIndex = ws.Cell(4, 2);        // row 4, column 2 — the same cell

// Relative navigation
var right = byAddress.CellRight();       // C4
var below = byAddress.CellBelow(3);      // B7
var left = byAddress.CellLeft();         // A4
var above = byAddress.CellAbove();       // B3
```

:::note
Rows and columns are **1-based**, matching Excel. `ws.Cell(1, 1)` is `A1`.
:::

Whole rows and columns are addressed the same way:

```csharp
var row = ws.Row(1);
var column = ws.Column("B");
var columnByIndex = ws.Column(2);
```

## Writing values

`Value` accepts the types Excel understands natively. There is no boxing and no
`object` — assignment goes through the `XLCellValue` struct:

```csharp
ws.Cell("A1").Value = "Contacts";                            // Text
ws.Cell("A2").Value = 42;                                    // Number
ws.Cell("A3").Value = 3.14159;                               // Number
ws.Cell("A4").Value = true;                                  // Boolean
ws.Cell("A5").Value = new DateTime(2026, 1, 21);             // DateTime
ws.Cell("A6").Value = TimeSpan.FromHours(7.5);               // TimeSpan
ws.Cell("A7").Value = Blank.Value;                           // explicitly blank
ws.Cell("A8").FormulaA1 = "=SUM(A2:A3)";                     // formula
```

`SetValue` does the same thing but returns the cell, which makes chained writes readable:

```csharp
ws.Cell("A1").SetValue("Name")
  .CellBelow().SetValue("Alice")
  .CellBelow().SetValue("Bob")
  .CellBelow().SetValue("Carol");
```

Values from other .NET types (`int`, `decimal`, `Guid`, enums, `null`) are converted on
assignment where a sensible Excel equivalent exists. Types with no equivalent — a POCO, a
collection — are not accepted; project them to text or numbers first, or use
[InsertData / InsertTable](./importing-exporting.md).

## Reading values

```csharp
var cell = ws.Cell("B10");

XLCellValue raw = cell.Value;        // the discriminated value
XLDataType type = cell.DataType;     // Blank, Boolean, Number, Text, Error, DateTime, TimeSpan

string text = cell.GetString();      // string form regardless of underlying type
double number = cell.GetDouble();    // throws if not a number
DateTime date = cell.GetDateTime();  // throws if not a date
bool flag = cell.GetBoolean();       // throws if not a boolean
```

`GetValue<T>()` converts, and throws when it cannot. When the content is not guaranteed, use
`TryGetValue<T>`:

```csharp
if (cell.TryGetValue<decimal>(out var amount))
{
    Console.WriteLine($"Amount: {amount:C}");
}
else
{
    Console.WriteLine($"Not a number: {cell.GetString()}");
}
```

To read the value as the user sees it — number format applied — use `GetFormattedString()`:

```csharp
ws.Cell("C2").Value = 1234.5;
ws.Cell("C2").Style.NumberFormat.Format = "$ #,##0.00";

ws.Cell("C2").GetString();            // "1234.5"
ws.Cell("C2").GetFormattedString();   // "$ 1,234.50"
```

Pattern matching over `XLCellValue` handles mixed columns cleanly:

```csharp
foreach (var cell in ws.Column("B").CellsUsed())
{
    var description = cell.DataType switch
    {
        XLDataType.Text => $"text: {cell.GetString()}",
        XLDataType.Number => $"number: {cell.GetDouble()}",
        XLDataType.DateTime => $"date: {cell.GetDateTime():d}",
        XLDataType.Boolean => $"bool: {cell.GetBoolean()}",
        XLDataType.Error => $"error: {cell.Value.GetError()}",
        _ => "blank",
    };

    Console.WriteLine(description);
}
```

## Defining ranges

A range is a rectangular block of cells. All the usual constructions are available:

```csharp
var byAddress = ws.Range("A1:D10");
var byCorners = ws.Range(ws.Cell("A1"), ws.Cell("D10"));
var byIndexes = ws.Range(1, 1, 10, 4);          // firstRow, firstCol, lastRow, lastCol

var wholeRow = ws.Row(1).AsRange();
var wholeColumn = ws.Column("B").AsRange();

// Several disjoint ranges as one unit
var multi = ws.Ranges("A1:B2,D4:E5,G7");
```

### The used range

Walking the full 1,048,576 × 16,384 grid is never what you want. These give you only the
populated part of the sheet:

```csharp
var used = ws.RangeUsed();                 // null when the sheet is empty
var first = ws.FirstCellUsed();
var last = ws.LastCellUsed();

foreach (var row in ws.RowsUsed())
{
    var name = row.Cell(1).GetString();
    var qty = row.Cell(2).GetValue<int>();
    Console.WriteLine($"{name}: {qty}");
}

foreach (var cell in ws.CellsUsed())
{
    // only cells that hold content, a formula, or formatting
}
```

`CellsUsed` takes an `XLCellsUsedOptions` flag when you want to be precise about what counts
as "used":

```csharp
// Cells with content only — ignore cells that are merely formatted
foreach (var cell in ws.CellsUsed(XLCellsUsedOptions.Contents))
{
    // ...
}
```

### Iterating a range

```csharp
var range = ws.Range("A1:C5");

foreach (var row in range.Rows())
{
    foreach (var cell in row.Cells())
    {
        // ...
    }
}

// Or cell by cell
foreach (var cell in range.Cells())
{
    // ...
}
```

## Clearing

`Clear` removes everything by default. The `XLClearOptions` flags let you keep some of it:

```csharp
ws.Cell("A1").Clear();                                 // contents + all formatting
ws.Cell("A1").Clear(XLClearOptions.Contents);          // value and formula only
ws.Range("A1:D10").Clear(XLClearOptions.AllFormats);   // formatting only, keep values

// Combine flags
ws.Range("A1:D10").Clear(XLClearOptions.Contents | XLClearOptions.Comments);
```

| Option | Removes |
|---|---|
| `Contents` | Cell values and formulas |
| `NormalFormats` | Styles applied directly to the cell |
| `ConditionalFormats` | Conditional formatting rules |
| `Comments` | Cell comments |
| `DataValidation` | Validation rules |
| `MergedRanges` | Merges overlapping the range |
| `Sparklines` | Sparklines |
| `AllFormats` | `NormalFormats` + `ConditionalFormats` |
| `AllContents` | `Contents` + `Comments` |
| `All` | Everything above |

## Copying and moving

```csharp
// Copy a single cell's value, formula, and style
ws.Cell("A1").CopyTo(ws.Cell("D1"));
ws.Cell("A1").CopyTo("D1");

// Copy a range — the target is the top-left cell of the destination
ws.Range("A1:C3").CopyTo(ws.Cell("E1"));

// Across sheets
var source = workbook.Worksheet("Data").Range("A1:C10");
source.CopyTo(workbook.Worksheet("Backup").Cell("A1"));
```

## Inserting and deleting

Inserting shifts the surrounding cells, exactly as it does in Excel:

```csharp
// Whole rows and columns
ws.Row(3).InsertRowsAbove(2);
ws.Row(3).InsertRowsBelow(1);
ws.Column("B").InsertColumnsBefore(1);
ws.Column("B").InsertColumnsAfter(3);

// Or within a range only
ws.Range("B2:D5").InsertRowsAbove(1);

// Deleting
ws.Row(5).Delete();
ws.Column("C").Delete();
ws.Range("B2:D5").Delete(XLShiftDeletedCells.ShiftCellsUp);
ws.Range("B2:D5").Delete(XLShiftDeletedCells.ShiftCellsLeft);
```

:::note
Insert and delete rewrite formula references across the workbook, the same way Excel does.
`=SUM(B2:B10)` becomes `=SUM(B2:B11)` after a row is inserted inside that span.
:::

## Merging

```csharp
ws.Cell("B2").Value = "Quarterly Report";
ws.Range("B2:E2").Merge();

// Centre the merged title
ws.Cell("B2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
ws.Cell("B2").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

// Merge just one row or column of a bigger range
ws.Range("B4:D6").Row(1).Merge();
ws.Range("F2:G8").Column(1).Merge();

ws.Range("B2:E2").Unmerge();
```

The value of a merged region lives in its top-left cell; the others read as blank.

## Named ranges

Defined names make formulas readable and are the cleanest way to hand a region to a formula
that lives elsewhere:

```csharp
// Workbook scope (default)
ws.Range("A1:A10").AddToNamed("SalesFigures");

// Worksheet scope
ws.Range("C1:C10").AddToNamed("LocalRates", XLScope.Worksheet);

// With a comment
ws.Range("E1:E10").AddToNamed("TaxRates", XLScope.Workbook, "Rates by region");

ws.Cell("B1").FormulaA1 = "=SUM(SalesFigures)";

// Read them back
var name = workbook.DefinedNames.DefinedName("SalesFigures");
foreach (var range in name.Ranges)
{
    range.Style.Fill.BackgroundColor = XLColor.LightYellow;
}
```

Applying one style to a set of named ranges is a neat way to keep formatting in one place:

```csharp
ws.Cell("A1").AsRange().AddToNamed("Titles");
ws.Range("C1:H1").AddToNamed("Titles");

var titleStyle = workbook.Style;
titleStyle.Font.Bold = true;
titleStyle.Fill.BackgroundColor = XLColor.Cyan;

workbook.DefinedNames.DefinedName("Titles").Ranges.Style = titleStyle;
```

## Sorting

```csharp
// Sort the used range by its first column
ws.Sort();

// By a specific column, descending
ws.Sort(2, XLSortOrder.Descending);

// Multi-column: "2 ASC, 3 DESC"
ws.Sort("2 ASC, 3 DESC");

// Sort a range rather than the sheet
ws.Range("A2:D100").Sort(1, XLSortOrder.Ascending);
```

## Transposing

```csharp
// Swap rows and columns in place
ws.Range("A1:C5").Transpose(XLTransposeOptions.MoveCells);

// Or keep surrounding cells where they are, replacing them instead
ws.Range("A1:C5").Transpose(XLTransposeOptions.ReplaceCells);
```

## Column widths and row heights

```csharp
ws.Column("A").Width = 30;
ws.Row(1).Height = 24;

// Size to content — needs a font engine, see the Fonts page
ws.Columns().AdjustToContents();
ws.Column("B").AdjustToContents();
ws.Rows().AdjustToContents();

// Hide and group
ws.Column("D").Hide();
ws.Column("D").Unhide();
ws.Rows(3, 8).Group();
ws.Rows(3, 8).Collapse();
```

## Hyperlinks and comments

```csharp
ws.Cell("A1").Value = "XLibur on GitHub";
ws.Cell("A1").SetHyperlink(new XLHyperlink("https://github.com/XLibur/XLibur"));

// Internal link to another sheet
ws.Cell("A2").Value = "Go to Summary";
ws.Cell("A2").SetHyperlink(new XLHyperlink("Summary!A1"));

// Comment
ws.Cell("B1").CreateComment()
  .AddText("Reviewed 2026-01-21")
  .SetBold();
```

## Where to next

- [Styling](./styling.md) — fonts, fills, borders, alignment, and number formats
- [Formulas](./formulas.md) — normal, array, and dynamic array formulas
- [Importing and Exporting Data](./importing-exporting.md) — bulk-loading collections and `DataTable`s
