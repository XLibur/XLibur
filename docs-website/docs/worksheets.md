---
id: worksheets
title: Worksheets
sidebar_label: Worksheets
description: Add, remove, rename, reorder, copy, and hide worksheets in an XLibur workbook.
---

# Worksheets

A workbook is a collection of worksheets. `XLWorkbook.Worksheets` is the collection API
(`Add`, `Delete`, `Contains`, indexing), while each `IXLWorksheet` exposes the operations that
act on a single sheet (`Name`, `Position`, `CopyTo`, `Delete`, `Hide`).

Two things are worth knowing up front:

- **Positions are 1-based.** `workbook.Worksheet(1)` is the leftmost tab.
- **Sheet names are case-insensitive.** `workbook.Worksheet("data")` finds a sheet named
  `Data`, and adding a second sheet called `DATA` throws.

## Adding

The simplest form takes a name:

```csharp
using var workbook = new XLWorkbook();

var sales = workbook.Worksheets.Add("Sales");
var costs = workbook.Worksheets.Add("Costs");
```

`XLWorkbook` also has `AddWorksheet` shortcuts that forward to the same collection, so both of
these are equivalent:

```csharp
var a = workbook.Worksheets.Add("Summary");
var b = workbook.AddWorksheet("Summary");   // same thing, fewer characters
```

### Inserting at a position

Pass a position to insert the sheet rather than appending it. Existing sheets at or after that
position shift right:

```csharp
workbook.Worksheets.Add("First");            // position 1
workbook.Worksheets.Add("Third");            // position 2
workbook.Worksheets.Add("Second", 2);        // inserted between them

// Order is now: First, Second, Third
```

### Auto-generated names

Omit the name and XLibur picks the next free `Sheet1`, `Sheet2`, … name:

```csharp
var sheet = workbook.Worksheets.Add();       // "Sheet1"
var next = workbook.Worksheets.Add(2);       // "Sheet2", inserted at position 2
```

### From a DataTable or DataSet

A `DataTable` can become a sheet in one call. The overloads let you control the sheet name and
the name of the Excel table that is created from the data:

```csharp
DataTable orders = LoadOrders();

workbook.Worksheets.Add(orders);                            // sheet named after the DataTable
workbook.Worksheets.Add(orders, "Orders");                  // explicit sheet name
workbook.Worksheets.Add(orders, "Orders", "OrdersTable");   // + explicit table name

// Every table in a DataSet, one sheet each
workbook.Worksheets.Add(dataSet);
```

:::note
Excel limits sheet names to 31 characters and forbids `: \ / ? * [ ]`. XLibur does not
validate this for you — a name Excel rejects produces a file Excel refuses to open, so
truncate and sanitise names that come from user input or database columns.
:::

## Finding sheets

```csharp
var byName = workbook.Worksheet("Sales");      // throws if missing
var byPosition = workbook.Worksheet(1);        // throws if missing

if (workbook.Worksheets.Contains("Sales"))
{
    // ...
}

// Non-throwing lookup
if (workbook.Worksheets.TryGetWorksheet("Sales", out var sheet))
{
    sheet.Cell("A1").Value = "Found";
}

// The collection is IEnumerable<IXLWorksheet>
foreach (var ws in workbook.Worksheets)
{
    Console.WriteLine($"{ws.Position}: {ws.Name}");
}
```

## Removing

Delete by name, by position, or from the sheet itself. Remaining sheets close the gap in
position order:

```csharp
workbook.Worksheets.Delete("Scratch");
workbook.Worksheets.Delete(3);

workbook.Worksheet("Old Data").Delete();
```

Deleting while enumerating the collection will throw, so materialise the list first:

```csharp
var temporary = workbook.Worksheets
    .Where(ws => ws.Name.StartsWith("tmp_", StringComparison.OrdinalIgnoreCase))
    .ToList();

foreach (var ws in temporary)
{
    ws.Delete();
}
```

:::warning
A workbook must contain at least one visible worksheet. Deleting the last one produces a file
Excel will refuse to open.
:::

## Moving

`Position` is settable. Assigning to it shifts every other sheet accordingly, so you never have
to renumber the rest yourself:

```csharp
var summary = workbook.Worksheet("Summary");

summary.Position = 1;                          // move to the front

var last = workbook.Worksheets.Count;
workbook.Worksheet("Appendix").Position = last; // move to the end
```

To move a sheet one place left or right:

```csharp
var ws = workbook.Worksheet("Detail");

if (ws.Position > 1)
{
    ws.Position--;                             // move left
}
```

Sorting all sheets alphabetically:

```csharp
var ordered = workbook.Worksheets.OrderBy(ws => ws.Name, StringComparer.OrdinalIgnoreCase).ToList();

for (var i = 0; i < ordered.Count; i++)
{
    ordered[i].Position = i + 1;
}
```

## Renaming

Setting `Name` also rewrites every formula and defined name that refers to the sheet, so
references do not break:

```csharp
var ws = workbook.Worksheet("Sheet1");
ws.Name = "Q1 Sales";

// A formula elsewhere reading "=Sheet1!A1" now reads "='Q1 Sales'!A1"
```

## Copying

`CopyTo` duplicates a sheet — values, formulas, formatting, tables, and merged ranges —
either within the same workbook or into a different one:

```csharp
var template = workbook.Worksheet("Template");

// Within the same workbook
var january = template.CopyTo("January");
var february = template.CopyTo("February", 2);   // and place it at position 2

// Into another workbook
using var target = new XLWorkbook();
template.CopyTo(target, "Imported Template");
```

A common pattern — one sheet per group, built from a single template:

```csharp
using var workbook = new XLWorkbook();
var template = workbook.Worksheets.Add("Template");
template.Cell("A1").Value = "Region";
template.Cell("A1").Style.Font.Bold = true;

foreach (var region in new[] { "North", "South", "East", "West" })
{
    var sheet = template.CopyTo(region);
    sheet.Cell("B1").Value = region;
}

template.Delete();   // drop the template once the copies exist
workbook.SaveAs("Regions.xlsx");
```

## Hiding

Sheets have three visibility states. `Hidden` sheets can be unhidden by the user through the
Excel UI; `VeryHidden` sheets can only be restored through VBA or code:

```csharp
var ws = workbook.Worksheet("Lookups");

ws.Hide();                                             // == Visibility = Hidden
ws.Unhide();                                           // == Visibility = Visible

ws.Visibility = XLWorksheetVisibility.VeryHidden;      // not listed in Excel's unhide dialog

if (ws.Visibility != XLWorksheetVisibility.Visible)
{
    ws.Unhide();
}
```

## Tab appearance and selection

```csharp
var ws = workbook.Worksheet("Sales");

ws.SetTabColor(XLColor.Red);
ws.TabColor = XLColor.FromHtml("#FF4F81BD");

ws.TabActive = true;      // the sheet shown when the file is opened
ws.TabSelected = true;    // part of the selected group of tabs
```

## Sheet view options

Each sheet carries its own view settings — gridlines, headers, zero display, and frozen panes:

```csharp
var ws = workbook.Worksheet("Report");

ws.ShowGridLines = false;
ws.ShowRowColHeaders = false;
ws.ShowZeros = false;
ws.SetShowFormulas(false);

// Freeze the header row and the first two columns
ws.SheetView.Freeze(1, 2);

// Or one axis at a time
ws.SheetView.FreezeRows(1);
ws.SheetView.FreezeColumns(2);

// A split (rather than frozen) view
ws.SheetView.SplitRow = 3;
ws.SheetView.SplitColumn = 3;
```

## Protecting a sheet

```csharp
var ws = workbook.Worksheet("Locked");

ws.Protect("s3cret")
  .AllowElement(XLSheetProtectionElements.SelectEverything)
  .AllowElement(XLSheetProtectionElements.FormatCells);

ws.Unprotect("s3cret");
```

:::note
Sheet protection is a UI convenience, not a security feature — the data is not encrypted and
any tool (including XLibur) can read it. To lock the *structure* of the workbook (adding,
deleting, or reordering sheets), see
[Workbook Settings](./workbook-settings.md#workbook-structure).
:::

## Default sizing and styling

Defaults apply to the whole sheet and are cheaper than styling individual cells:

```csharp
var ws = workbook.Worksheet("Data");

ws.ColumnWidth = 14;
ws.RowHeight = 18;
ws.Style.Font.FontName = "Calibri";
ws.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
```

## Where to next

- [Cells and Ranges](./cells-and-ranges.md) — addressing, reading, and writing content
- [Styling](./styling.md) — fonts, fills, borders, and alignment
- [Grouping and Outlines](./grouping-and-outlines.md) — collapsible sections of rows and columns
- [Workbook Settings](./workbook-settings.md) — document properties, protection, and save options
- [Page Setup and Printing](./page-setup.md) — print areas, headers, and scaling
