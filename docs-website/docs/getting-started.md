---
id: getting-started
title: Getting Started
sidebar_label: Getting Started
sidebar_position: 2
description: Install XLibur from NuGet and learn the basics — creating, loading, editing, and saving Excel workbooks.
---

# Getting Started

This page takes you from an empty project to reading and writing `.xlsx` files.

## Requirements

- .NET 8, .NET 9, or .NET 10
- No installation of Microsoft Excel is required — XLibur writes the file format directly

## Installation

The recommended package is **`XLibur.Bundle`**, which installs the core library together with the
default font engine and behaves like ClosedXML out of the box.

```sh
dotnet add package XLibur.Bundle
```

Or from the Visual Studio Package Manager Console:

```powershell
PM> Install-Package XLibur.Bundle
```

<details>
<summary>Why a bundle package?</summary>

XLibur keeps its core assembly free of any font library, and ships the font engine
(needed for text measurement — column auto-fit, row heights, glyph metrics) as a separate,
swappable package. `XLibur.Bundle` = `XLibur` + `XLibur.Fonts.SkiaSharp` (MIT licensed),
which is auto-registered the first time you create a workbook.

If you install the bare `XLibur` package with no font engine, creating a workbook throws an
`InvalidOperationException` telling you to add a font engine package. See the
[Migration from ClosedXML](./migration.md#font-engine-configuration-different-from-closedxml) for
the list of available engines.

</details>

Every example below assumes this using directive:

```csharp
using XLibur.Excel;
```

## Creating a workbook

Create a workbook, add a worksheet, write a couple of cells, and save it to disk:

```csharp
using XLibur.Excel;

using var workbook = new XLWorkbook();
var worksheet = workbook.Worksheets.Add("Sample Sheet");

worksheet.Cell("A1").Value = "Hello World!";
worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";

workbook.SaveAs("HelloWorld.xlsx");
```

`XLWorkbook` implements `IDisposable`, so prefer a `using` declaration (or `using` block) to
release the underlying resources when you are done.

## Loading an existing workbook

Pass a file path — or a `Stream` — to the `XLWorkbook` constructor:

```csharp
using var workbook = new XLWorkbook("Report.xlsx");

// By name...
var sheet = workbook.Worksheet("Sample Sheet");

// ...or by 1-based position
var firstSheet = workbook.Worksheet(1);
```

Loading from a stream is useful for uploaded files or blobs:

```csharp
await using var stream = File.OpenRead("Report.xlsx");
using var workbook = new XLWorkbook(stream);
```

:::note
Worksheet, row, and column indexes in XLibur are **1-based**, matching Excel itself —
`workbook.Worksheet(1)` is the first sheet, and `worksheet.Cell(1, 1)` is cell `A1`.
:::

## Reading cell values

Address a cell by its Excel address or by row/column index:

```csharp
var worksheet = workbook.Worksheet("Sample Sheet");

string title = worksheet.Cell("A1").GetString();
double total = worksheet.Cell("B10").GetValue<double>();
var cellByIndex = worksheet.Cell(1, 2); // row 1, column 2 == B1
```

`GetValue<T>()` throws if the value cannot be converted. When the content is not guaranteed,
use `TryGetValue<T>` instead:

```csharp
if (worksheet.Cell("B10").TryGetValue<double>(out var amount))
{
    Console.WriteLine($"Amount: {amount}");
}
```

To iterate only the populated part of a sheet, use `RowsUsed()` / `CellsUsed()` rather than
walking the full 1,048,576-row grid:

```csharp
foreach (var row in worksheet.RowsUsed())
{
    var name = row.Cell(1).GetString();
    var qty = row.Cell(2).GetValue<int>();
    Console.WriteLine($"{name}: {qty}");
}
```

## Amending cells

Assigning to `Cell(...).Value` sets the cell content. The value is strongly typed —
text, numbers, booleans, dates, and `TimeSpan` are all supported directly:

```csharp
var ws = workbook.Worksheet("Sample Sheet");

ws.Cell("A1").Value = "Contacts";                 // text
ws.Cell("B2").Value = 42;                         // number
ws.Cell("C2").Value = true;                       // boolean
ws.Cell("D2").Value = new DateTime(2026, 1, 21);  // date
ws.Cell("E2").FormulaA1 = "=SUM(B2:B10)";         // formula

ws.Cell("F2").Clear();                            // remove content and formatting
ws.Cell("F3").Clear(XLClearOptions.Contents);     // remove content only
```

Formatting is applied through `Style`, either per cell or across a range:

```csharp
var headers = ws.Range("A1:D1");
headers.Style.Font.Bold = true;
headers.Style.Fill.BackgroundColor = XLColor.Aqua;
headers.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

ws.Range("D2:D10").Style.NumberFormat.Format = "$ #,##0.00";
ws.Range("C2:C10").Style.DateFormat.Format = "yyyy-MM-dd";

ws.Columns().AdjustToContents(); // size columns to fit their content
```

## Saving

`SaveAs` writes to a new location; `Save` writes back to the file (or stream) the workbook
was loaded from:

```csharp
// Edit an existing file in place
using (var workbook = new XLWorkbook("Report.xlsx"))
{
    workbook.Worksheet(1).Cell("A1").Value = "Updated";
    workbook.Save();
}

// Or write a copy
workbook.SaveAs("Report-2026.xlsx");
```

To return a workbook from a web request, save it to a `MemoryStream` — no temporary
file needed:

```csharp
using var workbook = new XLWorkbook();
var ws = workbook.Worksheets.Add("Data");
ws.Cell("A1").Value = "Generated on the server";

using var stream = new MemoryStream();
workbook.SaveAs(stream);

return File(
    stream.ToArray(),
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "report.xlsx");
```

## Putting it together

A small end-to-end example — build a report, save it, reopen it, and amend a cell:

```csharp
using XLibur.Excel;

const string path = "SalesReport.xlsx";

// 1. Create
using (var workbook = new XLWorkbook())
{
    var ws = workbook.Worksheets.Add("Sales");

    ws.Cell("A1").Value = "Product";
    ws.Cell("B1").Value = "Units";
    ws.Cell("C1").Value = "Unit Price";
    ws.Cell("D1").Value = "Total";
    ws.Range("A1:D1").Style.Font.Bold = true;

    var products = new[]
    {
        ("Widget", 12, 9.99),
        ("Gadget", 4, 24.50),
        ("Doohickey", 27, 3.75),
    };

    var row = 2;
    foreach (var (name, units, price) in products)
    {
        ws.Cell(row, 1).Value = name;
        ws.Cell(row, 2).Value = units;
        ws.Cell(row, 3).Value = price;
        ws.Cell(row, 4).FormulaA1 = $"=B{row}*C{row}";
        row++;
    }

    ws.Range($"C2:D{row - 1}").Style.NumberFormat.Format = "$ #,##0.00";
    ws.Columns().AdjustToContents();

    workbook.SaveAs(path);
}

// 2. Reopen and amend
using (var workbook = new XLWorkbook(path))
{
    var ws = workbook.Worksheet("Sales");
    ws.Cell("A5").Value = "Thingamajig";
    ws.Cell("B5").Value = 8;
    ws.Cell("C5").Value = 14.25;
    ws.Cell("D5").FormulaA1 = "=B5*C5";
    workbook.Save();
}
```

## Where to next

- The [ClosedXML documentation](https://closedxml.github.io/ClosedXML/) covers the wider API
  (tables, pivot tables, conditional formatting, charts, page setup) and is still *mostly*
  valid for XLibur — the API surface is largely unchanged from ClosedXML 0.105.
- The [`XLibur.Examples`](https://github.com/XLibur/XLibur/tree/main/XLibur.Examples) project in
  the repository contains runnable samples for most features.
- Migrating an existing ClosedXML project? See
  [Migration from ClosedXML](./migration.md).
