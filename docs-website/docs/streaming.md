---
id: streaming
title: Streaming Writes for Large Files
sidebar_label: Streaming Writes
description: Write arbitrarily large .xlsx exports with flat memory using XLStreamingWorkbook — appending rows, styles, sheet layout, shared vs inline strings, and what the append-only model gives up.
---

# Streaming Writes for Large Files

`XLWorkbook` builds the entire workbook in memory and writes it out at the end. That is what makes
the rest of the API possible — you can read a cell back, edit it, insert a row above it — but it
also means the largest file you can produce is bounded by how much memory you are willing to give
it.

`XLStreamingWorkbook` is the other trade. It serialises each row into the file the moment you
append it and keeps nothing, so memory does not grow with the number of rows. In exchange it is
append-only: rows go in ascending order, one worksheet at a time, and nothing already written can
be read back or changed.

| | `XLWorkbook` | `XLStreamingWorkbook` |
|---|---|---|
| Peak memory, 1M rows × 10 columns | Needs about as much for **100K** rows | **108 MB**, or **14 MB** with inline strings |
| 50K rows × 3 columns | 250 ms, 67 MB allocated | **158 ms, 14 MB allocated** |
| Read a cell back | Yes | No |
| Edit what you wrote | Yes | No |
| Formulas evaluated | Optional | No — stored verbatim |

Reach for it when the output is a large, append-only export and memory is the constraint. For
everything else — anything you need to read, revise, or decorate with tables and pivots — stay on
[`XLWorkbook`](./importing-exporting.md).

## A first export

```csharp
using XLibur.Excel.Streaming;

using var workbook = XLStreamingWorkbook.Create("Orders.xlsx");

var sheet = workbook.AddWorksheet("Orders");
sheet.AppendRow("Order", "Customer", "Total");

foreach (var order in GetOrders())
    sheet.AppendRow(order.Reference, order.Customer, order.Total);

workbook.Finish();
```

`Create` also takes a `Stream`:

```csharp
using var file = File.Create("Orders.xlsx");
using var workbook = XLStreamingWorkbook.Create(file);
```

:::warning `Finish()` is not optional
`Finish()` writes the shared strings, the styles and the workbook part — none of which can be
known until the last row is in. A workbook that is disposed without it leaves an incomplete
package that no reader can open.

Disposal deliberately does **not** call it for you: if an exception were in flight, finishing
would bury it behind a second failure. Treat `Finish()` as the last statement of the write.
:::

## Appending rows

`AppendRow` is the short form. It takes values starting at column A:

```csharp
sheet.AppendRow("Widget", 12, 4.99, new DateTime(2026, 7, 27));
```

Values are `XLCellValue`, so strings, numbers, booleans, `DateTime`, `TimeSpan` and `XLError` all
convert implicitly — the same value model as `IXLCell.Value`. Dates and durations pick up the
number format that identifies them as such, exactly as they would on a normal cell, so they come
back as dates rather than as serial numbers.

To leave rows empty, skip them. Skipped rows cost nothing in the file:

```csharp
sheet.AppendRow("Section A");
sheet.SkipRows(2);
sheet.AppendRow("Section B");
```

`NextRowNumber` tells you where you are, which is useful for building a formula that refers to a
range you have just written:

```csharp
var firstDataRow = sheet.NextRowNumber;
foreach (var order in orders)
    sheet.AppendRow(order.Reference, order.Total);

var lastDataRow = sheet.NextRowNumber - 1;
```

## Building a row cell by cell

`AddRow()` opens a row you fill in yourself. It is the only way to write a formula or give a
single cell its own style, and it allocates nothing per row:

```csharp
using (var row = sheet.AddRow())
{
    row.Cell("Widget");
    row.Cell(12, highlight);            // this cell only
    row.Skip(1);                        // leave C empty
    row.Formula("B2*1.2", cachedValue: 14.4);
    row.At(10).Cell("note");            // jump to column J
}
```

Calls chain, if you prefer:

```csharp
sheet.AddRow().Cell("Widget").Cell(12).Cell(4.99);
```

The `using` is optional — a row also closes when the next one starts or the sheet completes — but
it makes the row's extent obvious and lets the compiler stop you using it afterwards. Writing to a
row that is no longer the open one throws `InvalidOperationException` rather than corrupting the
file.

Rows themselves can carry height, visibility and a style:

```csharp
sheet.AddRow(header, height: 24).Cell("Quarterly Report");
sheet.AddRow(style: null, hidden: true).Cell("internal");
```

### Formulas

Formula text is stored **verbatim** — it is never parsed or evaluated. A leading `=` is accepted
and stripped:

```csharp
row.Formula("SUM(B2:B100000)");             // stored as SUM(B2:B100000)
row.Formula("=SUM(B2:B100000)");            // identical
```

Without a `cachedValue` the cell has no stored result, so it shows as empty until Excel
recalculates the sheet — which it does on open. Supply the value when you already know it and the
cell displays immediately:

```csharp
row.Formula($"SUM(B{firstDataRow}:B{lastDataRow})", cachedValue: runningTotal);
```

## Styles

Styles come from `CreateStyle()`, which hands back a fresh `IXLStyle` you configure and pass to a
row or a cell:

```csharp
var header = workbook.CreateStyle();
header.Font.Bold = true;
header.Fill.BackgroundColor = XLColor.LightGray;

var money = workbook.CreateStyle();
money.NumberFormat.Format = "#,##0.00";

sheet.AppendRow(["Item", "Amount"], header);

using (var row = sheet.AddRow())
{
    row.Cell("Widget");
    row.Cell(1234.5, money);
}
```

A row style applies to every cell in that row that does not specify its own.

The writer interns a style's *value* at the moment you use it, so one instance can be reconfigured
and handed to later rows without disturbing rows already written:

```csharp
var rowStyle = workbook.CreateStyle();

foreach (var (item, isOverdue) in items)
{
    rowStyle.Font.FontColor = isOverdue ? XLColor.Red : XLColor.Black;
    sheet.AppendRow([item.Name, item.DaysOpen], rowStyle);
}
```

Distinct styles are held until `Finish()`, but their number is bounded by how many distinct formats
you actually use — not by row count. A thousand rows sharing two styles costs two.

## Sheet layout

Column widths, freeze panes and an autofilter range are all supported:

```csharp
var sheet = workbook.AddWorksheet("Orders");

sheet.Column(1).Width = 30;
sheet.Columns(2, 4).Width = 14;
sheet.Column(5).Hidden = true;

sheet.FreezeRows(1);                 // keep the header visible
sheet.AutoFilterRange = "A1:E1";

sheet.AppendRow("Order", "Customer", "Date", "Total", "Internal");
```

`FreezeColumns` and `FreezePanes(rows, columns)` cover the other two cases.

:::note Layout comes before the first row
Columns and panes are written into the file *ahead of* the rows, so they have to be set before you
append anything. Doing it afterwards throws `InvalidOperationException`, rather than silently
producing a file that ignores them.

`AutoFilterRange` is the exception — it is written after the rows, so it can be set at any point
before the sheet completes. That is handy when the range depends on how many rows you ended up
writing.
:::

## Several worksheets

Only one sheet is open at a time, because both would be writing to the same package. Adding a
worksheet completes the previous one automatically:

```csharp
var summary = workbook.AddWorksheet("Summary");
summary.AppendRow("Region", "Total");
// ...

var detail = workbook.AddWorksheet("Detail");   // Summary is completed here
detail.AppendRow("Order", "Amount");
// ...

workbook.Finish();                              // Detail is completed here
```

Sheets appear in the order you add them. You cannot go back to `summary` once `detail` has been
added — if a summary depends on totals you only learn later, write it as a formula over the detail
sheet and let Excel compute it, or make two passes over your data.

## Strings and memory

Memory is flat in the number of *rows*, but not unconditionally flat. By default text goes into a
shared string table, which holds each **distinct** string until `Finish()`. So the cost tracks how
many distinct text values you write rather than how many rows:

```csharp
// Cheap — a handful of distinct strings, however many rows
sheet.AppendRow(status, region, productName);

// Expensive — every row contributes a new string
sheet.AppendRow($"Order {Guid.NewGuid()}", description);
```

When the number of distinct strings is genuinely unbounded, switch to inline storage. Each string
is then written into its cell and nothing is retained:

```csharp
using var workbook = XLStreamingWorkbook.Create("Huge.xlsx", new XLStreamingOptions
{
    StringStorage = XLStreamingStringStorage.Inline,
});
```

| Mode | Memory | File size | Use when |
|---|---|---|---|
| `SharedStrings` (default) | Grows with distinct strings | Smaller when text repeats | Text repeats — the usual export |
| `Inline` | Flat | Larger when text repeats | Distinct strings are unbounded |

On a million rows where every row carries a distinct string — the worst case — that is the
difference between **108 MB and 14 MB**. Where text repeats, shared strings also produce a smaller
file, so the default is usually right.

## Writing to a web response

The streaming writer never seeks backwards, so its destination only has to be writable. That means
you can write straight to a response body, with neither a temporary file nor a `MemoryStream`
holding the whole workbook:

```csharp
[HttpGet("export")]
public async Task<IActionResult> Export()
{
    Response.ContentType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    Response.Headers.ContentDisposition = "attachment; filename=orders.xlsx";

    using (var workbook = XLStreamingWorkbook.Create(Response.Body))
    {
        var sheet = workbook.AddWorksheet("Orders");
        sheet.AppendRow("Order", "Customer", "Total");

        await foreach (var order in GetOrdersAsync())
            sheet.AppendRow(order.Reference, order.Customer, order.Total);

        workbook.Finish();
    }

    return new EmptyResult();
}
```

`XLWorkbook.SaveAs` cannot do this: it goes through `System.IO.Packaging`, which needs a seekable
stream and buffers the package regardless. See
[Importing and Exporting](./importing-exporting.md#streams-uploads-and-web-responses) for the
buffered equivalent.

:::note The stream stays open
`Create(Stream)` leaves the stream open when the workbook is disposed — the caller owns it. The
`Create(string)` overload owns the file it opens and closes it for you.
:::

## Compression

`CompressionLevel` trades file size against write time:

```csharp
using var workbook = XLStreamingWorkbook.Create("Fast.xlsx", new XLStreamingOptions
{
    CompressionLevel = CompressionLevel.Fastest,
});
```

`Fastest` is roughly 1.7× quicker than the default `Optimal`, for a larger file — often the right
call for a big export a user is waiting on. `SaveOptions.CompressionLevel` is the equivalent for an
ordinary save.

## What it cannot do

The append-only model rules out anything that needs to revisit what was written:

| Not supported | Use instead |
|---|---|
| Reading cells back, editing, inserting rows | `XLWorkbook` |
| Formula evaluation | Store a `cachedValue`, or let Excel recalculate on open |
| Tables, merged ranges, conditional formatting | `XLWorkbook` |
| Pivot tables, charts, images, comments | `XLWorkbook` |
| Data validation, hyperlinks, page setup | `XLWorkbook` |
| Encryption ([`SaveOptions.Password`](./encryption.md)) | `XLWorkbook` |
| Rows out of order, or a second pass over a sheet | Sort your data first |

If you need a large export *and* one of these, the usual answer is to generate the bulk with the
streaming writer and keep the decorated parts on a normal workbook in a separate file.

## A worked example

An export of arbitrary size, with a frozen and filtered header, formatted columns and a totals row
computed by Excel:

```csharp
using System.IO.Compression;
using XLibur.Excel;
using XLibur.Excel.Streaming;

using var workbook = XLStreamingWorkbook.Create("Orders.xlsx", new XLStreamingOptions
{
    CompressionLevel = CompressionLevel.Fastest,
});

var header = workbook.CreateStyle();
header.Font.Bold = true;
header.Fill.BackgroundColor = XLColor.LightGray;

var money = workbook.CreateStyle();
money.NumberFormat.Format = "#,##0.00";

var date = workbook.CreateStyle();
date.DateFormat.Format = "yyyy-MM-dd";

var sheet = workbook.AddWorksheet("Orders");
sheet.Column(1).Width = 16;
sheet.Column(2).Width = 32;
sheet.Column(3).Width = 12;
sheet.Column(4).Width = 14;
sheet.FreezeRows(1);
sheet.AutoFilterRange = "A1:D1";

sheet.AppendRow(["Order", "Customer", "Date", "Total"], header);

var firstDataRow = sheet.NextRowNumber;
var total = 0m;

foreach (var order in GetOrders())
{
    using var row = sheet.AddRow();
    row.Cell(order.Reference);
    row.Cell(order.Customer);
    row.Cell(order.Placed, date);
    row.Cell((double)order.Total, money);

    total += order.Total;
}

var lastDataRow = sheet.NextRowNumber - 1;

using (var totals = sheet.AddRow(header))
{
    totals.Cell("Total");
    totals.Skip(2);
    totals.Formula($"SUM(D{firstDataRow}:D{lastDataRow})", cachedValue: (double)total, style: money);
}

workbook.Finish();
```

## Where to next

- [Importing and Exporting](./importing-exporting.md#performance-notes) — bulk loading into a
  normal workbook, and the performance habits that apply when you stay on `XLWorkbook`
- [Workbook Settings](./workbook-settings.md#save-options) — `SaveOptions`, including
  `CompressionLevel` for ordinary saves
- [Styling](./styling.md) — the style model `CreateStyle()` returns
- [Formulas](./formulas.md) — the formula syntax stored verbatim here
