---
id: importing-exporting
title: Importing and Exporting Data
sidebar_label: Importing and Exporting
description: Bulk-load collections, DataTables, and DataSets into a workbook, and read data back out — including returning a generated file from a web response.
---

# Importing and Exporting Data

Most real workbooks are generated from data you already have: a query result, a `DataTable`, a
list of DTOs. XLibur has bulk-loading methods that are both shorter and considerably faster
than writing cells one at a time.

| Method | Writes | Creates a table |
|---|---|---|
| `InsertData` | Values only, no headers | No |
| `InsertTable` | Headers + values | Yes (optional) |
| `Worksheets.Add(DataTable)` | A whole new sheet | Yes |

## InsertData — values without a table

`InsertData` writes a collection into the sheet starting at the target cell and returns the
range it filled. It writes **no header row**:

```csharp
using XLibur.Excel;

var ws = workbook.Worksheets.Add("Data");

var numbers = new[] { 1, 2, 3, 4, 5 };
ws.Cell("A1").InsertData(numbers);           // A1:A5, one value per row

// Transposed — one value per column
ws.Cell("C1").InsertData(numbers, transpose: true);   // C1:G1
```

Collections of objects write one row per item, one column per public member:

```csharp
var people = new List<Person>
{
    new() { Name = "Ada", Age = 36 },
    new() { Name = "Grace", Age = 45 },
};

var range = ws.Cell("A2").InsertData(people);
Console.WriteLine(range!.RangeAddress);      // A2:B3
```

Collections of arrays or lists write a ragged grid — each inner collection is one row:

```csharp
var rows = new List<int[]>
{
    [1, 2, 3],
    [4, 5],
    [6, 7, 8, 9],
};

ws.Cell("A1").InsertData(rows);
```

A `DataTable`'s rows go in the same way, again without headers:

```csharp
DataTable orders = LoadOrders();
ws.Cell("A2").InsertData(orders);
```

The usual pattern is to write your own headers and then the data:

```csharp
string[] headers = ["Name", "Age"];
for (var i = 0; i < headers.Length; i++)
{
    ws.Cell(1, i + 1).Value = headers[i];
}

ws.Range(1, 1, 1, headers.Length).Style.Font.Bold = true;
ws.Cell(2, 1).InsertData(people);
```

## InsertTable — headers and a real table

`InsertTable` writes a header row from the member (or column) names and wraps the result in an
Excel table. See [Tables](./tables.md) for the full API:

```csharp
var table = ws.Cell("A1").InsertTable(people, "People", createTable: true);
```

Set `createTable: false` to get the headers and the table styling without registering a real
table object in the file:

```csharp
ws.Cell("A1").InsertTable(people, createTable: false);
```

## A DataTable as a whole sheet

`Worksheets.Add(DataTable)` creates a sheet, writes the data, and creates a table over it in
one call:

```csharp
DataTable orders = LoadOrders();

workbook.Worksheets.Add(orders);                            // sheet named after the DataTable
workbook.Worksheets.Add(orders, "Orders");                  // explicit sheet name
workbook.Worksheets.Add(orders, "Orders", "OrdersTable");   // + table name

// A sheet per table in a DataSet
workbook.Worksheets.Add(dataSet);
```

## Shaping the output with `XLColumn`

The `[XLColumn]` attribute controls headers, ordering, and exclusion for `InsertData` and
`InsertTable`:

```csharp
using XLibur.Attributes;

public class Order
{
    [XLColumn(Header = "Order ID", Order = 1)]
    public int Id { get; set; }

    [XLColumn(Header = "Customer", Order = 2)]
    public string CustomerName { get; set; } = "";

    [XLColumn(Header = "Placed on", Order = 3)]
    public DateTime OrderedAt { get; set; }

    [XLColumn(Header = "Total", Order = 4)]
    public decimal Total { get; set; }

    [XLColumn(Ignore = true)]
    public string InternalNotes { get; set; } = "";
}
```

Members without an `Order` are written after those that have one. Anonymous types work too, and
are often the simplest way to project exactly the shape you want:

```csharp
var projection = orders.Select(o => new
{
    Id = o.Id,
    Customer = o.CustomerName,
    Placed = o.OrderedAt,
    Total = o.Total,
});

ws.Cell("A1").InsertTable(projection, "Orders", createTable: true);
```

## Formatting after the fact

`InsertData` and `InsertTable` write values, not formats. Apply number formats and widths to the
returned range:

```csharp
var range = ws.Cell("A2").InsertData(orders);

if (range is not null)
{
    range.Column(3).Style.DateFormat.Format = "yyyy-MM-dd";
    range.Column(4).Style.NumberFormat.Format = "$ #,##0.00";
}

ws.Columns().AdjustToContents();
```

## Reading data back out

### Row by row

```csharp
using var workbook = new XLWorkbook("Orders.xlsx");
var ws = workbook.Worksheet("Orders");

foreach (var row in ws.RowsUsed().Skip(1))   // skip the header
{
    var order = new Order
    {
        Id = row.Cell(1).GetValue<int>(),
        CustomerName = row.Cell(2).GetString(),
        OrderedAt = row.Cell(3).GetDateTime(),
        Total = row.Cell(4).GetValue<decimal>(),
    };
}
```

### By field name, from a table

When the source is a table, read by column name and index arithmetic disappears:

```csharp
var table = ws.Table("Orders");

foreach (var row in table.DataRange!.Rows())
{
    var id = row.Field("Order ID").GetValue<int>();
    var customer = row.Field("Customer").GetString();
    var total = row.Field("Total").GetValue<decimal>();
}
```

### As a DataTable or dynamic objects

```csharp
var table = ws.Table("Orders");

DataTable dt = table.AsNativeDataTable();

foreach (var row in table.AsDynamicEnumerable())
{
    Console.WriteLine(row.Customer);
}
```

### Defensively, when the input is untrusted

Spreadsheets people have edited contain surprises. `TryGetValue<T>` keeps a bad cell from
taking down the import:

```csharp
foreach (var row in ws.RowsUsed().Skip(1))
{
    if (!row.Cell(1).TryGetValue<int>(out var id))
    {
        Console.WriteLine($"Row {row.RowNumber()}: bad id '{row.Cell(1).GetString()}'");
        continue;
    }

    if (!row.Cell(4).TryGetValue<decimal>(out var total))
    {
        total = 0;
    }

    // ...
}
```

## Streams, uploads, and web responses

Nothing needs to touch the file system. Load from a stream:

```csharp
await using var stream = File.OpenRead("Report.xlsx");
using var workbook = new XLWorkbook(stream);
```

From an uploaded file in ASP.NET Core:

```csharp
[HttpPost("import")]
public async Task<IActionResult> Import(IFormFile file)
{
    await using var stream = file.OpenReadStream();
    using var workbook = new XLWorkbook(stream);

    var ws = workbook.Worksheet(1);
    var rows = ws.RowsUsed().Skip(1).Count();

    return Ok(new { rows });
}
```

Return a generated workbook as a download:

```csharp
[HttpGet("export")]
public IActionResult Export()
{
    using var workbook = new XLWorkbook();
    var ws = workbook.Worksheets.Add("Orders");

    ws.Cell("A1").InsertTable(GetOrders(), "Orders", createTable: true);
    ws.Columns().AdjustToContents();

    using var stream = new MemoryStream();
    workbook.SaveAs(stream);

    return File(
        stream.ToArray(),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        $"orders-{DateTime.UtcNow:yyyyMMdd}.xlsx");
}
```

:::note
`SaveAs(stream)` leaves the stream at its end. Call `stream.Position = 0` before reading it
back, or use `ToArray()` on a `MemoryStream` as above.
:::

## Performance notes

:::tip When memory is the limit, not speed
Everything below makes an export cheaper *within* the in-memory model, where the whole workbook
exists before it is written. Past a few hundred thousand rows that model is itself the ceiling.

[`XLStreamingWorkbook`](./streaming.md) writes rows straight into the file as you append them, so
memory stays flat — a million rows by ten columns costs around 108 MB against roughly a gigabyte.
The trade is that it is append-only, with no reading back, no tables and no pivots.
:::

For large exports, a few habits make a substantial difference:

- **Use `InsertData` / `InsertTable`** rather than a cell-by-cell loop.
- **Style ranges, not cells.** One `ws.Range("D2:D50000").Style…` call stores one style; 50,000
  individual assignments store 50,000.
- **Set explicit column widths** instead of `AdjustToContents()` when the shape is known —
  auto-fit measures every cell's text.
- **Skip `evaluateFormulae`** unless a downstream consumer genuinely needs cached values;
  Excel will calculate on open.
- **Dispose the workbook** (`using`) so its buffers are released promptly.

```csharp
using var workbook = new XLWorkbook();
var ws = workbook.Worksheets.Add("Export");

ws.Cell("A1").InsertTable(largeCollection, "Data", createTable: true);

ws.Column(1).Width = 12;
ws.Column(2).Width = 30;
ws.Column(3).Width = 14;
ws.Range("C2:C100000").Style.NumberFormat.Format = "$ #,##0.00";

workbook.SaveAs("Export.xlsx");
```

## A worked example — round trip

```csharp
using System.Data;
using XLibur.Excel;
using XLibur.Attributes;

public class Order
{
    [XLColumn(Header = "Order ID", Order = 1)] public int Id { get; set; }
    [XLColumn(Header = "Customer", Order = 2)] public string Customer { get; set; } = "";
    [XLColumn(Header = "Placed on", Order = 3)] public DateTime Placed { get; set; }
    [XLColumn(Header = "Total", Order = 4)] public decimal Total { get; set; }
    [XLColumn(Ignore = true)] public string Notes { get; set; } = "";
}

const string path = "Orders.xlsx";

// --- Export ---
var orders = new List<Order>
{
    new() { Id = 1001, Customer = "Acme", Placed = new DateTime(2026, 1, 12), Total = 1200m },
    new() { Id = 1002, Customer = "Globex", Placed = new DateTime(2026, 1, 19), Total = 380m },
    new() { Id = 1003, Customer = "Initech", Placed = new DateTime(2026, 2, 3), Total = 4500m },
};

using (var workbook = new XLWorkbook())
{
    var ws = workbook.Worksheets.Add("Orders");

    var table = ws.Cell("A1").InsertTable(orders, "Orders", createTable: true);
    table.Theme = XLTableTheme.TableStyleMedium2;

    table.Field("Placed on").Column.Style.DateFormat.Format = "yyyy-MM-dd";
    table.Field("Total").Column.Style.NumberFormat.Format = "$ #,##0.00";

    table.ShowTotalsRow = true;
    table.Field("Order ID").TotalsRowLabel = "Totals";
    table.Field("Total").TotalsRowFunction = XLTotalsRowFunction.Sum;

    ws.Columns().AdjustToContents();
    workbook.SaveAs(path);
}

// --- Import ---
using (var workbook = new XLWorkbook(path))
{
    var table = workbook.Worksheet("Orders").Table("Orders");

    var imported = table.DataRange!.Rows().Select(row => new Order
    {
        Id = row.Field("Order ID").GetValue<int>(),
        Customer = row.Field("Customer").GetString(),
        Placed = row.Field("Placed on").GetDateTime(),
        Total = row.Field("Total").GetValue<decimal>(),
    }).ToList();

    Console.WriteLine($"{imported.Count} orders, total {imported.Sum(o => o.Total):C}");
}
```

## Where to next

- [Tables](./tables.md) — the full table API these methods produce
- [Cells and Ranges](./cells-and-ranges.md) — typed reads and the used range
- [Styling](./styling.md) — formatting the imported data
