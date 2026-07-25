---
id: autofilter
title: AutoFilter
sidebar_label: AutoFilter
description: Define autofilters on ranges and tables — value, custom, top/bottom, dynamic, date-group, and colour filters, plus sorting.
---

# AutoFilter

An autofilter turns a block of data into something the user can filter and sort from the
Excel UI. XLibur does more than write the dropdown arrows into the file: it *evaluates* the
filter conditions and marks the rows that fail them as hidden, so the workbook you save
already shows the filtered result.

The shape is always the same: get an `IXLAutoFilter`, pick a column, apply a condition.

## Define an autofilter

Call `SetAutoFilter()` on a range. The first row of the range becomes the header row:

```csharp
using XLibur.Excel;

using var workbook = new XLWorkbook();
var ws = workbook.Worksheets.Add("Data");

ws.Cell("A1").SetValue("Product")
  .CellBelow().SetValue("Widget")
  .CellBelow().SetValue("Gadget")
  .CellBelow().SetValue("Doohickey");

ws.Cell("B1").SetValue("Units")
  .CellBelow().SetValue(120)
  .CellBelow().SetValue(45)
  .CellBelow().SetValue(310);

var autoFilter = ws.RangeUsed()!.SetAutoFilter();
```

You can also filter a fixed range, or reach the sheet's filter afterwards through
`ws.AutoFilter`:

```csharp
ws.Range("A1:D100").SetAutoFilter();

var filter = ws.AutoFilter;               // the sheet's single autofilter
Console.WriteLine(filter.IsEnabled);
Console.WriteLine(filter.Range.RangeAddress.ToString());
```

Tables have their own autofilter, independent of the sheet's:

```csharp
var table = ws.Range("A1:D100").CreateTable("Sales");
var tableFilter = table.AutoFilter;
```

:::note
A worksheet has **at most one** sheet-level autofilter. Calling `SetAutoFilter()` on a second
range replaces the first. Each table, however, carries its own — so use tables when a sheet
holds several independently filtered blocks.
:::

### Enabling and clearing

```csharp
ws.AutoFilter.IsEnabled = false;    // hide the arrows, show every row
ws.AutoFilter.IsEnabled = true;

ws.AutoFilter.Clear();              // remove all filters and unhide every row
ws.AutoFilter.Column(2).Clear();    // clear one column's filters only
```

### Addressing columns

Filter columns are numbered **relative to the autofilter range**, starting at 1:

```csharp
ws.AutoFilter.Column(1);      // first column of the range
ws.AutoFilter.Column("A");    // same thing — "A" is the range's first column
```

For a filter on `C5:H100`, `Column(1)` is sheet column `C`.

## Value filters

The everyday case: a checkbox list of allowed values. Each `AddFilter` adds one permitted
value; a row is visible if its cell matches any of them.

```csharp
var autoFilter = ws.RangeUsed()!.SetAutoFilter();

autoFilter.Column(1).AddFilter("Widget")
                    .AddFilter("Gadget");

autoFilter.Column(2).AddFilter(120)
                    .AddFilter(310);
```

:::warning
Value filters compare the cell's **formatted string** against the filter value converted to a
string using the current culture. A cell holding `2.5` formatted as `2.50` will not match
`AddFilter(2.5)`, and results differ between locales. When filtering numbers, prefer the
custom comparison filters below.
:::

## Custom filters

Comparison-based conditions. Each returns a connector so you can add a second condition with
`.Or` or `.And` — Excel allows exactly two per column:

```csharp
var autoFilter = ws.RangeUsed()!.SetAutoFilter();

// Numeric comparisons
autoFilter.Column(2).EqualTo(3).Or.GreaterThan(4);
autoFilter.Column(2).EqualOrGreaterThan(100).And.LessThan(500);

// Range comparisons (no connector — these stand alone)
autoFilter.Column(2).Between(100, 500);
autoFilter.Column(2).NotBetween(100, 500);
```

Text conditions:

```csharp
autoFilter.Column(1).BeginsWith("J");
autoFilter.Column(1).EndsWith("son");
autoFilter.Column(1).Contains("dget");
autoFilter.Column(1).NotContains("test");
autoFilter.Column(1).NotBeginsWith("_");

// Combined
autoFilter.Column(1).BeginsWith("A").Or.BeginsWith("B");
```

The full comparison set: `EqualTo`, `NotEqualTo`, `GreaterThan`, `LessThan`,
`EqualOrGreaterThan`, `EqualOrLessThan`, `Between`, `NotBetween`, `BeginsWith`,
`NotBeginsWith`, `EndsWith`, `NotEndsWith`, `Contains`, `NotContains`.

## Top and bottom filters

Show only the highest or lowest values, by count or by percentile:

```csharp
autoFilter.Column(2).Top(10);                              // top 10 items
autoFilter.Column(2).Top(10, XLTopBottomType.Percent);     // top 10%
autoFilter.Column(2).Bottom(5);                            // bottom 5 items
autoFilter.Column(2).Bottom(50, XLTopBottomType.Percent);  // bottom half
```

## Dynamic filters

Computed against the column's own data:

```csharp
autoFilter.Column(2).AboveAverage();
autoFilter.Column(2).BelowAverage();
```

## Date group filters

Filter dates at a chosen granularity. `XLDateTimeGrouping` selects how much of the date is
compared — everything more precise than the grouping level is ignored:

```csharp
var target = new DateTime(2018, 1, 4);

autoFilter.Column(1).AddDateGroupFilter(target, XLDateTimeGrouping.Day);     // that exact day
autoFilter.Column(1).AddDateGroupFilter(target, XLDateTimeGrouping.Month);   // all of January 2018
autoFilter.Column(1).AddDateGroupFilter(target, XLDateTimeGrouping.Year);    // all of 2018
```

Grouping levels: `Year`, `Month`, `Day`, `Hour`, `Minute`, `Second`.

Like value filters, these accumulate — call it more than once to allow several periods:

```csharp
autoFilter.Column(1)
    .AddDateGroupFilter(new DateTime(2018, 1, 1), XLDateTimeGrouping.Month)
    .AddDateGroupFilter(new DateTime(2018, 3, 1), XLDateTimeGrouping.Month);
```

## Colour filters

Filter by fill colour or font colour:

```csharp
autoFilter.Column(3).ColorFilter(XLColor.Yellow);       // cells with a yellow fill
autoFilter.Column(3).FontColorFilter(XLColor.Red);      // cells with red text
```

## One filter type per column

A column carries a single filter *type*. Applying a different type replaces what was there —
`AddFilter` after `Top(10)` discards the top-10 rule rather than combining with it:

```csharp
autoFilter.Column(2).Top(10);
autoFilter.Column(2).AddFilter(5);   // the Top(10) filter is gone
```

Different columns combine with AND: a row must satisfy every column's filter to stay visible.

```csharp
autoFilter.Column(1).BeginsWith("J");     // name starts with J
autoFilter.Column(2).GreaterThan(100);    // AND units > 100
```

## Sorting

`Sort` orders the rows of the autofilter range by one column. It also records the sort
state in the file, so Excel shows the sort indicator on that column:

```csharp
ws.AutoFilter.Sort();                                      // column 1, ascending
ws.AutoFilter.Sort(2);                                     // by column 2
ws.AutoFilter.Sort(2, XLSortOrder.Descending);
ws.AutoFilter.Sort(2, XLSortOrder.Ascending, matchCase: true, ignoreBlanks: false);
```

| Parameter | Effect |
|---|---|
| `columnToSortBy` | 1-based column within the autofilter range |
| `sortOrder` | `Ascending` (default) or `Descending` |
| `matchCase` | Case-sensitive text comparison; default `false` |
| `ignoreBlanks` | `true` (default) puts blanks last regardless of order; `false` sorts them as empty strings |

Reading the sort state back:

```csharp
if (ws.AutoFilter.Sorted)
{
    Console.WriteLine($"Sorted by column {ws.AutoFilter.SortColumn} {ws.AutoFilter.SortOrder}");
}
```

## Inspecting the result

Because XLibur applies the filters, you can read which rows survived:

```csharp
foreach (var row in ws.AutoFilter.VisibleRows)
{
    Console.WriteLine(row.Cell(1).GetString());
}

Console.WriteLine($"{ws.AutoFilter.HiddenRows.Count()} rows filtered out");
```

## Reapplying after edits

Filters are re-evaluated automatically whenever the filter configuration changes. They are
*not* re-evaluated when you change cell values or delete rows afterwards — call `Reapply()`
then:

```csharp
ws.Cell("B3").Value = 999;
ws.AutoFilter.Reapply();
```

Every filter method also takes a `reapply` flag. Setting it to `false` on all but the last
call avoids re-filtering the range once per condition:

```csharp
var column = ws.AutoFilter.Column(2);
column.AddFilter(1, reapply: false);
column.AddFilter(2, reapply: false);
column.AddFilter(3, reapply: true);   // evaluate once, at the end
```

## Autofilters and column widths

`AdjustToContents()` sizes columns to their content but does not account for the space the
filter dropdown arrow needs, so the arrow can overlap the header text. Add a little padding
to filtered columns:

```csharp
ws.Columns().AdjustToContents();

foreach (var column in ws.AutoFilter.Range.Columns())
{
    var sheetColumn = ws.Column(column.RangeAddress.FirstAddress.ColumnNumber);
    sheetColumn.Width += 3;
}
```

## A worked example

```csharp
using XLibur.Excel;

using var workbook = new XLWorkbook();
var ws = workbook.Worksheets.Add("Orders");

ws.Cell("A1").Value = "Customer";
ws.Cell("B1").Value = "Region";
ws.Cell("C1").Value = "Amount";
ws.Cell("D1").Value = "Ordered";

var orders = new[]
{
    ("Acme", "North", 1200m, new DateTime(2026, 1, 12)),
    ("Globex", "South", 380m, new DateTime(2026, 1, 19)),
    ("Initech", "North", 4500m, new DateTime(2026, 2, 3)),
    ("Umbrella", "East", 910m, new DateTime(2026, 2, 21)),
    ("Soylent", "North", 2750m, new DateTime(2026, 3, 8)),
};

var row = 2;
foreach (var (customer, region, amount, ordered) in orders)
{
    ws.Cell(row, 1).Value = customer;
    ws.Cell(row, 2).Value = region;
    ws.Cell(row, 3).Value = amount;
    ws.Cell(row, 4).Value = ordered;
    row++;
}

ws.Range($"C2:C{row - 1}").Style.NumberFormat.Format = "$ #,##0.00";
ws.Range($"D2:D{row - 1}").Style.DateFormat.Format = "yyyy-MM-dd";

var filter = ws.RangeUsed()!.SetAutoFilter();

filter.Column(2).AddFilter("North");          // North region only
filter.Column(3).GreaterThan(1000);           // AND over $1,000

ws.AutoFilter.Sort(3, XLSortOrder.Descending);

ws.Range("A1:D1").Style.Font.Bold = true;
ws.Columns().AdjustToContents();

workbook.SaveAs("FilteredOrders.xlsx");
```

## Where to next

- [Tables](./tables.md) — table-scoped autofilters and totals rows
- [Cells and Ranges](./cells-and-ranges.md) — range sorting outside an autofilter
