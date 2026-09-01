---
id: grouping-and-outlines
title: Grouping and Outlines
sidebar_label: Grouping and Outlines
description: Group rows and columns into collapsible outline levels, nest them, and control where the summary row or column sits.
---

# Grouping and Outlines

Grouping turns a block of rows or columns into a collapsible section, with the `+`/`−` buttons
and level selectors down the left edge (or across the top). It is the standard way to ship a
detailed report that opens summarised — the reader expands only the sections they care about.

```csharp
using XLibur.Excel;

var ws = workbook.Worksheet("Report");

ws.Rows(3, 8).Group();       // group rows 3-8
ws.Rows(3, 8).Collapse();    // and start collapsed
```

## Grouping rows

```csharp
ws.Rows(3, 8).Group();                    // one outline level deeper
ws.Rows(3, 8).Group(collapse: true);      // group and collapse in one call
ws.Rows(3, 8).Group(outlineLevel: 2);     // set an explicit level
ws.Rows(3, 8).Group(2, collapse: true);

// A single row
ws.Row(5).Group();
```

Ungrouping removes one level; `fromAll: true` removes every level at once:

```csharp
ws.Rows(3, 8).Ungroup();
ws.Rows(3, 8).Ungroup(fromAll: true);
```

## Grouping columns

Identical API on columns:

```csharp
ws.Columns(3, 8).Group();
ws.Columns("C", "H").Group(collapse: true);
ws.Column("D").Group(2);

ws.Columns(3, 8).Ungroup(fromAll: true);
```

## Collapsing and expanding

At the group level:

```csharp
ws.Rows(3, 8).Collapse();
ws.Rows(3, 8).Expand();
```

Or across the whole sheet, optionally targeting one outline level:

```csharp
ws.CollapseRows();            // collapse every grouped row
ws.ExpandRows();
ws.CollapseColumns();
ws.ExpandColumns();

ws.CollapseRows(2);           // collapse only level 2 and deeper
ws.ExpandRows(1);
```

This is how you control what the reader sees on open — build the sheet fully expanded, then
collapse the detail as the last step:

```csharp
// ... build the report ...
ws.CollapseRows(2);   // show level 1 summaries, hide the detail beneath
```

## Outline levels

Levels nest, up to Excel's maximum of 7. Grouping an inner block first and then an outer block
containing it produces the nesting you would get by hand:

```csharp
// Detail rows for two sub-sections
ws.Rows(4, 7).Group();      // level 1
ws.Rows(10, 13).Group();    // level 1

// Then the whole section containing both
ws.Rows(3, 14).Group();     // level 1 for the outer, pushing the inner ones to level 2
```

Reading and setting the level directly:

```csharp
Console.WriteLine(ws.Row(5).OutlineLevel);
ws.Row(5).OutlineLevel = 2;
ws.Column("D").OutlineLevel = 1;
```

:::note
Setting `OutlineLevel = 0` removes the row or column from the outline entirely — the same as
`Ungroup(fromAll: true)`.
:::

:::caution Files written before this release carry a wrong outline summary
Setting a row's outline level used to count into the *column* outline tally. Two things followed:
a sheet with grouped rows declared `sheetFormatPr/@outlineLevelRow="0"`, and every row group
instead raised `@outlineLevelCol`, so a sheet with no grouped columns could still claim column
outlines. Because loading a file also sets each row's outline level, opening a file with row
groups and re-saving it inflated that file's `@outlineLevelCol` a little further each time.

Opening such a file and saving it with a current build corrects both attributes. Per-row
`row/@outlineLevel` was always written correctly, so only the sheet-level summary was ever wrong
— what Excel renders is unaffected.
:::

## Where the summary sits

Excel assumes the summary row is *below* its detail and the summary column *right* of it, and
places the collapse button accordingly. If your report puts totals at the top, tell the sheet:

```csharp
ws.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;    // totals above the detail
ws.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Bottom; // default

ws.Outline.SummaryHLocation = XLOutlineSummaryHLocation.Left;   // totals left of the detail
ws.Outline.SummaryHLocation = XLOutlineSummaryHLocation.Right;  // default
```

Getting this wrong is the usual reason a group's `+` button appears on the wrong row.

## Showing and hiding the outline controls

```csharp
ws.ShowOutlineSymbols = false;    // hide the +/- gutter but keep the grouping
ws.SetShowOutlineSymbols(true);
```

## Grouping vs hiding

Both make rows disappear, but they are not the same thing:

| | Grouping | Hiding |
|---|---|---|
| Reader can restore it | Yes, one click | Only via right-click → Unhide |
| Shows structure | Yes — levels and buttons | No |
| Survives filtering | Independent of autofilter | Autofilter also hides rows |

```csharp
ws.Rows(3, 8).Group(collapse: true);   // collapsible
ws.Rows(3, 8).Hide();                  // just gone
```

Use grouping for detail the reader may want; use hiding for scaffolding they should not see at
all (helper columns, lookup blocks).

## A worked example

A P&amp;L with collapsible sections, totals at the top of each group:

```csharp
using XLibur.Excel;

using var workbook = new XLWorkbook();
var ws = workbook.Worksheets.Add("P&L");

// Totals sit above their detail
ws.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;

ws.Cell("A1").Value = "Line item";
ws.Cell("B1").Value = "Amount";
ws.Range("A1:B1").Style.Font.Bold = true;

var sections = new (string Heading, (string Item, double Amount)[] Items)[]
{
    ("Revenue",
    [
        ("Product sales", 480_000),
        ("Services", 120_000),
        ("Licensing", 45_000),
    ]),
    ("Cost of sales",
    [
        ("Materials", 190_000),
        ("Direct labour", 145_000),
        ("Shipping", 22_000),
    ]),
    ("Operating expenses",
    [
        ("Salaries", 210_000),
        ("Premises", 48_000),
        ("Marketing", 36_000),
        ("Software", 19_000),
    ]),
};

var row = 2;
var sectionTotalRows = new List<int>();

foreach (var (heading, items) in sections)
{
    var headingRow = row;
    ws.Cell(row, 1).Value = heading;
    ws.Cell(row, 1).Style.Font.Bold = true;
    row++;

    var firstDetail = row;
    foreach (var (item, amount) in items)
    {
        ws.Cell(row, 1).Value = item;
        ws.Cell(row, 1).Style.Alignment.Indent = 2;
        ws.Cell(row, 2).Value = amount;
        row++;
    }

    var lastDetail = row - 1;

    // Section total on the heading row, above its detail
    ws.Cell(headingRow, 2).FormulaA1 = $"=SUM(B{firstDetail}:B{lastDetail})";
    ws.Cell(headingRow, 2).Style.Font.Bold = true;
    sectionTotalRows.Add(headingRow);

    // The detail rows become a collapsible group
    ws.Rows(firstDetail, lastDetail).Group();
}

// Grand total
ws.Cell(row + 1, 1).Value = "Net";
ws.Cell(row + 1, 1).Style.Font.Bold = true;
ws.Cell(row + 1, 2).FormulaA1 =
    $"=B{sectionTotalRows[0]}-B{sectionTotalRows[1]}-B{sectionTotalRows[2]}";
ws.Cell(row + 1, 2).Style.Font.Bold = true;
ws.Cell(row + 1, 2).Style.Border.TopBorder = XLBorderStyleValues.Thin;
ws.Cell(row + 1, 2).Style.Border.BottomBorder = XLBorderStyleValues.Double;

ws.Range($"B2:B{row + 1}").Style.NumberFormat.Format = "$ #,##0";
ws.Columns().AdjustToContents();

// Open summarised — the reader expands what they need
ws.CollapseRows();

workbook.SaveAs("ProfitAndLoss.xlsx");
```

## Where to next

- [Worksheets](./worksheets.md) — freeze panes and other view settings
- [Cells and Ranges](./cells-and-ranges.md) — hiding rows and columns outright
- [Page Setup](./page-setup.md) — printing a collapsed outline
