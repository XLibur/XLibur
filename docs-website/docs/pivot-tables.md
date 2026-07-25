---
id: pivot-tables
title: Pivot Tables
sidebar_label: Pivot Tables
description: Create pivot tables from ranges or tables, add row/column/filter fields, values with summary functions, grand totals, and layouts.
---

# Pivot Tables

A pivot table summarises a block of source data by grouping it along one or more fields and
aggregating a value field for each group. In XLibur you describe the pivot *specification* —
which fields go on rows, on columns, into the filter area, and what to aggregate — and Excel
renders the result when the file is opened.

A pivot table has four areas:

| Area | Property | Role |
|---|---|---|
| Rows | `RowLabels` | Fields that group data down the left |
| Columns | `ColumnLabels` | Fields that group data across the top |
| Values | `Values` | Fields that are aggregated in the body |
| Filters | `ReportFilters` | Fields exposed as page-level dropdowns |

:::note
XLibur writes the pivot table definition and a cache of the source data; it does not compute
the summarised cells itself. Excel (or LibreOffice) recalculates the body on open. That means
you will not see aggregated values by reading the pivot sheet's cells back with XLibur.
:::

## Inserting data

A pivot table needs a source. The cleanest one is an Excel table, because the pivot follows the
table as it grows:

```csharp
using XLibur.Excel;

var pastries = new List<Pastry>
{
    new("Croissant", 150, 60.2, "Apr"),
    new("Croissant", 250, 50.42, "May"),
    new("Doughnut", 250, 89.99, "Apr"),
    new("Doughnut", 225, 70, "May"),
    new("Danish", 394, 20.24, "Apr"),
    new("Danish", 190, 60, "May"),
};

using var workbook = new XLWorkbook();
var dataSheet = workbook.Worksheets.Add("PastrySalesData");

var source = dataSheet.Cell(1, 1).InsertTable(pastries, "PastrySalesData", createTable: true);
dataSheet.Columns().AdjustToContents();
```

A plain range works equally well when the data will not change size:

```csharp
var sourceRange = dataSheet.Range("A1:D7");
```

## Creating

Put the pivot on its own sheet and anchor it at a target cell. That cell becomes the pivot's
top-left corner:

```csharp
var pivotSheet = workbook.Worksheets.Add("Summary");

var pivot = pivotSheet.PivotTables.Add("SalesPivot", pivotSheet.Cell(1, 1), source);
```

`PivotTables.Add` has three overloads, one per source kind:

```csharp
pivotSheet.PivotTables.Add("P1", pivotSheet.Cell("A1"), table);        // from an IXLTable
pivotSheet.PivotTables.Add("P2", pivotSheet.Cell("A1"), range);        // from an IXLRange
pivotSheet.PivotTables.Add("P3", pivotSheet.Cell("A1"), pivotCache);   // reuse an existing cache
```

There is also a shortcut straight from the source:

```csharp
var pivot = source.CreatePivotTable(pivotSheet.FirstCell(), "SalesPivot");
```

:::tip
Two pivot tables built from the same range share one pivot cache automatically, which keeps the
file smaller. To share deliberately, pass `existingPivot.PivotCache` to the `Add` overload.
:::

Finding and removing pivot tables:

```csharp
var pivot = pivotSheet.PivotTables.PivotTable("SalesPivot");

if (pivotSheet.PivotTables.Contains("SalesPivot"))
{
    pivotSheet.PivotTables.Delete("SalesPivot");
}

pivotSheet.PivotTables.DeleteAll();
```

## Adding fields

Field names are the **source column names** — the table's header text, or the property names
of the objects you inserted.

```csharp
// Rows: one group per pastry name, then per month within it
pivot.RowLabels.Add("Name");
pivot.RowLabels.Add("Month");

// Columns: one column group per month
pivot.ColumnLabels.Add("Month");

// Filters: a page-level dropdown
pivot.ReportFilters.Add("Name");
```

Give a field a display name different from the source column with the two-argument overload:

```csharp
pivot.RowLabels.Add("Name", "Pastry");
```

### Field options

Each `Add` returns the field, so options chain off it:

```csharp
pivot.RowLabels.Add("Name")
    .SetSort(XLPivotSortType.Ascending)
    .SetCollapsed()
    .SetRepeatItemLabels()
    .SetInsertBlankLines(false);

pivot.RowLabels.Add("Month")
    .SetSort(XLPivotSortType.Descending)
    .SetShowBlankItems(false);
```

| Option | Effect |
|---|---|
| `SetSort(XLPivotSortType)` | `Default` (manual), `Ascending`, `Descending` |
| `SetCollapsed(bool)` | Start the group collapsed |
| `SetLayout(XLPivotLayout)` | `Compact`, `Outline`, or `Tabular` for this field |
| `SetRepeatItemLabels(bool)` | Repeat the group label on every row |
| `SetInsertBlankLines(bool)` | Blank line after each group |
| `SetShowBlankItems(bool)` | Include items with no data |
| `SetInsertPageBreaks(bool)` | Page break after each group when printing |
| `SetSubtotalsAtTop(bool)` | Subtotals above rather than below the group |
| `SetSubtotalCaption(string)` | Custom subtotal label |

### Filter fields with pre-selected values

Report filters can be pre-set to a subset of values:

```csharp
pivot.ReportFilters.Add("Name")
    .AddSelectedValue("Croissant")
    .AddSelectedValue("Doughnut");

pivot.ReportFilters.Add("Month")
    .AddSelectedValues(["Apr", "May"]);
```

## Values and totals columns

`Values.Add` puts a field in the body. The default aggregation is `Sum`:

```csharp
pivot.Values.Add("NumberOfOrders");
pivot.Values.Add("Quality");
```

Set the aggregation explicitly with `SetSummaryFormula`:

```csharp
pivot.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);
pivot.Values.Add("Quality").SetSummaryFormula(XLPivotSummary.Average);
pivot.Values.Add("Name").SetSummaryFormula(XLPivotSummary.Count);
```

Available summaries: `Sum`, `Count`, `Average`, `Minimum`, `Maximum`, `Product`,
`CountNumbers`, `StandardDeviation`, `PopulationStandardDeviation`, `Variance`,
`PopulationVariance`.

### The same field twice

Use the two-argument overload to add one source field under several names — a common way to
show both a total and a count of the same column:

```csharp
pivot.Values.Add("NumberOfOrders", "Total orders").SetSummaryFormula(XLPivotSummary.Sum);
pivot.Values.Add("NumberOfOrders", "Order count").SetSummaryFormula(XLPivotSummary.Count);
pivot.Values.Add("NumberOfOrders", "Average order").SetSummaryFormula(XLPivotSummary.Average);
```

### Number formats on values

```csharp
pivot.Values.Add("Quality", "Sum of Quality")
    .NumberFormat.SetFormat("#,##0.00");

pivot.Values.Add("Revenue", "Revenue")
    .NumberFormat.Format = "$ #,##0";
```

### "Show values as" calculations

Beyond raw aggregation, a value can be shown relative to something else:

```csharp
pivot.Values.Add("NumberOfOrders", "% of total").ShowAsPercentageOfTotal();
pivot.Values.Add("NumberOfOrders", "% of row").ShowAsPercentageOfRow();
pivot.Values.Add("NumberOfOrders", "% of column").ShowAsPercentageOfColumn();
pivot.Values.Add("NumberOfOrders", "Running total").ShowAsRunningTotalIn("Month");

// Relative to a specific item of another field
pivot.Values.Add("NumberOfOrders", "% of Danish")
    .ShowAsPercentageFrom("Name").And("Danish")
    .NumberFormat.Format = "0%";

// Difference from the previous item
pivot.Values.Add("NumberOfOrders", "Change")
    .ShowAsDifferenceFrom("Month").AndPrevious();

pivot.Values.Add("NumberOfOrders").ShowAsNormal();   // back to plain aggregation
```

### Where the value headers sit

With more than one value field, Excel adds a "Values" pseudo-field. Place it explicitly on rows
or columns using the sentinel label:

```csharp
// Value names down the rows
pivot.RowLabels.Add(XLConstants.PivotTable.ValuesSentinalLabel);
pivot.RowLabels.Add("Name");

// ...or across the columns
pivot.ColumnLabels.Add("Month");
pivot.ColumnLabels.Add(XLConstants.PivotTable.ValuesSentinalLabel);
```

## Grand totals and subtotals

```csharp
pivot.ShowGrandTotalsRows = true;      // total row at the bottom
pivot.ShowGrandTotalsColumns = true;   // total column on the right

// Fluent equivalents
pivot.SetShowGrandTotalsRows(false)
     .SetShowGrandTotalsColumns(true);
```

Subtotals for the intermediate groups are controlled at the table level:

```csharp
pivot.Subtotals = XLPivotSubtotals.DoNotShow;   // or AtTop, AtBottom
pivot.SetSubtotals(XLPivotSubtotals.AtBottom);
```

Per-field subtotal functions — a field can carry several at once:

```csharp
var field = pivot.RowLabels.Add("Name");
field.AddSubtotal(XLSubtotalFunction.Sum);
field.AddSubtotal(XLSubtotalFunction.Average);
field.SetSubtotal(XLSubtotalFunction.Count, enabled: false);
```

## Layout and appearance

```csharp
pivot.Layout = XLPivotLayout.Tabular;      // Compact (default), Outline, or Tabular
pivot.SetLayout(XLPivotLayout.Outline);

pivot.Theme = XLPivotTableTheme.PivotStyleMedium9;

pivot.ShowRowHeaders = true;
pivot.ShowColumnHeaders = true;
pivot.SetShowRowStripes()
     .SetShowColumnStripes(false);

pivot.SetRowHeaderCaption("Pastry name");
pivot.SetColumnHeaderCaption("Measures");

pivot.AutofitColumns = true;
pivot.PreserveCellFormatting = true;
```

Handling gaps and errors in the source:

```csharp
pivot.EmptyCellReplacement = "—";
pivot.ErrorValueReplacement = "n/a";
pivot.SetShowEmptyItemsOnRows(false);
pivot.SetShowEmptyItemsOnColumns(false);
```

Interaction switches, which matter mostly for the printed or exported view:

```csharp
pivot.SetShowExpandCollapseButtons(false);
pivot.SetDisplayCaptionsAndDropdowns(false);
pivot.SetClassicPivotTableLayout();
```

## A worked example

```csharp
using XLibur.Excel;

public record Pastry(string Name, int NumberOfOrders, double Quality, string Month);

using var workbook = new XLWorkbook();

// 1. The source data, as a table
var dataSheet = workbook.Worksheets.Add("PastrySalesData");
var pastries = new List<Pastry>
{
    new("Croissant", 150, 60.2, "Apr"),
    new("Croissant", 250, 50.42, "May"),
    new("Croissant", 134, 22.12, "Jun"),
    new("Doughnut", 250, 89.99, "Apr"),
    new("Doughnut", 225, 70, "May"),
    new("Doughnut", 210, 75.33, "Jun"),
    new("Danish", 394, 20.24, "Apr"),
    new("Danish", 190, 60, "May"),
    new("Danish", 221, 24.76, "Jun"),
};

var source = dataSheet.Cell(1, 1).InsertTable(pastries, "PastrySalesData", createTable: true);
dataSheet.Columns().AdjustToContents();

// 2. The pivot
var pivotSheet = workbook.Worksheets.Add("Summary");
var pivot = pivotSheet.PivotTables.Add("SalesPivot", pivotSheet.Cell(1, 1), source);

pivot.RowLabels.Add("Name").SetSort(XLPivotSortType.Ascending);
pivot.ColumnLabels.Add("Month");

pivot.Values.Add("NumberOfOrders", "Orders")
     .SetSummaryFormula(XLPivotSummary.Sum);

pivot.Values.Add("Quality", "Avg quality")
     .SetSummaryFormula(XLPivotSummary.Average)
     .NumberFormat.SetFormat("#,##0.00");

pivot.SetRowHeaderCaption("Pastry")
     .SetColumnHeaderCaption("Month")
     .SetShowGrandTotalsRows()
     .SetShowGrandTotalsColumns();

pivot.Subtotals = XLPivotSubtotals.DoNotShow;
pivot.Theme = XLPivotTableTheme.PivotStyleMedium9;

pivotSheet.Columns().AdjustToContents();
workbook.SaveAs("PastrySales.xlsx");
```

## Where to next

- [Tables](./tables.md) — the recommended pivot source
- [Theming](./theming.md) — pivot styles follow the workbook theme colours
