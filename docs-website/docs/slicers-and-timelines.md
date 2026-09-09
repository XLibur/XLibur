---
id: slicers-and-timelines
title: Slicers and Timelines
sidebar_label: Slicers and Timelines
description: Read, create, style and position the slicer button panels and timeline date scrubbers that filter pivot tables and tables.
---

# Slicers and Timelines

A **slicer** is the panel of buttons Excel draws beside a pivot table or a table so a reader can
filter it without opening a dropdown. A **timeline** is the same idea for dates: a scrubber band
with a Years / Quarters / Months / Days level chooser.

Both are drawings on a worksheet, and both filter through a *cache* rather than through the thing
they sit next to — which is how one slicer can drive several pivot tables at once on a dashboard.

| | Filters | Added with |
|---|---|---|
| Pivot slicer | One or more pivot tables sharing a pivot cache | `Slicers.Add(pivotTable, fieldName)` |
| Table slicer | One column of a table, through that table's autofilter | `Slicers.Add(table, columnName)` |
| Timeline | One or more pivot tables, on a date field | `Timelines.Add(pivotTable, dateFieldName)` |

:::note Selection is read-only
You can read which items a slicer has selected and which range a timeline covers, but you cannot
set them. Writing a selection means writing the pivot table's item visibility — and, for a
timeline, its `dateBetween` filter — in the same breath, and a model that moved one without the
others would produce a workbook that disagrees with itself in a way no validator can see. Create
the control, and let the reader click it.
:::

## Reading what a file already has

Slicers and timelines are owned by the worksheet they are drawn on:

```csharp
using var workbook = new XLWorkbook("Dashboard.xlsx");
var sheet = workbook.Worksheet("Pivot");

foreach (var slicer in sheet.Slicers)
{
    Console.WriteLine($"{slicer.Name}: {slicer.Caption} on {slicer.SourceFieldName}");
}

var byName = sheet.Slicers.Slicer("Region");             // throws if there is none
if (sheet.Slicers.TryGetSlicer("Region", out var region))
{
    // ...
}
```

What each control *filters* is a separate relationship, so a pivot table can list the controls
pointing at it even when they live on another sheet:

```csharp
var pivot = workbook.Worksheet("Pivot").PivotTables.PivotTable("SalesPivot");

foreach (var slicer in pivot.Slicers)      // IEnumerable<IXLSlicer>
foreach (var timeline in pivot.Timelines)  // IEnumerable<IXLTimeline>
```

Reading a slicer's current state:

```csharp
var slicer = sheet.Slicers.Slicer("Region");

slicer.SourceKind;        // XLSlicerSourceKind.PivotTable or .Table
slicer.SourceFieldName;   // the pivot cache field, or the table column
slicer.PivotTables;       // IReadOnlyList<IXLPivotTable> — empty for a table slicer
slicer.Table;             // IXLTable? — null for a pivot slicer

if (slicer.HasSelection)
{
    foreach (XLCellValue item in slicer.SelectedItems)
        Console.WriteLine(item);
}
```

`HasSelection` is `false` when the slicer records no explicit selection, which is how Excel
represents a slicer nobody has clicked — every item is showing. A table column filtered by
something a slicer cannot produce (a custom or top-ten filter applied by hand) reports no
selection.

A timeline reports a range rather than a list:

```csharp
var timeline = sheet.Timelines.Timeline("Date");

timeline.Level;            // XLTimelineLevel.Years / Quarters / Months / Days
timeline.BoundsStart;      // DateTime? — the extent of the band
timeline.BoundsEnd;
timeline.HasSelection;
timeline.SelectionStart;   // DateTime? — null when HasSelection is false
timeline.SelectionEnd;
```

`BoundsStart` and `BoundsEnd` are read-only because Excel recomputes the extent whenever the pivot
cache refreshes; a settable bound would be honest in only one direction.

## Adding a slicer

```csharp
var pivotSheet = workbook.Worksheet("Pivot");
var pivot = pivotSheet.PivotTables.PivotTable("SalesPivot");

// Filters the pivot table on one of its cache fields
var region = pivotSheet.Slicers.Add(pivot, "Region");

// Filters a table on one of its columns
var dataSheet = workbook.Worksheet("Data");
var amount = dataSheet.Slicers.Add(dataSheet.Tables.Table("Sales"), "Amount");
```

Both overloads throw `ArgumentException` when the field or column does not exist. The new slicer
starts with every item selected, so it filters nothing until someone clicks it.

The slicer is placed to the right of whatever it filters. Move it by assigning a cell:

```csharp
region.Position = pivotSheet.Cell("E3");
```

A slicer is anchored to the grid like any other drawing, so it travels when rows or columns are
inserted above or to the left of it. Moving one read from a file shifts both of its corners
together, so it keeps the size it had.

:::tip A control need not sit on the sheet it filters
`Slicers` and `Timelines` belong to the worksheet the control is *drawn* on; the pivot table or
table passed to `Add` may live anywhere in the workbook. That is how a dashboard sheet drives
pivot tables kept out of sight.
:::

## Adding a timeline

```csharp
var timeline = pivotSheet.Timelines.Add(pivot, "Date");
timeline.Caption = "Pick a period";
timeline.Style = "TimeSlicerStyleLight2";
timeline.Level = XLTimelineLevel.Months;
timeline.Position = pivotSheet.Cell("E3");
```

The field must hold dates — `Add` throws `ArgumentException` when the pivot cache has no field of
that name, or when the field it names holds something else.

## Styling

Everything below is settable on a control you created *and* on one loaded from a file:

```csharp
slicer.Caption = "Pick a region";   // the heading; defaults to Name
slicer.ShowCaption = true;
slicer.Style = "SlicerStyleDark3";  // null means the workbook default
slicer.ColumnCount = 2;             // columns of buttons
slicer.RowHeightPt = 19.5;          // one button row, in points; null leaves it to Excel

timeline.Caption = "Pick a period";
timeline.ShowHeader = true;
timeline.ShowSelectionLabel = true;        // the range written under the header
timeline.ShowTimeLevel = true;             // the Years/Quarters/Months/Days chooser
timeline.ShowHorizontalScrollbar = true;
timeline.Style = "TimeSlicerStyleLight2";
```

`Style` is a plain string rather than an enumeration of the built-in styles, deliberately: a
workbook may name a custom style, and a model that could only report the styles it knows about
would silently lose the rest.

`Name` is read-only. It is the internal identifier the drawing anchor refers to and must be unique
across the workbook — `Caption` is what the reader sees.

:::note Edits are patched, not regenerated
A control loaded from a file is edited by patching the change into the part it was read from, so
everything alongside the edited attribute survives — the `xr10:uid`, the `startItem`, the
extension list. A control nobody assigns to is not written to at all: its part is not opened, and
comes through a save byte for byte.
:::

## Removing

There is no public `Remove`. Deleting the pivot table takes its slicers and timelines with it,
along with their caches, registrations and drawing anchors, so nothing is left for Excel to offer
to repair:

```csharp
pivotSheet.PivotTables.Delete("SalesPivot");
```

## A worked example

A pivot table with both control types on the same sheet:

```csharp
using XLibur.Excel;

using var workbook = new XLWorkbook();

// 1. Source data with a date column
var data = workbook.AddWorksheet("Data");
data.Cell("A1").Value = "Date";
data.Cell("B1").Value = "Region";
data.Cell("C1").Value = "Amount";

var start = new DateTime(2024, 1, 15);
for (var i = 0; i < 24; i++)
{
    data.Cell(i + 2, 1).Value = start.AddDays(i * 11);
    data.Cell(i + 2, 2).Value = i % 2 == 0 ? "North" : "South";
    data.Cell(i + 2, 3).Value = 100 + (i * 7);
}

data.Column(1).Style.DateFormat.Format = "yyyy-mm-dd";

// 2. The pivot table
var pivotSheet = workbook.AddWorksheet("Pivot");
var pivot = pivotSheet.PivotTables.Add("SalesPivot", pivotSheet.Cell("A3"), data.Range("A1:C25"));
pivot.RowLabels.Add("Region");
pivot.Values.Add("Amount");

// 3. A timeline on the date field, and a slicer on the region field
var timeline = pivotSheet.Timelines.Add(pivot, "Date");
timeline.Caption = "Pick a period";
timeline.Style = "TimeSlicerStyleLight2";
timeline.Position = pivotSheet.Cell("E3");

var slicer = pivotSheet.Slicers.Add(pivot, "Region");
slicer.Caption = "Region";
slicer.ColumnCount = 2;
slicer.Position = pivotSheet.Cell("E20");

workbook.SaveAs("Dashboard.xlsx");
```

## What is not modelled

| | Why |
|---|---|
| Setting a slicer's selected items | Has to move the pivot table's item visibility with it |
| Setting a timeline's selected range | Excel records it in three places at once — cache state, a `dateBetween` pivot filter, and hidden-item flags |
| Setting a timeline's bounds | Excel recomputes them on every cache refresh |
| Removing a control on its own | Delete the pivot table, which cascades |
| An enumeration of built-in styles | A workbook may name a custom style; `Style` carries whatever the file says |

Anything else a file carries and XLibur has no model for — a custom style name, an offset into the
anchor cell, an extension list — is preserved through a save rather than dropped.

## Where to next

- [Pivot Tables](./pivot-tables.md) — building the pivot a slicer or timeline filters
- [Tables](./tables.md) — the source for a table slicer
- [AutoFilter](./autofilter.md) — where a table slicer's selection is actually stored
