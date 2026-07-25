---
id: sparklines
title: Sparklines
sidebar_label: Sparklines
description: Add in-cell line, column, and win/loss sparklines, group them, and control markers, axis scaling, and colours.
---

# Sparklines

A sparkline is a tiny chart drawn inside a single cell — a trend line beside a row of monthly
figures, without the space a real chart needs. Excel manages them in **groups**: every
sparkline in a group shares one type, style, and axis configuration.

```csharp
using XLibur.Excel;

using var workbook = new XLWorkbook();
var ws = workbook.Worksheets.Add("Trends");

ws.Cell("A1").Value = "Region";
ws.Cell("B1").Value = "Trend";
for (var month = 1; month <= 12; month++)
{
    ws.Cell(1, month + 2).Value = new DateTime(2026, month, 1);
}

// ... write region names in A2:A5 and monthly figures in C2:N5 ...

ws.SparklineGroups.Add("B2:B5", "C2:N5");

workbook.SaveAs("Trends.xlsx");
```

That single `Add` creates one group holding four sparklines: `B2` plots `C2:N2`, `B3` plots
`C3:N3`, and so on — the location range and the source range are matched row by row.

## Creating groups

`SparklineGroups.Add` has four overloads:

```csharp
// Address strings
ws.SparklineGroups.Add("B2:B5", "C2:N5");

// Range objects — one sparkline per row
ws.SparklineGroups.Add(ws.Range("B2:B5"), ws.Range("C2:N5"));

// A single sparkline in one cell
ws.SparklineGroups.Add(ws.Cell("B2"), ws.Range("C2:N2"));

// An existing group (e.g. copied from another sheet)
ws.SparklineGroups.Add(otherGroup);
```

Adding more sparklines to an existing group, each with its own source range:

```csharp
var group = ws.SparklineGroups.Add("B2", "C2:N2");
group.Add(ws.Cell("B3"), ws.Range("C3:K3"));    // shorter history
group.Add(ws.Cell("B4"), ws.Range("C4:E4"));    // shorter still
```

Source ranges within a group do not have to be the same length — Excel scales each sparkline
to whatever data it has.

## Type

Three types, set on the group:

```csharp
group.Type = XLSparklineType.Line;      // default — a trend line
group.Type = XLSparklineType.Column;    // a small bar chart
group.Type = XLSparklineType.Stacked;   // win/loss: equal-height bars above/below the axis

group.SetType(XLSparklineType.Column);  // fluent form
```

`Stacked` is Excel's *Win/Loss* sparkline: it ignores magnitude and only shows sign, which
suits pass/fail or profit/loss series.

## Styles

`XLSparklineTheme` provides the built-in colour schemes Excel offers, as `IXLSparklineStyle`
values:

```csharp
group.SetStyle(XLSparklineTheme.Colorful1);
group.SetStyle(XLSparklineTheme.Dark3);
group.SetStyle(XLSparklineTheme.Accent2);
group.SetStyle(XLSparklineTheme.Default);
```

Families: `Dark1`–`Dark6`, `Colorful1`–`Colorful6`, and `Accent1`–`Accent6` (plus the tinted
variants Excel shows in its gallery). `Default` is `Dark5`.

Colours can also be set individually:

```csharp
group.Style
    .SetSeriesColor(XLColor.FromTheme(XLThemeColor.Accent1))
    .SetNegativeColor(XLColor.Red)
    .SetHighMarkerColor(XLColor.Green)
    .SetLowMarkerColor(XLColor.Red)
    .SetFirstMarkerColor(XLColor.Gray)
    .SetLastMarkerColor(XLColor.Black)
    .SetMarkersColor(XLColor.DarkGray);
```

## Markers

`XLSparklineMarkers` is a flags enum — combine the points you want highlighted:

```csharp
group.SetShowMarkers(XLSparklineMarkers.All);

group.SetShowMarkers(
    XLSparklineMarkers.FirstPoint |
    XLSparklineMarkers.LastPoint |
    XLSparklineMarkers.HighPoint |
    XLSparklineMarkers.LowPoint);

group.ShowMarkers = XLSparklineMarkers.None;
```

| Flag | Highlights |
|---|---|
| `HighPoint` / `LowPoint` | The maximum and minimum |
| `FirstPoint` / `LastPoint` | The ends of the series |
| `NegativePoints` | Every value below zero |
| `Markers` | Every data point (line sparklines only) |
| `All` | All of the above |

## Axis scaling

By default each sparkline scales to its own data, so a row with small numbers looks just as
dramatic as one with large numbers. `SameForAll` makes the group share one scale, which is
what you usually want when comparing rows:

```csharp
group.VerticalAxis
    .SetMinAxisType(XLSparklineAxisMinMax.SameForAll)
    .SetMaxAxisType(XLSparklineAxisMinMax.SameForAll);
```

Fixed bounds:

```csharp
group.VerticalAxis
    .SetMinAxisType(XLSparklineAxisMinMax.Custom)
    .SetMaxAxisType(XLSparklineAxisMinMax.Custom)
    .SetManualMin(-80)
    .SetManualMax(100);
```

| Axis type | Behaviour |
|---|---|
| `Automatic` | Each sparkline scales to its own data (default) |
| `SameForAll` | One scale across the whole group |
| `Custom` | Use `ManualMin` / `ManualMax` |

The horizontal axis can be shown as a zero line:

```csharp
group.HorizontalAxis
    .SetVisible(true)
    .SetColor(XLColor.Red)
    .SetRightToLeft(false);
```

## Date axis

Passing a date range makes the horizontal spacing proportional to time rather than to point
count — so a gap in the data reads as a gap:

```csharp
group.SetDateRange(ws.Range("C1:N1"));

Console.WriteLine(group.HorizontalAxis.DateAxis);   // true
group.SetDateRange(null);                            // back to even spacing
```

## Blanks and hidden data

```csharp
group.DisplayEmptyCellsAs = XLDisplayBlanksAsValues.Interpolate;   // bridge the gap
group.DisplayEmptyCellsAs = XLDisplayBlanksAsValues.Zero;          // plot as zero
group.DisplayEmptyCellsAs = XLDisplayBlanksAsValues.NotPlotted;    // leave a gap

group.DisplayHidden = true;   // include rows/columns the user has hidden
```

## Line weight

Line sparklines only:

```csharp
group.SetLineWeight(2);
group.LineWeight = 0.75;
```

## Finding and removing

```csharp
foreach (var group in ws.SparklineGroups)
{
    Console.WriteLine($"{group.Type}, {group.Count()} sparklines");

    foreach (var sparkline in group)
    {
        Console.WriteLine($"  {sparkline.Location.Address} <- {sparkline.SourceData.RangeAddress}");
    }
}

group.Remove(sparkline);
ws.Range("B2:B5").Clear(XLClearOptions.Sparklines);
```

A sparkline can be re-pointed after creation:

```csharp
sparkline.SetLocation(ws.Cell("B7"))
         .SetSourceData(ws.Range("C7:N7"));
```

## Sparklines vs charts vs data bars

Three ways to put a visual in a sheet, and they solve different problems:

| Use | When |
|---|---|
| **Sparkline** | A trend across many points, one row at a time, inline with the data |
| **[Data bar](./conditional-formatting.md#data-bars)** | Comparing a *single* value per row against the others |
| **[Chart](./charts.md)** | The visual is the point, and it needs a title, axes, and a legend |

## A worked example

```csharp
using XLibur.Excel;

using var workbook = new XLWorkbook();
var ws = workbook.Worksheets.Add("Monthly");

// Header: region, sparkline column, then twelve months
ws.Cell("A1").Value = "Region";
ws.Cell("B1").Value = "Trend";
for (var month = 1; month <= 12; month++)
{
    ws.Cell(1, month + 2).Value = new DateTime(2026, month, 1);
}

ws.Range(1, 3, 1, 14).Style.DateFormat.Format = "MMM";
ws.Range(1, 1, 1, 14).Style.Font.Bold = true;

var data = new[]
{
    ("North", new[] { 31d, 35, 29, 40, 44, 39, 42, 48, 51, 47, 53, 58 }),
    ("South", new[] { 22d, 19, 25, 21, 18, 24, 20, 17, 23, 19, 16, 21 }),
    ("East",  new[] { 12d, 18, 24, 33, 41, 52, 60, 71, 78, 84, 92, 99 }),
    ("West",  new[] { 45d, 42, 38, 40, 35, 33, 36, 31, 29, 32, 27, 25 }),
};

var row = 2;
foreach (var (region, values) in data)
{
    ws.Cell(row, 1).Value = region;
    for (var i = 0; i < values.Length; i++)
    {
        ws.Cell(row, i + 3).Value = values[i];
    }

    row++;
}

var last = row - 1;

// One group, shared scale so the four regions are directly comparable
var group = ws.SparklineGroups.Add($"B2:B{last}", $"C2:N{last}");

group.SetType(XLSparklineType.Line)
     .SetStyle(XLSparklineTheme.Colorful1)
     .SetShowMarkers(XLSparklineMarkers.HighPoint | XLSparklineMarkers.LowPoint | XLSparklineMarkers.LastPoint)
     .SetDateRange(ws.Range("C1:N1"))
     .SetLineWeight(1.25);

group.VerticalAxis
     .SetMinAxisType(XLSparklineAxisMinMax.SameForAll)
     .SetMaxAxisType(XLSparklineAxisMinMax.SameForAll);

ws.Column("B").Width = 18;
ws.Columns(3, 14).Width = 5;
ws.Rows(2, last).Height = 22;

workbook.SaveAs("MonthlyTrends.xlsx");
```

## Where to next

- [Charts](./charts.md) — full charts when a sparkline is not enough
- [Conditional Formatting](./conditional-formatting.md) — data bars and colour scales
