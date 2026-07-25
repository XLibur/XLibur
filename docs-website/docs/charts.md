---
id: charts
title: Charts
sidebar_label: Charts
description: Add column, line, pie, scatter, stock, surface, combo, and modern Excel charts to a worksheet, with series, titles, and anchor positioning.
---

# Charts

XLibur can embed charts in a worksheet. You choose a chart type, add one or more series that
point at ranges of data, and anchor the chart to a rectangle of cells.

:::note
Chart support here is substantially wider than in ClosedXML 0.105, which could not create
charts at all. Over 70 chart types are available, including combo charts and the Office 2016+
"extended" types (Waterfall, Funnel, Treemap, Sunburst, Box &amp; Whisker).
:::

## A first chart

```csharp
using XLibur.Excel;

using var workbook = new XLWorkbook();
var ws = workbook.Worksheets.Add("Sales");

ws.Cell("A1").Value = "Quarter";
ws.Cell("B1").Value = "Revenue";

string[] quarters = ["Q1", "Q2", "Q3", "Q4"];
double[] revenue = [30_000, 45_000, 28_000, 50_000];

for (var i = 0; i < quarters.Length; i++)
{
    ws.Cell(i + 2, 1).Value = quarters[i];
    ws.Cell(i + 2, 2).Value = revenue[i];
}

var chart = ws.Charts.Add(XLChartType.ColumnClustered);
chart.SetTitle("Revenue by quarter");
chart.Series.Add("Revenue", "Sales!$B$2:$B$5", "Sales!$A$2:$A$5");

chart.Position.SetColumn(3).SetRow(1);
chart.SecondPosition.SetColumn(11).SetRow(16);

workbook.SaveAs("Revenue.xlsx");
```

Three things are going on: `Charts.Add` creates the chart, `Series.Add` tells it what to plot,
and the two positions define the rectangle it occupies.

## Series references

`Series.Add(name, valueReferences, categoryReferences)` takes **string references**, not range
objects, and they must be fully qualified with the sheet name and absolute (`$`) addressing —
this is the form Excel stores in the chart part:

```csharp
chart.Series.Add("Revenue", "Sales!$B$2:$B$5", "Sales!$A$2:$A$5");
```

| Argument | Meaning |
|---|---|
| `name` | Series label, shown in the legend |
| `valueReferences` | The numbers to plot |
| `categoryReferences` | Optional axis labels; pass `null` to omit |

Sheet names containing spaces must be single-quoted, exactly as in a formula:

```csharp
chart.Series.Add("Revenue", "'Q1 Sales'!$B$2:$B$5", "'Q1 Sales'!$A$2:$A$5");
```

Building the references from live range objects avoids hand-typed typos:

```csharp
var values = ws.Range("B2:B5");
var categories = ws.Range("A2:A5");

static string Ref(IXLRange range)
{
    var sheet = range.Worksheet.Name;
    var quoted = sheet.Contains(' ') ? $"'{sheet}'" : sheet;
    return $"{quoted}!{range.RangeAddress.ToStringFixed()}";
}

chart.Series.Add("Revenue", Ref(values), Ref(categories));
```

### Several series

Add one per data column. All of them normally share the same category reference:

```csharp
var categories = "Sales!$A$2:$A$5";

var chart = ws.Charts.Add(XLChartType.ColumnClustered);
chart.SetTitle("Revenue vs cost");
chart.Series.Add("Revenue", "Sales!$B$2:$B$5", categories);
chart.Series.Add("Cost", "Sales!$C$2:$C$5", categories);
chart.Series.Add("Margin", "Sales!$D$2:$D$5", categories);

Console.WriteLine(chart.Series.Count);   // 3
```

Each series exposes its plot order and can be re-pointed after the fact:

```csharp
foreach (var series in chart.Series)
{
    Console.WriteLine($"{series.Order}: {series.Name} -> {series.ValueReferences}");
}

var first = chart.Series.First();
first.Name = "Net revenue";
first.ValueReferences = "Sales!$B$2:$B$9";
```

## Positioning

A chart is anchored to two cells: `Position` is the top-left corner and `SecondPosition` the
bottom-right. The rectangle between them is the chart's size and location.

```csharp
chart.Position.SetColumn(3).SetRow(1);        // top-left
chart.SecondPosition.SetColumn(11).SetRow(16); // bottom-right
```

:::warning
Chart anchor row and column indexes are **0-based**, unlike the 1-based cell addressing used
everywhere else in XLibur. `SetColumn(0).SetRow(0)` is the top-left of the sheet (cell `A1`);
`SetColumn(3)` is column `D`.
:::

Fine-tune with fractional offsets inside the anchor cell:

```csharp
chart.Position.SetColumn(3).SetColumnOffset(0.5).SetRow(1).SetRowOffset(0.25);
```

Other drawing-level properties:

```csharp
chart.Visible = true;
chart.ZOrder = 2;              // stacking order against other drawings
Console.WriteLine(chart.ShapeId);
```

A small helper keeps multi-chart sheets readable:

```csharp
static void PlaceChart(IXLChart chart, int row, int column, int width = 8, int height = 15)
{
    chart.Position.SetColumn(column).SetRow(row);
    chart.SecondPosition.SetColumn(column + width).SetRow(row + height);
}

PlaceChart(revenueChart, row: 1, column: 4);
PlaceChart(marginChart, row: 17, column: 4);
```

## Combo charts

Set `SecondaryChartType` and add series to `SecondarySeries`. Both types share one plot area —
the classic "bars with a line over them":

```csharp
var combo = ws.Charts.Add(XLChartType.ColumnClustered);
combo.SetTitle("Units sold and average price");

combo.Series.Add("Units", "Sales!$B$2:$B$5", "Sales!$A$2:$A$5");

combo.SecondaryChartType = XLChartType.Line;
combo.SecondarySeries.Add("Avg price", "Sales!$C$2:$C$5", "Sales!$A$2:$A$5");

combo.Position.SetColumn(0).SetRow(7);
combo.SecondPosition.SetColumn(10).SetRow(24);
```

Set `SecondaryChartType` back to `null` to turn a combo chart into a single-type chart.

## Chart types

`XLChartType` covers the full Excel catalogue. Pick by data shape:

| Family | Types |
|---|---|
| **Column** | `ColumnClustered`, `ColumnStacked`, `ColumnStacked100Percent`, `Column3D`, `ColumnClustered3D`, `ColumnStacked3D`, `ColumnStacked100Percent3D` |
| **Bar** (horizontal) | `BarClustered`, `BarStacked`, `BarStacked100Percent`, `BarClustered3D`, `BarStacked3D`, `BarStacked100Percent3D` |
| **Line** | `Line`, `LineStacked`, `LineStacked100Percent`, `LineWithMarkers`, `LineWithMarkersStacked`, `LineWithMarkersStacked100Percent`, `Line3D` |
| **Area** | `Area`, `AreaStacked`, `AreaStacked100Percent`, `Area3D`, `AreaStacked3D`, `AreaStacked100Percent3D` |
| **Pie / Doughnut** | `Pie`, `PieExploded`, `Pie3D`, `PieExploded3D`, `PieToPie`, `PieToBar`, `Doughnut`, `DoughnutExploded` |
| **Scatter (XY)** | `XYScatterMarkers`, `XYScatterStraightLinesWithMarkers`, `XYScatterStraightLinesNoMarkers`, `XYScatterSmoothLinesWithMarkers`, `XYScatterSmoothLinesNoMarkers` |
| **Bubble** | `Bubble`, `Bubble3D` |
| **Radar** | `Radar`, `RadarWithMarkers`, `RadarFilled` |
| **Stock** | `StockHighLowClose`, `StockOpenHighLowClose`, `StockVolumeHighLowClose`, `StockVolumeOpenHighLowClose` |
| **Surface** | `Surface`, `SurfaceWireframe`, `SurfaceContour`, `SurfaceContourWireframe` |
| **Cone** | `Cone`, `ConeClustered`, `ConeStacked`, `ConeStacked100Percent`, `ConeHorizontalClustered`, `ConeHorizontalStacked`, `ConeHorizontalStacked100Percent` |
| **Cylinder** | `Cylinder`, `CylinderClustered`, `CylinderStacked`, `CylinderStacked100Percent`, `CylinderHorizontalClustered`, `CylinderHorizontalStacked`, `CylinderHorizontalStacked100Percent` |
| **Pyramid** | `Pyramid`, `PyramidClustered`, `PyramidStacked`, `PyramidStacked100Percent`, `PyramidHorizontalClustered`, `PyramidHorizontalStacked`, `PyramidHorizontalStacked100Percent` |
| **Extended** (Office 2016+) | `Waterfall`, `Funnel`, `Treemap`, `Sunburst`, `BoxWhisker` |

### Data shape by family

Most families take one category column plus one value column per series. A few are different:

**Stock charts** expect the price columns in a fixed order — high/low/close, optionally
preceded by open, optionally preceded by volume:

```csharp
var stock = ws.Charts.Add(XLChartType.StockHighLowClose);
stock.Series.Add("High", "Prices!$B$2:$B$11", "Prices!$A$2:$A$11");
stock.Series.Add("Low", "Prices!$C$2:$C$11", "Prices!$A$2:$A$11");
stock.Series.Add("Close", "Prices!$D$2:$D$11", "Prices!$A$2:$A$11");
```

**Hierarchical charts** (`Treemap`, `Sunburst`) take a multi-column category reference, one
column per level of the hierarchy:

```csharp
// Columns A, B, C hold Branch / Category / Item; column D holds the value
var sunburst = ws.Charts.Add(XLChartType.Sunburst);
sunburst.SetTitle("Spend breakdown");
sunburst.Series.Add("Value", "'Spend'!$D$2:$D$8", "'Spend'!$A$2:$C$8");
```

**Box &amp; Whisker** takes the raw observations with a grouping column as the category — Excel
computes the quartiles:

```csharp
var box = ws.Charts.Add(XLChartType.BoxWhisker);
box.Series.Add("Value", "Samples!$B$2:$B$9", "Samples!$A$2:$A$9");
```

**Waterfall** takes a single series of signed amounts:

```csharp
var waterfall = ws.Charts.Add(XLChartType.Waterfall);
waterfall.Series.Add("Amount", "Bridge!$B$2:$B$6", "Bridge!$A$2:$A$6");
```

:::note
The extended types are written to an `ExtendedChartPart` using the Office 2016 `cx` namespace.
Excel 2016 and later render them; older versions and some third-party readers show a
placeholder instead.
:::

## 3D charts

`RightAngleAxes` keeps the axes square regardless of the chart's rotation, which usually reads
better than the default perspective:

```csharp
var chart = ws.Charts.Add(XLChartType.Column3D);
chart.SetRightAngleAxes();
chart.SetRightAngleAxes(false);   // back to perspective
```

## Finding charts

```csharp
Console.WriteLine(ws.Charts.Count);

foreach (var chart in ws.Charts)
{
    Console.WriteLine($"{chart.Title} ({chart.ChartType}), {chart.Series.Count} series");
}
```

Chart type and title are settable after creation:

```csharp
var chart = ws.Charts.First();
chart.ChartType = XLChartType.BarClustered;
chart.Title = "Revised";
chart.Title = null;               // remove the title
```

## What is not covered

The chart API is deliberately narrow: type, title, series, and placement. Axis titles, legend
placement, gridlines, data labels, per-series colours, and trend lines are not exposed. If you
need that level of control, the usual approach is to build a template workbook in Excel with
the chart formatted exactly as you want it, then use XLibur to write the source data the chart
already points at:

```csharp
using var workbook = new XLWorkbook("ChartTemplate.xlsx");
var ws = workbook.Worksheet("Data");

var row = 2;
foreach (var (label, value) in results)
{
    ws.Cell(row, 1).Value = label;
    ws.Cell(row, 2).Value = value;
    row++;
}

workbook.Save();   // the pre-formatted chart picks up the new data
```

## A worked example

```csharp
using XLibur.Excel;

using var workbook = new XLWorkbook();
var ws = workbook.Worksheets.Add("Quarterly");

// Data
ws.Cell("A1").Value = "Quarter";
ws.Cell("B1").Value = "Revenue";
ws.Cell("C1").Value = "Cost";
ws.Cell("D1").Value = "Margin %";
ws.Range("A1:D1").Style.Font.Bold = true;

var data = new[]
{
    ("Q1", 30_000d, 21_000d),
    ("Q2", 45_000d, 29_000d),
    ("Q3", 28_000d, 20_500d),
    ("Q4", 50_000d, 31_000d),
};

var row = 2;
foreach (var (quarter, revenue, cost) in data)
{
    ws.Cell(row, 1).Value = quarter;
    ws.Cell(row, 2).Value = revenue;
    ws.Cell(row, 3).Value = cost;
    ws.Cell(row, 4).FormulaA1 = $"=(B{row}-C{row})/B{row}";
    row++;
}

var last = row - 1;
ws.Range($"B2:C{last}").Style.NumberFormat.Format = "$ #,##0";
ws.Range($"D2:D{last}").Style.NumberFormat.Format = "0.0%";
ws.Columns().AdjustToContents();

const string sheet = "Quarterly";
var categories = $"{sheet}!$A$2:$A${last}";

// Chart 1: revenue vs cost
var bars = ws.Charts.Add(XLChartType.ColumnClustered);
bars.SetTitle("Revenue vs cost");
bars.Series.Add("Revenue", $"{sheet}!$B$2:$B${last}", categories);
bars.Series.Add("Cost", $"{sheet}!$C$2:$C${last}", categories);
bars.Position.SetColumn(5).SetRow(1);
bars.SecondPosition.SetColumn(13).SetRow(16);

// Chart 2: combo — revenue bars with the margin line over them
var combo = ws.Charts.Add(XLChartType.ColumnClustered);
combo.SetTitle("Revenue and margin");
combo.Series.Add("Revenue", $"{sheet}!$B$2:$B${last}", categories);
combo.SecondaryChartType = XLChartType.LineWithMarkers;
combo.SecondarySeries.Add("Margin %", $"{sheet}!$D$2:$D${last}", categories);
combo.Position.SetColumn(5).SetRow(18);
combo.SecondPosition.SetColumn(13).SetRow(33);

// Chart 3: share of annual revenue
var pie = ws.Charts.Add(XLChartType.Pie);
pie.SetTitle("Share of annual revenue");
pie.Series.Add("Revenue", $"{sheet}!$B$2:$B${last}", categories);
pie.Position.SetColumn(14).SetRow(1);
pie.SecondPosition.SetColumn(21).SetRow(16);

workbook.SaveAs("QuarterlyCharts.xlsx");
```

## Where to next

- [Sparklines](./sparklines.md) — in-cell mini charts, when a full chart is too much
- [Images and Pictures](./images.md) — the other kind of drawing, and shared anchor rules
- [Pivot Tables](./pivot-tables.md) — summarising the data a chart plots
