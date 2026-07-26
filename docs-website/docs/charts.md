---
id: charts
title: Charts
sidebar_label: Charts
description: Add column, line, pie, scatter, stock, surface, combo, and modern Excel charts to a worksheet, with series, titles, and anchor positioning.
---

# Charts

XLibur can embed charts in a worksheet. You choose a chart type, add one or more series that
point at ranges of data, and anchor the chart to a rectangle of cells.

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

### Anchoring

By default a chart spans the rectangle between `Position` and `SecondPosition`, so inserting rows or
columns moves and resizes it. `Anchor` picks a different rule:

```csharp
// Keep the size, move with the cell underneath the top-left corner
chart.Anchor = XLDrawingAnchor.MoveWithCells;
chart.Position.SetColumn(4).SetRow(3);
chart.Width = 480;    // pixels
chart.Height = 288;

// Pin to a spot on the sheet, ignoring the grid
chart.Anchor = XLDrawingAnchor.Absolute;
chart.Left = 200;     // pixels from the left edge
chart.Top = 120;
chart.Width = 480;
chart.Height = 288;
```

`SecondPosition` is only used by the default `MoveAndSizeWithCells`; `Width`/`Height` are only used
by the other two, and `Left`/`Top` only by `Absolute`. All four are read back from a file, whichever
anchor the chart came with.

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

## Series formatting

Each series carries its own fill, outline and marker. Every property is optional: leave it alone
and Excel picks the automatic colour from the workbook theme, which is usually what you want for
all but one or two highlighted series.

```csharp
var chart = ws.Charts.Add(XLChartType.ColumnClustered);
var revenue = chart.Series.Add("Revenue", "Sales!$B$2:$B$5", "Sales!$A$2:$A$5");

revenue.FillColor = XLColor.FromHtml("#4472C4");     // bar interior
revenue.LineColor = XLColor.FromHtml("#203864");     // bar border
revenue.LineWidthPt = 1.5;
```

On a line or scatter series `LineColor` and `LineWidthPt` draw the line itself. The marker is a
separate thing: `MarkerStyle` picks the symbol and `MarkerFillColor` fills it.

```csharp
chart.SecondaryChartType = XLChartType.LineWithMarkers;
var trend = chart.SecondarySeries.Add("Margin %", "Sales!$D$2:$D$5", "Sales!$A$2:$A$5");

trend.LineColor = XLColor.FromTheme(XLThemeColor.Accent2);
trend.LineWidthPt = 2.25;
trend.MarkerStyle = XLMarkerStyle.Circle;
trend.MarkerSize = 7;                                 // 2 to 72 points
trend.MarkerFillColor = XLColor.FromHtml("#70AD47");
trend.Smooth = true;                                  // curve rather than straight segments
```

| Property | Applies to | Default |
|---|---|---|
| `FillColor` | bar, column, area and pie interiors | automatic (theme) |
| `LineColor` | line and scatter lines; borders elsewhere | automatic (theme) |
| `LineWidthPt` | any outline, 0 to 1584 points | Excel's own |
| `MarkerStyle` | line, scatter, radar | `Auto` |
| `MarkerSize` | line, scatter, radar, 2 to 72 points | Excel's own |
| `MarkerFillColor` | line, scatter, radar | automatic (theme) |
| `Smooth` | line, scatter | the chart type's default |

`XLMarkerStyle` offers `Auto`, `None`, `Circle`, `Dash`, `Diamond`, `Dot`, `Plus`, `Square`,
`Star`, `Triangle` and `X`. `Auto` — the default — writes nothing, so the chart type decides:
the `LineWithMarkers*` types draw markers, plain `Line` does not.

:::note
Setting a colour to `null` clears it back to automatic; it never writes an explicit black. A
theme colour is written as a DrawingML scheme colour, but its `ThemeTint` is not applied — pass an
RGB colour if you need a specific tint.

The extended (Office 2016+) types — Waterfall, Funnel, Treemap, Sunburst, Box &amp; Whisker —
ignore series formatting.
:::

### Secondary value axis

`UseSecondaryAxis` moves a series onto a second value axis drawn on the right, which is what
makes a two-scale chart readable — units in the thousands next to a percentage, say:

```csharp
var chart = ws.Charts.Add(XLChartType.ColumnClustered);
chart.Series.Add("Units", "Sales!$B$2:$B$5", "Sales!$A$2:$A$5");

chart.SecondaryChartType = XLChartType.LineWithMarkers;
var margin = chart.SecondarySeries.Add("Margin %", "Sales!$D$2:$D$5", "Sales!$A$2:$A$5");
margin.UseSecondaryAxis = true;
```

It works on any series of a chart type that has one category and one value axis — bar, column, line,
area, radar and stock — including series of the primary chart type. Pie and doughnut charts have no
value axis, scatter and bubble charts already have two, and surface charts add a series axis, so all
of those ignore it.

:::caution
`UseSecondaryAxis` can only be set on charts you create. Moving a series of a chart **loaded from
a file** onto a secondary axis would mean regrouping the chart's XML, which XLibur does not do, so
the setter throws `NotSupportedException`. The colour and marker properties have no such limit.
:::

## Data labels

`DataLabels` exists on the chart and on each series. Set it on the chart to label every series the
same way, then override the odd series that needs something different:

```csharp
var chart = ws.Charts.Add(XLChartType.ColumnClustered);
chart.Series.Add("Revenue", "Sales!$B$2:$B$5", "Sales!$A$2:$A$5");
var cost = chart.Series.Add("Cost", "Sales!$C$2:$C$5", "Sales!$A$2:$A$5");

chart.DataLabels.ShowValue = true;
chart.DataLabels.NumberFormat = "$ #,##0";
chart.DataLabels.Position = XLDataLabelPosition.OutsideEnd;

cost.DataLabels.ShowValue = true;
cost.DataLabels.ShowSeriesName = true;
cost.DataLabels.Position = XLDataLabelPosition.InsideEnd;
```

Pie and doughnut charts usually want the share rather than the number:

```csharp
var pie = ws.Charts.Add(XLChartType.Pie);
var share = pie.Series.Add("Revenue", "Sales!$B$2:$B$5", "Sales!$A$2:$A$5");
share.DataLabels.ShowCategoryName = true;
share.DataLabels.ShowPercentage = true;
share.DataLabels.Position = XLDataLabelPosition.BestFit;
```

| Property | Meaning | Default |
|---|---|---|
| `ShowValue` | the point's value | `false` |
| `ShowCategoryName` | the category axis label | `false` |
| `ShowSeriesName` | the series name | `false` |
| `ShowPercentage` | share of the total, pie and doughnut only | `false` |
| `NumberFormat` | format code for the label, e.g. `"0.0%"` | from the source cells |
| `Position` | where the label sits | `Auto` |

Nothing is written to the file until one of these is set, so an untouched chart keeps Excel's own
defaults.

### Label positions

What Excel offers depends on the chart type, and it refuses to open a file that uses a position it
does not offer — so the setter throws `ArgumentException` for a combination Excel would reject, with
the allowed values in the message:

| Chart type | Positions |
|---|---|
| clustered bar and column | `Center`, `InsideEnd`, `InsideBase`, `OutsideEnd` |
| stacked bar and column | `Center`, `InsideEnd`, `InsideBase` |
| line, scatter, radar | `Center`, `Left`, `Right`, `Above`, `Below` |
| pie | `BestFit`, `Center`, `InsideEnd`, `OutsideEnd` |
| area, doughnut, bubble, stock, every 3D type | `Auto` only |

`Auto` — the default — writes no position and lets Excel place the label. In a combo chart, a
secondary series is judged against `SecondaryChartType`, so a line series over columns takes the line
positions.

:::note
Surface charts and the extended (Office 2016+) types have no data label support in the file format
and ignore these properties. If a chart type change makes an already-set position invalid, the
position is dropped on save rather than written — the alternative is a file Excel refuses to open.
:::

## Legend

A chart XLibur creates has no legend until you ask for one:

```csharp
chart.Legend.Visible = true;
chart.Legend.Position = XLLegendPosition.Bottom;   // Right (default), Bottom, Left, Top, TopRight
chart.Legend.Overlay = true;                       // draw over the plot area rather than beside it
```

Setting `Visible = false` on a chart read from a file removes the legend it came with.

## Axes

`CategoryAxis` is the horizontal axis and `ValueAxis` the vertical one. `SecondaryValueAxis` is the
one on the right, which exists while a series has `UseSecondaryAxis` set:

```csharp
chart.CategoryAxis.Title = "Quarter";

chart.ValueAxis.Title = "Revenue";
chart.ValueAxis.NumberFormat = "$ #,##0";
chart.ValueAxis.Min = 0;
chart.ValueAxis.Max = 60_000;
chart.ValueAxis.MajorUnit = 10_000;
chart.ValueAxis.MajorGridlines = true;

chart.SecondaryValueAxis.Title = "Margin";
chart.SecondaryValueAxis.NumberFormat = "0%";
```

| Property | Meaning | Default |
|---|---|---|
| `Title` | axis title text | `null` — no title |
| `NumberFormat` | format code for the labels | from the source cells |
| `Min` / `Max` | ends of the scale | chosen from the data |
| `MajorUnit` / `MinorUnit` | tick intervals | chosen by Excel |
| `Visible` | whether the axis is drawn | `true` |
| `MajorGridlines` | gridlines across the plot | `false` |
| `Orientation` | `MinMax`, or `MaxMin` to reverse the axis | `MinMax` |
| `LogScale` / `LogBase` | logarithmic scale, base 2 to 1000 | `false` / `10` |

A log scale is what makes values spanning orders of magnitude readable:

```csharp
chart.ValueAxis.LogScale = true;
chart.ValueAxis.LogBase = 10;
```

:::note
`MajorUnit`, `MinorUnit` and `LogScale` belong to a value axis in the file format and are skipped on
a category axis — except on a scatter or bubble chart, where the horizontal axis holds numbers and
takes them too. Pie and doughnut charts have no axes at all and ignore everything here.
:::

### Editing charts loaded from a file

XLibur never regenerates the XML of a chart it read from a file — it patches in just the
properties you assign. Everything it does not model stays exactly as Excel wrote it: trendlines,
error bars, gradient and picture fills, per-point colours and label overrides, label and axis fonts,
tick marks, the chart's style and colour parts.

```csharp
using var workbook = new XLWorkbook("Report.xlsx");
var chart = workbook.Worksheet("Data").Charts.First();

chart.Series.First().FillColor = XLColor.FromHtml("#C00000");
chart.ValueAxis.Max = 60_000;
workbook.Save();   // only the fill and the scale change; the trendline stays
```

Chart type, title and series references are settable on charts you create, but on a loaded chart
only the series formatting, data labels, legend and axes are written back.

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

Charts anchored any of the three ways are listed, as are the chart types XLibur reads but does not
write itself: 3D pie, line, area and surface groups, and pie-of-pie / bar-of-pie.

Chart type and title are settable after creation:

```csharp
var chart = ws.Charts.First();
chart.ChartType = XLChartType.BarClustered;
chart.Title = "Revised";
chart.Title = null;               // remove the title
```

## What is not covered

The chart API covers type, title, series, series formatting, data labels, legend, axes and placement.
Fonts, fills and borders of the chart furniture, gradient and picture fills, per-data-point
formatting, trend lines and error bars are not exposed — though a chart read from a file keeps all of
them. If you need that level of control over a chart you are creating, the usual approach is to build
a template workbook in Excel with the chart formatted exactly as you want it, then use XLibur to write
the source data the chart already points at:

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

// Chart 1: revenue vs cost, with cost played down
var bars = ws.Charts.Add(XLChartType.ColumnClustered);
bars.SetTitle("Revenue vs cost");
bars.Series.Add("Revenue", $"{sheet}!$B$2:$B${last}", categories).FillColor =
    XLColor.FromHtml("#4472C4");
bars.Series.Add("Cost", $"{sheet}!$C$2:$C${last}", categories).FillColor =
    XLColor.FromHtml("#A5A5A5");
bars.Position.SetColumn(5).SetRow(1);
bars.SecondPosition.SetColumn(13).SetRow(16);

// Chart 2: combo — revenue bars with the margin line on its own axis
var combo = ws.Charts.Add(XLChartType.ColumnClustered);
combo.SetTitle("Revenue and margin");
combo.Series.Add("Revenue", $"{sheet}!$B$2:$B${last}", categories);
combo.SecondaryChartType = XLChartType.LineWithMarkers;

var margin = combo.SecondarySeries.Add("Margin %", $"{sheet}!$D$2:$D${last}", categories);
margin.UseSecondaryAxis = true;
margin.LineColor = XLColor.FromHtml("#ED7D31");
margin.LineWidthPt = 2.25;
margin.MarkerStyle = XLMarkerStyle.Circle;
margin.MarkerSize = 7;

combo.Position.SetColumn(5).SetRow(18);
combo.SecondPosition.SetColumn(13).SetRow(33);

combo.Legend.Visible = true;
combo.Legend.Position = XLLegendPosition.Bottom;
combo.ValueAxis.Title = "Currency";
combo.ValueAxis.NumberFormat = "$ #,##0";
combo.ValueAxis.MajorGridlines = true;
combo.SecondaryValueAxis.Title = "Margin";
combo.SecondaryValueAxis.NumberFormat = "0%";

// Chart 3: share of annual revenue
var pie = ws.Charts.Add(XLChartType.Pie);
pie.SetTitle("Share of annual revenue");
var share = pie.Series.Add("Revenue", $"{sheet}!$B$2:$B${last}", categories);
share.DataLabels.ShowCategoryName = true;
share.DataLabels.ShowPercentage = true;
share.DataLabels.Position = XLDataLabelPosition.BestFit;
pie.Position.SetColumn(14).SetRow(1);
pie.SecondPosition.SetColumn(21).SetRow(16);

workbook.SaveAs("QuarterlyCharts.xlsx");
```

## Where to next

- [Sparklines](./sparklines.md) — in-cell mini charts, when a full chart is too much
- [Images and Pictures](./images.md) — the other kind of drawing, and shared anchor rules
- [Pivot Tables](./pivot-tables.md) — summarising the data a chart plots
