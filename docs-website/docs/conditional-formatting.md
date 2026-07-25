---
id: conditional-formatting
title: Conditional Formatting
sidebar_label: Conditional Formatting
description: Style cells based on their values — comparison rules, text rules, colour scales, data bars, icon sets, and formula-driven rules.
---

# Conditional Formatting

Conditional formatting applies a style to a cell only when a condition holds. Unlike a direct
style, the rule travels with the file and re-evaluates whenever the data changes — so a
"highlight overdue rows" rule keeps working after the user edits the sheet.

Every rule starts with `AddConditionalFormat()` on a range, then a `When…` method that returns
the `IXLStyle` to apply:

```csharp
using XLibur.Excel;

var ws = workbook.Worksheet("Data");

ws.Range("B2:B100").AddConditionalFormat()
    .WhenGreaterThan(1000)
    .Fill.SetBackgroundColor(XLColor.LightGreen)
    .Font.SetBold();
```

Multiple rules can target the same range; Excel evaluates them in the order they were added.

## Comparison rules

Numeric and string overloads exist for each:

```csharp
var range = ws.Range("B2:B100");

range.AddConditionalFormat().WhenEquals(0).Fill.SetBackgroundColor(XLColor.LightGray);
range.AddConditionalFormat().WhenNotEquals(0).Font.SetBold();
range.AddConditionalFormat().WhenGreaterThan(1000).Font.SetFontColor(XLColor.DarkGreen);
range.AddConditionalFormat().WhenLessThan(0).Font.SetFontColor(XLColor.Red);
range.AddConditionalFormat().WhenEqualOrGreaterThan(500).Fill.SetBackgroundColor(XLColor.LightYellow);
range.AddConditionalFormat().WhenEqualOrLessThan(100).Fill.SetBackgroundColor(XLColor.MistyRose);
range.AddConditionalFormat().WhenBetween(100, 500).Fill.SetBackgroundColor(XLColor.LightBlue);
range.AddConditionalFormat().WhenNotBetween(100, 500).Font.SetItalic();
```

## Text rules

```csharp
var range = ws.Range("A2:A100");

range.AddConditionalFormat().WhenContains("URGENT").Font.SetFontColor(XLColor.Red);
range.AddConditionalFormat().WhenNotContains("draft").Font.SetBold();
range.AddConditionalFormat().WhenStartsWith("TMP").Fill.SetBackgroundColor(XLColor.LightGray);
range.AddConditionalFormat().WhenEndsWith("_old").Font.SetStrikethrough();
```

## Blank, error, duplicate, and unique

```csharp
var range = ws.Range("A2:D100");

range.AddConditionalFormat().WhenIsBlank().Fill.SetBackgroundColor(XLColor.LightGray);
range.AddConditionalFormat().WhenNotBlank().Border.SetBottomBorder(XLBorderStyleValues.Hair);
range.AddConditionalFormat().WhenIsError().Fill.SetBackgroundColor(XLColor.Salmon);
range.AddConditionalFormat().WhenNotError().Font.SetFontColor(XLColor.Black);

range.AddConditionalFormat().WhenIsDuplicate().Fill.SetBackgroundColor(XLColor.Yellow);
range.AddConditionalFormat().WhenIsUnique().Font.SetBold();
```

## Top and bottom

```csharp
var range = ws.Range("B2:B100");

range.AddConditionalFormat().WhenIsTop(10).Fill.SetBackgroundColor(XLColor.LightGreen);
range.AddConditionalFormat().WhenIsTop(10, XLTopBottomType.Percent).Font.SetBold();
range.AddConditionalFormat().WhenIsBottom(5, XLTopBottomType.Items).Fill.SetBackgroundColor(XLColor.MistyRose);
```

## Date rules

`WhenDateIs` takes a relative time period, evaluated against the day the file is opened:

```csharp
var dates = ws.Range("C2:C100");

dates.AddConditionalFormat().WhenDateIs(XLTimePeriod.Today).Font.SetBold();
dates.AddConditionalFormat().WhenDateIs(XLTimePeriod.Yesterday).Font.SetFontColor(XLColor.Gray);
dates.AddConditionalFormat().WhenDateIs(XLTimePeriod.InTheLast7Days).Fill.SetBackgroundColor(XLColor.LightYellow);
dates.AddConditionalFormat().WhenDateIs(XLTimePeriod.NextMonth).Fill.SetBackgroundColor(XLColor.LightBlue);
```

Periods: `Yesterday`, `Today`, `Tomorrow`, `InTheLast7Days`, `LastWeek`, `ThisWeek`,
`NextWeek`, `LastMonth`, `ThisMonth`, `NextMonth`.

## Formula rules

The most flexible option: `WhenIsTrue` takes a formula written relative to the **top-left cell
of the range**, and applies the style wherever it evaluates true. This is how you style a whole
row based on one column:

```csharp
// Highlight the entire row when column E says "Overdue"
ws.Range("A2:F100").AddConditionalFormat()
    .WhenIsTrue("=$E2=\"Overdue\"")
    .Fill.SetBackgroundColor(XLColor.MistyRose);

// Alternating row banding
ws.Range("A2:F100").AddConditionalFormat()
    .WhenIsTrue("=MOD(ROW(),2)=0")
    .Fill.SetBackgroundColor(XLColor.FromHtml("#FFF6F8FA"));

// Compare two columns
ws.Range("A2:F100").AddConditionalFormat()
    .WhenIsTrue("=$C2>$D2")
    .Font.SetFontColor(XLColor.Red);
```

:::note
Anchor the column with `$` (as in `$E2`) so the rule reads the same column for every cell in
the row, while the row number stays relative.
:::

## Colour scales

A two- or three-point gradient across the range. Each stop takes an `XLCFContentType` — how the
threshold value should be interpreted:

```csharp
// Three-colour scale: red → yellow → green
ws.Range("B2:B100").AddConditionalFormat().ColorScale()
    .LowestValue(XLColor.Red)
    .Midpoint(XLCFContentType.Percent, "50", XLColor.Yellow)
    .HighestValue(XLColor.Green);

// Two-colour scale between explicit thresholds
ws.Range("C2:C100").AddConditionalFormat().ColorScale()
    .Minimum(XLCFContentType.Number, "0", XLColor.White)
    .Maximum(XLCFContentType.Percentile, "90", XLColor.DarkBlue);
```

Content types: `Number`, `Percent`, `Percentile`, `Formula`, `Minimum`, `Maximum`.

## Data bars

An in-cell bar chart:

```csharp
// Simple bar
ws.Range("B2:B100").AddConditionalFormat()
    .DataBar(XLColor.CornflowerBlue)
    .LowestValue()
    .HighestValue();

// Bar only, no number shown, solid rather than gradient
ws.Range("C2:C100").AddConditionalFormat()
    .DataBar(XLColor.CornflowerBlue, showBarOnly: true, gradient: false)
    .Minimum(XLCFContentType.Number, 0)
    .Maximum(XLCFContentType.Number, 100);

// Separate colours for positive and negative values
var bars = ws.Range("D2:D100").AddConditionalFormat()
    .DataBar(XLColor.Green, XLColor.Red)
    .LowestValue()
    .HighestValue();

bars.BarAxisPosition = XLDataBarAxisPosition.Middle;
```

## Icon sets

```csharp
ws.Range("E2:E100").AddConditionalFormat()
    .IconSet(XLIconSetStyle.ThreeTrafficLights1)
    .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, 0, XLCFContentType.Percent)
    .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, 33, XLCFContentType.Percent)
    .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, 67, XLCFContentType.Percent);

// Reverse the icon order and hide the underlying value
ws.Range("F2:F100").AddConditionalFormat()
    .IconSet(XLIconSetStyle.FiveArrows, reverseIconOrder: true, showIconOnly: true)
    .AddValue(XLCFIconSetOperator.GreaterThan, 0, XLCFContentType.Percentile)
    .AddValue(XLCFIconSetOperator.GreaterThan, 20, XLCFContentType.Percentile)
    .AddValue(XLCFIconSetOperator.GreaterThan, 40, XLCFContentType.Percentile)
    .AddValue(XLCFIconSetOperator.GreaterThan, 60, XLCFContentType.Percentile)
    .AddValue(XLCFIconSetOperator.GreaterThan, 80, XLCFContentType.Percentile);
```

Add one threshold per icon in the set — three for the `Three…` styles, four for `Four…`, five
for `Five…`.

Available styles: `ThreeArrows`, `ThreeArrowsGray`, `ThreeFlags`, `ThreeTrafficLights1`,
`ThreeTrafficLights2`, `ThreeSigns`, `ThreeSymbols`, `ThreeSymbols2`, `FourArrows`,
`FourArrowsGray`, `FourRedToBlack`, `FourRating`, `FourTrafficLights`, `FiveArrows`,
`FiveArrowsGray`, `FiveRating`, `FiveQuarters`.

## Managing rules

```csharp
// Every rule on the sheet
foreach (var format in ws.ConditionalFormats)
{
    Console.WriteLine($"{format.ConditionalFormatType} on {format.Range.RangeAddress}");
}

// Remove them
ws.ConditionalFormats.RemoveAll();

// Or clear just one range's rules
ws.Range("B2:B100").Clear(XLClearOptions.ConditionalFormats);
```

## A worked example

```csharp
using XLibur.Excel;

using var workbook = new XLWorkbook();
var ws = workbook.Worksheets.Add("Tasks");

ws.Cell("A1").Value = "Task";
ws.Cell("B1").Value = "Owner";
ws.Cell("C1").Value = "Due";
ws.Cell("D1").Value = "Progress";
ws.Cell("E1").Value = "Status";
ws.Range("A1:E1").Style.Font.Bold = true;

var tasks = new[]
{
    ("Design review", "Ada", new DateTime(2026, 2, 1), 1.00, "Done"),
    ("Migration script", "Grace", new DateTime(2026, 2, 14), 0.60, "In progress"),
    ("Load testing", "Alan", new DateTime(2026, 1, 20), 0.10, "Overdue"),
    ("Docs pass", "Ada", new DateTime(2026, 3, 2), 0.35, "In progress"),
};

var row = 2;
foreach (var (task, owner, due, progress, status) in tasks)
{
    ws.Cell(row, 1).Value = task;
    ws.Cell(row, 2).Value = owner;
    ws.Cell(row, 3).Value = due;
    ws.Cell(row, 4).Value = progress;
    ws.Cell(row, 5).Value = status;
    row++;
}

var last = row - 1;
ws.Range($"C2:C{last}").Style.DateFormat.Format = "yyyy-MM-dd";
ws.Range($"D2:D{last}").Style.NumberFormat.Format = "0%";

// Whole row red when overdue
ws.Range($"A2:E{last}").AddConditionalFormat()
    .WhenIsTrue("=$E2=\"Overdue\"")
    .Fill.SetBackgroundColor(XLColor.MistyRose);

// Progress as a data bar
ws.Range($"D2:D{last}").AddConditionalFormat()
    .DataBar(XLColor.CornflowerBlue)
    .Minimum(XLCFContentType.Number, 0)
    .Maximum(XLCFContentType.Number, 1);

// Due dates in the next week stand out
ws.Range($"C2:C{last}").AddConditionalFormat()
    .WhenDateIs(XLTimePeriod.NextWeek)
    .Font.SetBold();

ws.Columns().AdjustToContents();
workbook.SaveAs("Tasks.xlsx");
```

## Where to next

- [Styling](./styling.md) — the style API these rules produce
- [Data Validation](./data-validation.md) — constraining what users may enter
- [Sparklines](./sparklines.md) — in-cell trend charts, where a data bar is not enough
