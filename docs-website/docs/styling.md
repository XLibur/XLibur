---
id: styling
title: Styling
sidebar_label: Styling
description: Apply fonts, fills, borders, alignment, text rotation, and number formats to cells, ranges, rows, columns, and whole worksheets.
---

# Styling

Every object that holds cells — a cell, a range, a row, a column, a worksheet, the workbook
itself — exposes a `Style` property of the same type, `IXLStyle`. Setting a style on a range
applies it to every cell in that range; setting it on the workbook changes the default for
everything.

`IXLStyle` has six parts:

| Part | Controls |
|---|---|
| `Font` | Typeface, size, weight, colour, underline, super/subscript |
| `Fill` | Background colour and pattern |
| `Border` | Line style and colour on each edge, plus diagonals |
| `Alignment` | Horizontal/vertical placement, wrapping, indent, rotation |
| `NumberFormat` | How numbers are displayed |
| `Protection` | Whether the cell is locked or hidden when the sheet is protected |

## Applying a style

```csharp
var ws = workbook.Worksheet("Report");

// One cell
ws.Cell("A1").Style.Font.Bold = true;

// A range — applies to all cells in it
ws.Range("A1:D1").Style.Font.Bold = true;

// A whole column or row
ws.Column("B").Style.NumberFormat.Format = "#,##0.00";
ws.Row(1).Style.Fill.BackgroundColor = XLColor.LightGray;

// The whole sheet
ws.Style.Font.FontName = "Calibri";

// The workbook default — every new cell inherits this
workbook.Style.Font.FontSize = 11;
```

Every setter also has a fluent `Set…` form that returns the style, so a block of related
settings reads as one statement:

```csharp
ws.Range("A1:D1").Style
    .Font.SetBold()
    .Font.SetFontColor(XLColor.White)
    .Fill.SetBackgroundColor(XLColor.FromHtml("#FF4F81BD"))
    .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
    .Border.SetBottomBorder(XLBorderStyleValues.Thin);
```

### Reusing a style

Build a style once from `workbook.Style` and assign it wherever it is needed. This is both
tidier and cheaper than repeating the property sets, because XLibur interns identical styles:

```csharp
var headerStyle = workbook.Style;
headerStyle.Font.Bold = true;
headerStyle.Font.FontColor = XLColor.White;
headerStyle.Fill.BackgroundColor = XLColor.FromHtml("#FF4F81BD");
headerStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

ws.Range("A1:F1").Style = headerStyle;
ws.Range("A20:F20").Style = headerStyle;
```

:::tip
Style whole columns or ranges rather than looping over cells. A single
`ws.Range("D2:D10000").Style.NumberFormat.Format = "0.00"` stores one style; a loop over
10,000 cells does 10,000 assignments to reach the same result.
:::

## Font

```csharp
var style = ws.Cell("B2").Style;

style.Font.FontName = "Segoe UI";
style.Font.FontSize = 14;
style.Font.Bold = true;
style.Font.Italic = true;
style.Font.Underline = XLFontUnderlineValues.Single;   // None, Single, Double, SingleAccounting, DoubleAccounting
style.Font.Strikethrough = true;
style.Font.FontColor = XLColor.DarkRed;
style.Font.VerticalAlignment = XLFontVerticalTextAlignmentValues.Superscript;  // or Subscript, Baseline
```

Fluent equivalent:

```csharp
ws.Cell("B2").Style
    .Font.SetFontName("Segoe UI")
    .Font.SetFontSize(14)
    .Font.SetBold()
    .Font.SetFontColor(XLColor.DarkRed);
```

Fonts follow the workbook theme when you use `FontScheme`, which keeps text in step with
whatever theme fonts the file carries:

```csharp
ws.Cell("A1").Style.Font.FontScheme = XLFontScheme.Major;   // heading font
ws.Cell("A2").Style.Font.FontScheme = XLFontScheme.Minor;   // body font
```

### Rich text — mixed formatting in one cell

When a single cell needs more than one format, write rich text instead of a plain value:

```csharp
var cell = ws.Cell("A1");
cell.GetRichText()
    .AddText("Total: ").SetFontColor(XLColor.Gray)
    .AddText("1,240").SetBold().SetFontColor(XLColor.DarkGreen)
    .AddText(" units");
```

## Background colour

`Fill` is the cell background. In almost every case you want a solid fill, which XLibur sets
for you when you assign a background colour:

```csharp
ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.Yellow;
ws.Range("A1:D1").Style.Fill.SetBackgroundColor(XLColor.FromHtml("#FFEEECE1"));
```

Patterned fills need all three properties:

```csharp
var style = ws.Cell("B2").Style;
style.Fill.PatternType = XLFillPatternValues.LightGrid;
style.Fill.BackgroundColor = XLColor.White;
style.Fill.PatternColor = XLColor.LightBlue;
```

### Specifying colours

`XLColor` accepts named colours, HTML hex, ARGB components, Excel's indexed palette, and
theme colours:

```csharp
XLColor.Red                                  // one of ~140 named colours
XLColor.FromName("CornflowerBlue")
XLColor.FromHtml("#FF4F81BD")                // AARRGGBB or RRGGBB
XLColor.FromArgb(0x4F, 0x81, 0xBD)           // r, g, b
XLColor.FromArgb(255, 0x4F, 0x81, 0xBD)      // a, r, g, b
XLColor.FromIndex(44)                        // legacy indexed palette
XLColor.FromTheme(XLThemeColor.Accent1)      // follows the workbook theme
XLColor.FromTheme(XLThemeColor.Accent1, 0.4) // theme colour, lightened
```

Theme colours are the ones to reach for when you want the workbook to stay coherent — see
[Theming](./theming.md).

## Borders

Each edge has a style and a colour:

```csharp
var style = ws.Cell("B2").Style;

style.Border.TopBorder = XLBorderStyleValues.Thin;
style.Border.TopBorderColor = XLColor.Black;
style.Border.BottomBorder = XLBorderStyleValues.Double;
style.Border.BottomBorderColor = XLColor.Black;
style.Border.LeftBorder = XLBorderStyleValues.Thin;
style.Border.RightBorder = XLBorderStyleValues.Thin;
```

Available line styles: `None`, `Hair`, `Dotted`, `DashDotDot`, `DashDot`, `Dashed`, `Thin`,
`MediumDashDotDot`, `SlantDashDot`, `MediumDashDot`, `MediumDashed`, `Medium`, `Thick`,
`Double`.

### Outside and inside borders

On a range, `OutsideBorder` draws the perimeter and `InsideBorder` draws the grid between
cells. This is the concise way to box a table:

```csharp
var range = ws.Range("A1:D10");

range.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
range.Style.Border.OutsideBorderColor = XLColor.Black;
range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
range.Style.Border.InsideBorderColor = XLColor.LightGray;
```

:::note
`OutsideBorder` and `InsideBorder` are write-only — they are a shorthand that expands into the
individual edge settings, so there is nothing meaningful to read back.
:::

### Diagonal borders

```csharp
var style = ws.Cell("B2").Style;
style.Border.DiagonalBorder = XLBorderStyleValues.Thin;
style.Border.DiagonalBorderColor = XLColor.Red;
style.Border.DiagonalUp = true;      // bottom-left to top-right
style.Border.DiagonalDown = true;    // top-left to bottom-right
```

## Alignment and orientation

```csharp
var style = ws.Cell("B2").Style;

style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
style.Alignment.WrapText = true;
style.Alignment.ShrinkToFit = false;
style.Alignment.Indent = 2;
```

Horizontal values: `General`, `Left`, `Center`, `Right`, `Fill`, `Justify`, `CenterContinuous`,
`Distributed`.
Vertical values: `Top`, `Center`, `Bottom`, `Justify`, `Distributed`.

### Text rotation

`TextRotation` is in degrees, from `-90` (rotated clockwise) to `90` (counterclockwise). The
special value `255` stacks the characters vertically:

```csharp
// Angled column headers — the classic use
ws.Range("B1:H1").Style.Alignment.TextRotation = 45;
ws.Row(1).Height = 60;

ws.Cell("A5").Style.Alignment.TextRotation = -90;   // read bottom-to-top
ws.Cell("A6").Style.Alignment.TextRotation = 255;   // stacked letters
```

`TopToBottom` is a separate switch that lays the text out downwards rather than rotating it:

```csharp
ws.Cell("A7").Style.Alignment.TopToBottom = true;
```

### Reading order

For right-to-left content:

```csharp
ws.Cell("A1").Style.Alignment.ReadingOrder = XLAlignmentReadingOrderValues.RightToLeft;
ws.RightToLeft = true;   // flips the whole sheet
```

## Number formats

Two ways to set a format — a custom format string, or one of Excel's built-in format IDs:

```csharp
ws.Range("D2:D100").Style.NumberFormat.Format = "$ #,##0.00";
ws.Range("E2:E100").Style.NumberFormat.Format = "0.00%";
ws.Range("F2:F100").Style.NumberFormat.Format = "yyyy-MM-dd";
ws.Range("G2:G100").Style.NumberFormat.Format = "[h]:mm:ss";

// Built-in format IDs — 15 is "d-mmm-yy", 3 is "#,##0"
ws.Column("C").Style.NumberFormat.NumberFormatId = 15;
ws.Column("H").Style.NumberFormat.NumberFormatId = (int)XLPredefinedFormat.Number.IntegerWithSeparator;
```

Some formats worth keeping to hand:

| Format string | Renders `1234.5` as |
|---|---|
| `#,##0` | `1,235` |
| `#,##0.00` | `1,234.50` |
| `$ #,##0.00` | `$ 1,234.50` |
| `0.00%` | `123450.00%` |
| `0.00E+00` | `1.23E+03` |
| `#,##0;[Red](#,##0)` | `1,235` — negatives red and parenthesised |
| `yyyy-mm-dd hh:mm` | (dates) `2026-01-21 09:30` |
| `[h]:mm:ss` | (durations) elapsed hours beyond 24 |

There is also a `DateFormat` shortcut, which is the same underlying property scoped to dates:

```csharp
ws.Range("C2:C100").Style.DateFormat.Format = "dd/MM/yyyy";
```

:::note
A number format changes only how a value is *displayed*. `GetDouble()` still returns the raw
number; use `GetFormattedString()` to read it as the user sees it.
:::

## Protection

Locking only takes effect once the sheet itself is protected:

```csharp
ws.Style.Protection.Locked = true;                      // default for all cells
ws.Range("B2:B10").Style.Protection.Locked = false;     // leave these editable
ws.Range("C2:C10").Style.Protection.Hidden = true;      // hide formulas in the formula bar

ws.Protect("s3cret");
```

## A worked example

```csharp
using XLibur.Excel;

using var workbook = new XLWorkbook();
var ws = workbook.Worksheets.Add("Sales");

string[] headers = ["Region", "Q1", "Q2", "Q3", "Q4", "Total"];
for (var i = 0; i < headers.Length; i++)
{
    ws.Cell(1, i + 1).Value = headers[i];
}

var data = new[]
{
    ("North", 12000d, 13500d, 12800d, 15100d),
    ("South", 9800d, 10200d, 11000d, 11900d),
    ("East", 15300d, 14100d, 16200d, 17400d),
    ("West", 7400d, 8100d, 7900d, 8600d),
};

var row = 2;
foreach (var (region, q1, q2, q3, q4) in data)
{
    ws.Cell(row, 1).Value = region;
    ws.Cell(row, 2).Value = q1;
    ws.Cell(row, 3).Value = q2;
    ws.Cell(row, 4).Value = q3;
    ws.Cell(row, 5).Value = q4;
    ws.Cell(row, 6).FormulaA1 = $"=SUM(B{row}:E{row})";
    row++;
}

var lastRow = row - 1;

// Header band
ws.Range(1, 1, 1, 6).Style
    .Font.SetBold()
    .Font.SetFontColor(XLColor.White)
    .Fill.SetBackgroundColor(XLColor.FromHtml("#FF4F81BD"))
    .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

// Money columns
ws.Range(2, 2, lastRow, 6).Style.NumberFormat.Format = "$ #,##0";

// Emphasise the totals column
ws.Range(1, 6, lastRow, 6).Style.Font.Bold = true;

// Box the table
var table = ws.Range(1, 1, lastRow, 6);
table.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
table.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
table.Style.Border.InsideBorderColor = XLColor.LightGray;

ws.Columns().AdjustToContents();
workbook.SaveAs("StyledSales.xlsx");
```

## Where to next

- [Theming](./theming.md) — theme colours so styles stay consistent across a workbook
- [Conditional Formatting](./conditional-formatting.md) — styles driven by cell values
- [Tables](./tables.md) — built-in table styles, no manual banding required
