---
id: theming
title: Theming
sidebar_label: Theming
description: Set the workbook theme colour scheme, use theme colours and tints in styles, and apply built-in table and pivot table themes.
---

# Theming

An Excel workbook carries a **theme** — a palette of twelve named colours that styles refer to
by role rather than by value. Style a header with `Accent1` instead of `#4F81BD` and the whole
document restyles when the theme changes, exactly as it does when a user picks a different
theme from Excel's Page Layout tab.

XLibur exposes the colour scheme through `workbook.Theme`, and lets any style reference it with
`XLColor.FromTheme`.

## The colour scheme

`IXLTheme` has twelve slots:

| Slot | Role |
|---|---|
| `Background1` | Primary background (light) |
| `Text1` | Primary text (dark) |
| `Background2` | Secondary background |
| `Text2` | Secondary text |
| `Accent1` … `Accent6` | The six accent colours used by charts, tables, and pivot styles |
| `Hyperlink` | Unvisited link colour |
| `FollowedHyperlink` | Visited link colour |

The defaults match Excel's classic Office theme:

```csharp
Background1        #FFFFFFFF     Text1              #FF000000
Background2        #FFEEECE1     Text2              #FF1F497D
Accent1            #FF4F81BD     Accent2            #FFC0504D
Accent3            #FF9BBB59     Accent4            #FF8064A2
Accent5            #FF4BACC6     Accent6            #FFF79646
Hyperlink          #FF0000FF     FollowedHyperlink  #FF800080
```

## Setting the workbook theme

Assign to the slots you want to change:

```csharp
using XLibur.Excel;

using var workbook = new XLWorkbook();

workbook.Theme.Accent1 = XLColor.FromHtml("#FF1F6FEB");
workbook.Theme.Accent2 = XLColor.FromHtml("#FF2DA44E");
workbook.Theme.Accent3 = XLColor.FromHtml("#FFBF3989");
workbook.Theme.Text2 = XLColor.FromHtml("#FF1F2328");
workbook.Theme.Background2 = XLColor.FromHtml("#FFF6F8FA");
workbook.Theme.Hyperlink = XLColor.FromHtml("#FF0969DA");
```

A whole scheme in one go:

```csharp
static void ApplyBrandTheme(IXLTheme theme)
{
    theme.Background1 = XLColor.FromHtml("#FFFFFFFF");
    theme.Text1 = XLColor.FromHtml("#FF1F2328");
    theme.Background2 = XLColor.FromHtml("#FFF6F8FA");
    theme.Text2 = XLColor.FromHtml("#FF24292F");
    theme.Accent1 = XLColor.FromHtml("#FF1F6FEB");
    theme.Accent2 = XLColor.FromHtml("#FF2DA44E");
    theme.Accent3 = XLColor.FromHtml("#FFBF3989");
    theme.Accent4 = XLColor.FromHtml("#FF9A6700");
    theme.Accent5 = XLColor.FromHtml("#FFCF222E");
    theme.Accent6 = XLColor.FromHtml("#FF8250DF");
    theme.Hyperlink = XLColor.FromHtml("#FF0969DA");
    theme.FollowedHyperlink = XLColor.FromHtml("#FF8250DF");
}

using var workbook = new XLWorkbook();
ApplyBrandTheme(workbook.Theme);
```

:::warning
The theme part is written **only when the file does not already have one** — that is, for
workbooks you create from scratch. Loading an existing `.xlsx` and changing `workbook.Theme`
will not change the saved file, because the original theme part is carried through untouched.

To restyle an existing workbook, either build a new `XLWorkbook` and copy the sheets into it,
or set explicit colours on the styles rather than relying on theme slots.
:::

Reading the current scheme back — including from a loaded workbook, where it reflects whatever
theme that file carries:

```csharp
using var workbook = new XLWorkbook("Report.xlsx");

Console.WriteLine(workbook.Theme.Accent1);
Console.WriteLine(workbook.Theme.ResolveThemeColor(XLThemeColor.Accent1));
```

## Using theme colours in styles

`XLColor.FromTheme` produces a colour that *references* a theme slot rather than baking in an
RGB value:

```csharp
var ws = workbook.Worksheets.Add("Report");

ws.Range("A1:F1").Style
    .Fill.SetBackgroundColor(XLColor.FromTheme(XLThemeColor.Accent1))
    .Font.SetFontColor(XLColor.FromTheme(XLThemeColor.Background1));

ws.Range("A2:F20").Style.Border.BottomBorderColor = XLColor.FromTheme(XLThemeColor.Background2);
```

### Tints

The two-argument overload applies a *tint*: a value from `-1.0` (fully darkened) through `0`
(the colour itself) to `1.0` (fully lightened). This is how Excel builds the shade ramps you
see under each theme colour in its colour picker:

```csharp
XLColor.FromTheme(XLThemeColor.Accent1, -0.5);   // 50% darker
XLColor.FromTheme(XLThemeColor.Accent1, 0.0);    // the accent itself
XLColor.FromTheme(XLThemeColor.Accent1, 0.4);    // 40% lighter
XLColor.FromTheme(XLThemeColor.Accent1, 0.8);    // very pale — good for banding
```

A banded table built entirely from one accent:

```csharp
var accent = XLThemeColor.Accent1;

ws.Range("A1:F1").Style
    .Fill.SetBackgroundColor(XLColor.FromTheme(accent, -0.25))
    .Font.SetFontColor(XLColor.FromTheme(XLThemeColor.Background1))
    .Font.SetBold();

for (var row = 2; row <= 20; row += 2)
{
    ws.Range($"A{row}:F{row}").Style.Fill.BackgroundColor = XLColor.FromTheme(accent, 0.8);
}
```

### Theme fonts

Fonts have the same idea. `XLFontScheme.Major` is the theme's heading font and
`XLFontScheme.Minor` its body font:

```csharp
ws.Range("A1:F1").Style.Font.FontScheme = XLFontScheme.Major;
ws.Range("A2:F20").Style.Font.FontScheme = XLFontScheme.Minor;
ws.Cell("H1").Style.Font.FontScheme = XLFontScheme.None;   // use FontName verbatim
```

## Table themes

Excel tables have their own catalogue of styles, independent of the colour scheme but drawing
their colours from it. Set one through `IXLTable.Theme`:

```csharp
var table = ws.Range("A1:F20").CreateTable("Sales");

table.Theme = XLTableTheme.TableStyleMedium2;
table.Theme = XLTableTheme.TableStyleLight9;
table.Theme = XLTableTheme.TableStyleDark3;
table.Theme = XLTableTheme.None;
```

There are 60 built-in styles, in three families:

| Family | Range | Look |
|---|---|---|
| `TableStyleLight1` – `TableStyleLight21` | 21 styles | Outlines and light banding |
| `TableStyleMedium1` – `TableStyleMedium28` | 28 styles | Filled header, banded rows |
| `TableStyleDark1` – `TableStyleDark11` | 11 styles | Solid dark fills |

Within each family the styles cycle through the theme's accent colours, so
`TableStyleMedium2` follows `Accent1`, `TableStyleMedium3` follows `Accent2`, and so on.
Change `workbook.Theme.Accent1` and every `TableStyleMedium2` table in the workbook follows.

Fine-grained switches on top of the chosen style:

```csharp
table.ShowRowStripes = true;
table.ShowColumnStripes = false;
table.EmphasizeFirstColumn = true;
table.EmphasizeLastColumn = true;
```

## Pivot table themes

Pivot tables use a parallel catalogue via the `XLPivotTableTheme` enum:

```csharp
pivot.Theme = XLPivotTableTheme.PivotStyleMedium9;
pivot.Theme = XLPivotTableTheme.PivotStyleLight16;
pivot.Theme = XLPivotTableTheme.PivotStyleDark4;
pivot.Theme = XLPivotTableTheme.None;

pivot.SetShowRowStripes()
     .SetShowColumnStripes(false);
```

## Tab colours

Sheet tabs accept theme colours too, which is a cheap way to make a multi-sheet workbook read
as one document:

```csharp
workbook.Worksheet("Summary").SetTabColor(XLColor.FromTheme(XLThemeColor.Accent1));
workbook.Worksheet("Detail").SetTabColor(XLColor.FromTheme(XLThemeColor.Accent1, 0.4));
workbook.Worksheet("Notes").SetTabColor(XLColor.FromTheme(XLThemeColor.Accent1, 0.7));
```

## A worked example

```csharp
using XLibur.Excel;

using var workbook = new XLWorkbook();

// 1. Brand colours in the theme
workbook.Theme.Accent1 = XLColor.FromHtml("#FF1F6FEB");
workbook.Theme.Accent2 = XLColor.FromHtml("#FF2DA44E");
workbook.Theme.Hyperlink = XLColor.FromHtml("#FF0969DA");

var ws = workbook.Worksheets.Add("Q1");
ws.SetTabColor(XLColor.FromTheme(XLThemeColor.Accent1));

// 2. Data
string[] headers = ["Region", "Revenue", "Growth"];
for (var i = 0; i < headers.Length; i++)
{
    ws.Cell(1, i + 1).Value = headers[i];
}

var rows = new[]
{
    ("North", 128_400m, 0.12),
    ("South", 96_100m, -0.03),
    ("East", 154_700m, 0.21),
    ("West", 71_900m, 0.05),
};

var r = 2;
foreach (var (region, revenue, growth) in rows)
{
    ws.Cell(r, 1).Value = region;
    ws.Cell(r, 2).Value = revenue;
    ws.Cell(r, 3).Value = growth;
    r++;
}

// 3. Style entirely from the theme — no hard-coded colours
ws.Range(1, 1, 1, 3).Style
    .Fill.SetBackgroundColor(XLColor.FromTheme(XLThemeColor.Accent1))
    .Font.SetFontColor(XLColor.FromTheme(XLThemeColor.Background1))
    .Font.SetFontScheme(XLFontScheme.Major)
    .Font.SetBold();

ws.Range(2, 1, r - 1, 3).Style
    .Font.SetFontScheme(XLFontScheme.Minor)
    .Border.SetBottomBorder(XLBorderStyleValues.Thin)
    .Border.SetBottomBorderColor(XLColor.FromTheme(XLThemeColor.Accent1, 0.7));

ws.Range(2, 2, r - 1, 2).Style.NumberFormat.Format = "$ #,##0";
ws.Range(2, 3, r - 1, 3).Style.NumberFormat.Format = "0.0%;[Red]-0.0%";

ws.Columns().AdjustToContents();
workbook.SaveAs("ThemedReport.xlsx");
```

## Where to next

- [Styling](./styling.md) — the full style API these colours plug into
- [Tables](./tables.md) — table themes in context
- [Fonts](./fonts.md) — the font engine behind `AdjustToContents`
