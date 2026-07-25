---
id: comments-and-hyperlinks
title: Comments and Hyperlinks
sidebar_label: Comments and Hyperlinks
description: Attach comments to cells with formatted text and styling, and add internal or external hyperlinks.
---

# Comments and Hyperlinks

Two ways to attach something to a cell beyond its value: a **comment** (Excel's classic yellow
note, shown on hover) and a **hyperlink** (clickable navigation to a URL, a file, or another
part of the workbook).

## Comments

### Creating

`CreateComment()` makes a new comment; `GetComment()` returns the existing one, creating it if
there isn't one yet — so in practice `GetComment()` is what you want:

```csharp
using XLibur.Excel;

var ws = workbook.Worksheet("Data");

ws.Cell("B2").GetComment().AddText("Reviewed on 2026-01-21");

if (ws.Cell("B2").HasComment)
{
    Console.WriteLine(ws.Cell("B2").GetComment().Text);
}

ws.Cell("B2").GetComment().Delete();
```

Comments are rich text, so `AddText` appends runs you can format independently:

```csharp
ws.Cell("C4").GetComment()
    .AddText("Warning: ").SetBold().SetFontColor(XLColor.Red)
    .AddText("this figure is provisional.").SetItalic();
```

### Author and signature

A comment carries an author name. `AddSignature()` prepends a bold line with it, which is what
Excel does when a user adds a note:

```csharp
using var workbook = new XLWorkbook { Author = "Reporting Service" };
var ws = workbook.Worksheets.Add("Data");

var comment = ws.Cell("A1").GetComment();
comment.SetAuthor("Data Team");
comment.AddSignature();
comment.AddText("Figures are unaudited.");
```

The workbook's `Author` is the default for new comments.

### Visibility

Comments are hidden by default and appear on hover. Pin one open with `SetVisible()`:

```csharp
ws.Cell("A1").GetComment().SetVisible();
ws.Cell("A1").GetComment().SetVisible(false);

// Show every comment on the sheet
foreach (var cell in ws.CellsUsed(XLCellsUsedOptions.All, c => c.HasComment))
{
    cell.GetComment().SetVisible();
}
```

### Positioning and size

A comment is a drawing, so it has the same anchor model as a picture, plus its own size:

```csharp
var comment = ws.Cell("B2").GetComment();

comment.Position.SetColumn(4).SetRow(2);

comment.Style
    .Size.SetWidth(30)      // in column-width units, like IXLColumn.Width
    .Size.SetHeight(10);    // in row-height units, like IXLRow.Height

// Or let Excel size it to the text
comment.Style.Size.SetAutomaticSize();
```

### Appearance

`Style` groups the same options Excel exposes in the Format Comment dialog:

```csharp
var comment = ws.Cell("B2").GetComment();

comment.Style
    .ColorsAndLines.SetFillColor(XLColor.LightYellow)
    .ColorsAndLines.SetLineColor(XLColor.DarkGray)
    .ColorsAndLines.SetLineWeight(1)
    .ColorsAndLines.SetFillTransparency(0.1);

comment.Style
    .Alignment.SetHorizontal(XLDrawingHorizontalAlignment.Left)
    .Alignment.SetVertical(XLDrawingVerticalAlignment.Top)
    .Alignment.SetDirection(XLDrawingTextDirection.Context);

comment.Style.Margins.SetAll(0.1);

comment.Style
    .Protection.SetLocked(true)
    .Protection.SetLockText(true);

// Alternate text for accessibility / web export
comment.Style.Web.AlternateText = "Provisional figure note";
```

Style groups available: `Alignment`, `ColorsAndLines`, `Size`, `Margins`, `Protection`,
`Properties`, `Web`.

### A note on comments vs threaded comments

These are Excel's *legacy* notes — the yellow sticky kind. Modern threaded comments (the ones
with replies, introduced in Excel 2016) are a different part of the file format and are not
created by XLibur. Excel displays legacy notes fine; it just labels them "Notes" rather than
"Comments" in its newer UI.

## Hyperlinks

### External links

```csharp
ws.Cell("A1").Value = "XLibur on GitHub";
ws.Cell("A1").SetHyperlink(new XLHyperlink("https://github.com/XLibur/XLibur"));

// With a tooltip
ws.Cell("A2").Value = "Documentation";
ws.Cell("A2").SetHyperlink(new XLHyperlink("https://xlibur.github.io/XLibur/", "Open the docs"));

// From a Uri
ws.Cell("A3").SetHyperlink(new XLHyperlink(new Uri("https://example.com/report")));
```

Email and file links use the same constructor — the scheme decides the behaviour:

```csharp
ws.Cell("A4").Value = "Email support";
ws.Cell("A4").SetHyperlink(new XLHyperlink("mailto:support@example.com?subject=Report"));

ws.Cell("A5").Value = "Source data";
ws.Cell("A5").SetHyperlink(new XLHyperlink(@"file:///C:\Reports\source.csv"));
```

:::note
Setting a hyperlink does not change the cell's text. Assign `Value` as well, or the cell will
be clickable but blank.
:::

### Internal links

Pass a cell or a range to link within the workbook:

```csharp
var summary = workbook.Worksheet("Summary");

ws.Cell("B1").Value = "Back to summary";
ws.Cell("B1").SetHyperlink(new XLHyperlink(summary.Cell("A1")));

ws.Cell("B2").Value = "Jump to the detail block";
ws.Cell("B2").SetHyperlink(new XLHyperlink(ws.Range("D40:H60")));

// Or by address string
ws.Cell("B3").Value = "Q1 figures";
ws.Cell("B3").SetHyperlink(new XLHyperlink("'Q1 Sales'!A1"));
```

A defined name works too, and survives the target moving:

```csharp
ws.Range("D40:H60").AddToNamed("DetailBlock");
ws.Cell("B4").SetHyperlink(new XLHyperlink("DetailBlock"));
```

### Reading hyperlinks back

```csharp
var cell = ws.Cell("A1");

if (cell.HasHyperlink)
{
    var link = cell.GetHyperlink();

    Console.WriteLine(link.IsExternal
        ? $"external: {link.ExternalAddress}"
        : $"internal: {link.InternalAddress}");

    Console.WriteLine(link.Tooltip);
}

// Every hyperlink on the sheet. Cell is null for a hyperlink
// not attached to a worksheet, so guard before dereferencing.
foreach (var link in ws.Hyperlinks)
{
    Console.WriteLine($"{link.Cell?.Address.ToString() ?? "(detached)"}: {link.Tooltip}");
}
```

### Removing

```csharp
ws.Cell("A1").SetHyperlink(null);
ws.Hyperlinks.Delete(ws.Cell("A1").Address);
```

Removing a hyperlink also resets the cell's font colour and underline back to the sheet default
if it was using the theme hyperlink style.

### Styling

XLibur does not automatically apply the blue-and-underlined look. Use the theme's hyperlink
colour so it stays in step with the rest of the workbook:

```csharp
ws.Cell("A1").Style
    .Font.SetFontColor(XLColor.FromTheme(XLThemeColor.Hyperlink))
    .Font.SetUnderline(XLFontUnderlineValues.Single);
```

### The HYPERLINK function

For a link whose target is computed, the worksheet function is often simpler than a real
hyperlink object — and it recalculates:

```csharp
ws.Cell("C2").FormulaA1 = "=HYPERLINK(\"https://tracker.example.com/\" & A2, \"Open \" & A2)";
```

## A worked example

An index sheet linking to every data sheet, with a comment explaining each one:

```csharp
using XLibur.Excel;

using var workbook = new XLWorkbook { Author = "Reporting Service" };

var index = workbook.Worksheets.Add("Index");
index.Cell("A1").Value = "Report contents";
index.Cell("A1").Style.Font.SetBold().Font.SetFontSize(14);

var sheets = new[]
{
    ("Sales", "Revenue by region and quarter."),
    ("Costs", "Direct and indirect costs, excluding tax."),
    ("Margin", "Derived from Sales and Costs; do not edit directly."),
};

var row = 3;
foreach (var (name, description) in sheets)
{
    var sheet = workbook.Worksheets.Add(name);
    sheet.Cell("A1").Value = name;

    // Link from the index to the sheet
    index.Cell(row, 1).Value = name;
    index.Cell(row, 1).SetHyperlink(new XLHyperlink(sheet.Cell("A1"), $"Go to {name}"));
    index.Cell(row, 1).Style
        .Font.SetFontColor(XLColor.FromTheme(XLThemeColor.Hyperlink))
        .Font.SetUnderline(XLFontUnderlineValues.Single);

    // Explain it in a comment
    var comment = index.Cell(row, 1).GetComment();
    comment.SetAuthor("Reporting Service");
    comment.AddSignature();
    comment.AddText(description);
    comment.Style
        .Size.SetAutomaticSize()
        .ColorsAndLines.SetFillColor(XLColor.LightYellow);

    // And a link back
    sheet.Cell("C1").Value = "Back to index";
    sheet.Cell("C1").SetHyperlink(new XLHyperlink(index.Cell("A1")));

    row++;
}

index.Columns().AdjustToContents();
workbook.SaveAs("Report.xlsx");
```

## Where to next

- [Cells and Ranges](./cells-and-ranges.md) — addressing the cells these attach to
- [Defined Names](./defined-names.md) — stable link targets that survive edits
- [Styling](./styling.md) — rich text, which comments reuse
