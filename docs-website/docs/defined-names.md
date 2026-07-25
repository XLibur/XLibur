---
id: defined-names
title: Defined Names
sidebar_label: Defined Names
description: Create workbook- and worksheet-scoped named ranges, point them at ranges or formulas, and use them in formulas, validation, and charts.
---

# Defined Names

A defined name (Excel calls them *named ranges*) is a label attached to a range or a formula.
`=SUM(SalesFigures)` reads better than `=SUM(Data!$B$2:$B$500)`, survives the range moving, and
gives you one place to change what a formula points at.

Names come in two scopes:

| Scope | Visible from | Collection |
|---|---|---|
| `Workbook` | Anywhere in the file | `workbook.DefinedNames` |
| `Worksheet` | Its own sheet, or qualified as `Sheet1!Name` | `worksheet.DefinedNames` |

## Creating

The shortest route is from the range itself:

```csharp
using XLibur.Excel;

var ws = workbook.Worksheet("Data");

ws.Range("B2:B100").AddToNamed("SalesFigures");                     // workbook scope
ws.Range("C2:C100").AddToNamed("LocalRates", XLScope.Worksheet);    // sheet scope
ws.Range("D2:D100").AddToNamed("TaxRates", XLScope.Workbook, "Rates by region");
```

Or through the collection, which also accepts an address string:

```csharp
workbook.DefinedNames.Add("SalesFigures", ws.Range("B2:B100"));
workbook.DefinedNames.Add("SalesFigures", "Data!$B$2:$B$100");
workbook.DefinedNames.Add("SalesFigures", "Data!$B$2:$B$100", "Monthly totals");

// Sheet-scoped, by using that sheet's collection
ws.DefinedNames.Add("LocalRates", ws.Range("C2:C100"));

// A name covering several disjoint ranges
workbook.DefinedNames.Add("KeyCells", ws.Ranges("A1,C3,E5:E10"));
```

A single cell needs `AsRange()` first:

```csharp
ws.Cell("A1").AsRange().AddToNamed("ReportTitle");
```

:::note
Name rules follow Excel: start with a letter or underscore, no spaces, no characters that look
like a cell reference (`A1`, `R1C1`), and up to 255 characters. Names are matched
case-insensitively.
:::

## Using them

In formulas — this is the main point:

```csharp
ws.Cell("F1").FormulaA1 = "=SUM(SalesFigures)";
ws.Cell("F2").FormulaA1 = "=AVERAGE(SalesFigures)";
ws.Cell("F3").FormulaA1 = "=SUMPRODUCT(SalesFigures, TaxRates)";
```

In data validation, where a named list is the only clean way to source a dropdown from another
sheet:

```csharp
lookups.Range("A1:A4").AddToNamed("Regions");
ws.Range("G2:G100").CreateDataValidation().List("=Regions");
```

As a hyperlink target that survives the destination moving:

```csharp
ws.Cell("B1").SetHyperlink(new XLHyperlink("DetailBlock"));
```

And as a way to apply one style to several scattered ranges at once:

```csharp
ws.Cell("A1").AsRange().AddToNamed("Titles");
ws.Range("C1:H1").AddToNamed("Titles");

var titleStyle = workbook.Style;
titleStyle.Font.Bold = true;
titleStyle.Fill.BackgroundColor = XLColor.Cyan;

workbook.DefinedNames.DefinedName("Titles").Ranges.Style = titleStyle;
```

## Looking them up

```csharp
// Throws KeyNotFoundException if absent
var name = workbook.DefinedNames.DefinedName("SalesFigures");

// Non-throwing
if (workbook.DefinedNames.TryGetValue("SalesFigures", out var found))
{
    Console.WriteLine(found.RefersTo);
}

if (workbook.DefinedNames.Contains("SalesFigures"))
{
    // ...
}

// Enumerate
foreach (var definedName in workbook.DefinedNames)
{
    Console.WriteLine($"{definedName.Name} ({definedName.Scope}) -> {definedName.RefersTo}");
}
```

Sheet-scoped names are **not** in the workbook collection — look on the sheet:

```csharp
foreach (var definedName in ws.DefinedNames)
{
    Console.WriteLine(definedName.Name);
}
```

To see everything in the file, walk both:

```csharp
var all = workbook.DefinedNames
    .Concat(workbook.Worksheets.SelectMany(sheet => sheet.DefinedNames));
```

## Working with a name

```csharp
var name = workbook.DefinedNames.DefinedName("SalesFigures");

Console.WriteLine(name.Name);        // "SalesFigures"
Console.WriteLine(name.Scope);       // Workbook or Worksheet
Console.WriteLine(name.RefersTo);    // "Data!$B$2:$B$100"
Console.WriteLine(name.Comment);
Console.WriteLine(name.Visible);     // hidden names exist and are used by some add-ins

// The ranges it resolves to
foreach (var range in name.Ranges)
{
    range.Style.Fill.BackgroundColor = XLColor.LightYellow;
    Console.WriteLine(range.RangeAddress);
}
```

### Renaming and repointing

```csharp
name.Name = "MonthlySales";
name.Comment = "Updated for FY26";
name.Visible = false;

name.SetRefersTo(ws.Range("B2:B500"));
name.SetRefersTo(ws.Ranges("B2:B500,D2:D500"));
name.RefersTo = "Data!$B$2:$B$500";
```

:::warning
Changing `Name` does **not** rewrite formulas that already use the old name — those will break.
Repointing with `SetRefersTo` is safe, because the name itself is unchanged.
:::

### Names that hold a formula

A defined name does not have to point at a range. Anything valid on the right of `RefersTo`
works, which makes names a lightweight way to define a constant or a reusable expression:

```csharp
workbook.DefinedNames.Add("VatRate", "=0.20");
workbook.DefinedNames.Add("ReportYear", "=2026");

ws.Cell("E2").FormulaA1 = "=D2*VatRate";
ws.Cell("A1").FormulaA1 = "=\"Annual report \" & ReportYear";
```

Note that `name.Ranges` is empty for these — there is no range to return.

## Deleting

Deleting a name never deletes the cells it refers to:

```csharp
workbook.DefinedNames.Delete("SalesFigures");
workbook.DefinedNames.Delete(0);          // by index
workbook.DefinedNames.DeleteAll();

// Or from the name itself
workbook.DefinedNames.DefinedName("SalesFigures").Delete();

// Sheet-scoped names
ws.DefinedNames.Delete("LocalRates");
```

## Broken references

Deleting the sheet a name points at leaves the name behind with a `#REF!` reference. Excel
tolerates this but flags it in the Name Manager, and it is a common cause of "we opened your
file and it complained" reports. Two helpers let you audit and clean up:

```csharp
foreach (var broken in workbook.DefinedNames.InvalidNamedRanges())
{
    Console.WriteLine($"broken: {broken.Name} -> {broken.RefersTo}");
}

var valid = workbook.DefinedNames.ValidNamedRanges().ToList();
```

Stripping the broken ones before saving:

```csharp
foreach (var broken in workbook.DefinedNames.InvalidNamedRanges().ToList())
{
    broken.Delete();
}

foreach (var sheet in workbook.Worksheets)
{
    foreach (var broken in sheet.DefinedNames.InvalidNamedRanges().ToList())
    {
        broken.Delete();
    }
}
```

:::note
Materialise with `ToList()` before deleting — removing items while enumerating the collection
will throw.
:::

## Scope in practice

Use **workbook** scope by default: one name, visible everywhere, and formulas on any sheet can
use it.

Use **worksheet** scope when the same label needs a different meaning per sheet — a per-sheet
`Total` on twelve monthly sheets, for instance. Each sheet's formulas resolve `Total` to its
own range, and anything outside qualifies it:

```csharp
foreach (var month in months)
{
    var sheet = workbook.Worksheets.Add(month);
    sheet.Range("B2:B40").AddToNamed("Total", XLScope.Worksheet);
    sheet.Cell("B41").FormulaA1 = "=SUM(Total)";   // resolves to this sheet's range
}

// From elsewhere, qualify it
summary.Cell("B2").FormulaA1 = "=SUM(January!Total)";
```

## A worked example

```csharp
using XLibur.Excel;

using var workbook = new XLWorkbook();

// Lookup data on a hidden sheet, exposed by name
var lookups = workbook.Worksheets.Add("Lookups");
string[] regions = ["North", "South", "East", "West"];
double[] rates = [0.20, 0.18, 0.22, 0.19];

for (var i = 0; i < regions.Length; i++)
{
    lookups.Cell(i + 1, 1).Value = regions[i];
    lookups.Cell(i + 1, 2).Value = rates[i];
}

lookups.Range(1, 1, regions.Length, 1).AddToNamed("Regions", XLScope.Workbook, "Valid sales regions");
lookups.Range(1, 1, regions.Length, 2).AddToNamed("RegionRates", XLScope.Workbook, "Region -> tax rate");
lookups.Hide();

// A named constant
workbook.DefinedNames.Add("ReportYear", "=2026");

// The data sheet uses the names rather than raw addresses
var ws = workbook.Worksheets.Add("Orders");
ws.Position = 1;

ws.Cell("A1").FormulaA1 = "=\"Orders for \" & ReportYear";
ws.Cell("A1").Style.Font.SetBold().Font.SetFontSize(14);

ws.Cell("A3").Value = "Region";
ws.Cell("B3").Value = "Net";
ws.Cell("C3").Value = "Tax";
ws.Range("A3:C3").Style.Font.Bold = true;

ws.Range("A4:A100").CreateDataValidation().List("=Regions");

for (var row = 4; row <= 100; row++)
{
    ws.Cell(row, 3).FormulaA1 = $"=IF(A{row}=\"\",\"\",B{row}*VLOOKUP(A{row},RegionRates,2,FALSE))";
}

ws.Range("B4:C100").AddToNamed("OrderAmounts");
ws.Range("B4:C100").Style.NumberFormat.Format = "$ #,##0.00";

ws.Cell("E3").Value = "Total net";
ws.Cell("F3").FormulaA1 = "=SUM(OrderAmounts)";

ws.Columns().AdjustToContents();
workbook.SaveAs("Orders.xlsx");
```

## Where to next

- [Cells and Ranges](./cells-and-ranges.md) — the ranges names point at
- [Formulas](./formulas.md) — using names in formulas
- [Data Validation](./data-validation.md) — named lists for dropdowns
