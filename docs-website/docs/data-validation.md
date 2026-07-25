---
id: data-validation
title: Data Validation
sidebar_label: Data Validation
description: Restrict what users can type into cells — numbers, dates, times, text length, dropdown lists, and custom formulas.
---

# Data Validation

Data validation constrains what a user may enter into a cell, and shows an input hint or an
error message when they stray outside it. It is the standard way to make a generated workbook
behave like a form rather than a blank grid.

Create a validation with `CreateDataValidation()` on a cell or range, then apply a criterion:

```csharp
using XLibur.Excel;

var ws = workbook.Worksheet("Entry");

// Only decimals between 1 and 5
ws.Cell("A1").CreateDataValidation().Decimal.Between(1, 5);

// Only whole numbers of at least 0
ws.Range("B2:B100").CreateDataValidation().WholeNumber.EqualOrGreaterThan(0);
```

## Criteria types

A validation carries exactly one criterion type, chosen by the property you use:

| Property | Allows |
|---|---|
| `WholeNumber` | Integers |
| `Decimal` | Any number |
| `Date` | Dates |
| `Time` | Times |
| `TextLength` | Strings of a constrained length |
| `List(...)` | One of a fixed set of values |
| `Custom(...)` | Whatever a formula says |

Each of the first five supports the same comparisons: `EqualTo`, `NotEqualTo`, `GreaterThan`,
`LessThan`, `EqualOrGreaterThan`, `EqualOrLessThan`, `Between`, `NotBetween`.

### Numbers

```csharp
ws.Cell("A1").CreateDataValidation().WholeNumber.Between(1, 100);
ws.Cell("A2").CreateDataValidation().WholeNumber.EqualTo(2);
ws.Cell("A3").CreateDataValidation().Decimal.GreaterThan(0);
ws.Cell("A4").CreateDataValidation().Decimal.NotBetween(-1, 1);
```

Bounds can come from other cells, so the limits stay editable in the sheet:

```csharp
ws.Cell("H1").Value = 0;
ws.Cell("H2").Value = 1000;

ws.Range("B2:B100").CreateDataValidation()
    .Decimal.Between(ws.Cell("H1"), ws.Cell("H2"));
```

### Dates and times

```csharp
ws.Cell("C1").CreateDataValidation()
    .Date.EqualOrGreaterThan(new DateTime(2026, 1, 1));

ws.Range("C2:C100").CreateDataValidation()
    .Date.Between(new DateTime(2026, 1, 1), new DateTime(2026, 12, 31));

ws.Cell("D1").CreateDataValidation()
    .Time.Between(new TimeSpan(9, 0, 0), new TimeSpan(17, 0, 0));
```

### Text length

```csharp
ws.Range("E2:E100").CreateDataValidation().TextLength.Between(1, 50);
ws.Cell("E1").CreateDataValidation().TextLength.EqualOrLessThan(255);
```

## Dropdown lists

Three ways to supply the list.

**From a range on the sheet:**

```csharp
ws.Cell("H1").Value = "Yes";
ws.Cell("H2").Value = "No";
ws.Cell("H3").Value = "N/A";

ws.Range("F2:F100").CreateDataValidation().List(ws.Range("H1:H3"));
```

**From a named range** — the cleanest option when the list lives on another sheet, because
Excel's `List(range)` cannot cross sheets directly:

```csharp
var lookups = workbook.Worksheets.Add("Lookups");
lookups.Cell("A1").Value = "North";
lookups.Cell("A2").Value = "South";
lookups.Cell("A3").Value = "East";
lookups.Cell("A4").Value = "West";
lookups.Range("A1:A4").AddToNamed("Regions");
lookups.Hide();

ws.Range("G2:G100").CreateDataValidation().List("=Regions");
```

**Inline, as a literal list** — subject to Excel's 255-character limit on the whole string:

```csharp
ws.Range("A2:A100").CreateDataValidation().List("\"Low,Medium,High\"");
```

Hiding the dropdown arrow while keeping the constraint:

```csharp
ws.Range("F2:F100").CreateDataValidation().List(ws.Range("H1:H3"), inCellDropdown: false);
```

## Custom formula validation

`Custom` takes a formula that must evaluate to `TRUE` for the entry to be accepted. It is
written relative to the top-left cell of the validated range:

```csharp
// Must be a multiple of 5
ws.Range("B2:B100").CreateDataValidation().Custom("=MOD(B2,5)=0");

// Must not duplicate an existing entry in the column
ws.Range("A2:A100").CreateDataValidation().Custom("=COUNTIF($A$2:$A$100,A2)=1");

// End date must be on or after start date
ws.Range("D2:D100").CreateDataValidation().Custom("=D2>=C2");

// Must be uppercase
ws.Range("E2:E100").CreateDataValidation().Custom("=EXACT(E2,UPPER(E2))");
```

## Messages

Two separate messages: an *input* hint shown when the cell is selected, and an *error* shown
when the entry fails validation.

```csharp
var validation = ws.Range("B2:B100").CreateDataValidation();
validation.Decimal.Between(0, 1_000_000);

validation.InputTitle = "Order value";
validation.InputMessage = "Enter the value in GBP, excluding VAT.";

validation.ErrorTitle = "Value out of range";
validation.ErrorMessage = "The order value must be between 0 and 1,000,000.";
validation.ErrorStyle = XLErrorStyle.Stop;
```

`ErrorStyle` controls how hard the rejection is:

| Style | Behaviour |
|---|---|
| `Stop` | Entry is rejected outright (default) |
| `Warning` | User is warned but may proceed |
| `Information` | Purely informational; the entry is accepted |

Suppressing either message:

```csharp
validation.ShowInputMessage = false;
validation.ShowErrorMessage = false;
```

## Blanks

By default an empty cell passes validation. Set `IgnoreBlanks = false` to require a value:

```csharp
var validation = ws.Range("A2:A100").CreateDataValidation();
validation.WholeNumber.GreaterThan(0);
validation.IgnoreBlanks = false;
```

## Applying to several ranges at once

```csharp
var ranges = ws.Ranges("A1:B2,B4:D7,F4:G5");

var validation = ranges.CreateDataValidation();
validation.Decimal.EqualOrGreaterThan(0);
validation.IgnoreBlanks = false;
```

:::note
Validations that overlap are resolved by the later one winning for the shared cells. Adding a
validation to `B3:B4` after one on `B1:B4` leaves `B1:B2` on the first rule and `B3:B4` on the
second.
:::

## Inspecting and removing

```csharp
foreach (var validation in ws.DataValidations)
{
    var addresses = string.Join(", ", validation.Ranges.Select(r => r.RangeAddress.ToString()));
    Console.WriteLine($"{addresses}: {validation.AllowedValues} {validation.Value}");
}

// Rules covering a particular range
foreach (var validation in ws.DataValidations.GetAllInRange(ws.Range("A2:A100").RangeAddress))
{
    Console.WriteLine(validation.AllowedValues);
}

// Remove by predicate
ws.DataValidations.Delete(v => v.AllowedValues == XLAllowedValues.List);

// Or clear the rules on a range
ws.Range("A2:A100").Clear(XLClearOptions.DataValidation);
```

## A worked example

```csharp
using XLibur.Excel;

using var workbook = new XLWorkbook();

// A hidden sheet holding the lookup lists
var lookups = workbook.Worksheets.Add("Lookups");
string[] regions = ["North", "South", "East", "West"];
for (var i = 0; i < regions.Length; i++)
{
    lookups.Cell(i + 1, 1).Value = regions[i];
}

lookups.Range(1, 1, regions.Length, 1).AddToNamed("Regions");
lookups.Hide();

// The entry form
var ws = workbook.Worksheets.Add("Order Entry");
ws.Position = 1;

ws.Cell("A1").Value = "Customer";
ws.Cell("B1").Value = "Region";
ws.Cell("C1").Value = "Order date";
ws.Cell("D1").Value = "Quantity";
ws.Cell("E1").Value = "Unit price";
ws.Range("A1:E1").Style.Font.Bold = true;

// Customer: 1–100 characters, required
var customer = ws.Range("A2:A200").CreateDataValidation();
customer.TextLength.Between(1, 100);
customer.IgnoreBlanks = false;
customer.InputTitle = "Customer";
customer.InputMessage = "Enter the customer's registered name.";

// Region: dropdown from the named range
var region = ws.Range("B2:B200").CreateDataValidation();
region.List("=Regions");
region.ErrorTitle = "Unknown region";
region.ErrorMessage = "Pick one of the four sales regions.";

// Order date: this year only
var date = ws.Range("C2:C200").CreateDataValidation();
date.Date.Between(new DateTime(2026, 1, 1), new DateTime(2026, 12, 31));
date.ErrorStyle = XLErrorStyle.Warning;
date.ErrorTitle = "Date outside 2026";
date.ErrorMessage = "Dates outside 2026 need approval — continue only if that is intended.";

// Quantity: positive whole numbers
ws.Range("D2:D200").CreateDataValidation().WholeNumber.GreaterThan(0);

// Unit price: non-negative
ws.Range("E2:E200").CreateDataValidation().Decimal.EqualOrGreaterThan(0);

ws.Range("C2:C200").Style.DateFormat.Format = "yyyy-MM-dd";
ws.Range("E2:E200").Style.NumberFormat.Format = "$ #,##0.00";
ws.Columns().AdjustToContents();

workbook.SaveAs("OrderEntry.xlsx");
```

## Where to next

- [Conditional Formatting](./conditional-formatting.md) — flag bad data rather than block it
- [Worksheets](./worksheets.md) — hiding the sheets that hold lookup lists
