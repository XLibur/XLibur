---
id: workbook-settings
title: Workbook Settings
sidebar_label: Workbook Settings
description: Document properties, custom properties, protection, calculation mode, load and save options, and other workbook-level configuration.
---

# Workbook Settings

Settings that apply to the file as a whole rather than to any one sheet: the metadata Windows
shows in a file's Properties dialog, protection, how Excel recalculates, and the options that
govern loading and saving.

## Document properties

`workbook.Properties` maps to the fields Excel exposes under File → Info:

```csharp
using XLibur.Excel;

using var workbook = new XLWorkbook();

workbook.Properties.Author = "Reporting Service";
workbook.Properties.Title = "Q1 Sales Report";
workbook.Properties.Subject = "Regional performance";
workbook.Properties.Category = "Finance";
workbook.Properties.Keywords = "sales;q1;regional";
workbook.Properties.Comments = "Generated automatically — do not edit by hand.";
workbook.Properties.Status = "Final";
workbook.Properties.Company = "Example Ltd";
workbook.Properties.Manager = "A. Manager";
workbook.Properties.LastModifiedBy = "Reporting Service";
workbook.Properties.Created = new DateTime(2026, 1, 21);
workbook.Properties.Modified = DateTime.UtcNow;
```

`workbook.Author` is a shortcut that also seeds the author of new
[comments](./comments-and-hyperlinks.md#author-and-signature):

```csharp
using var workbook = new XLWorkbook { Author = "Reporting Service" };
```

Reading them from an existing file is a cheap way to audit where a spreadsheet came from:

```csharp
using var workbook = new XLWorkbook("Report.xlsx");

Console.WriteLine(workbook.Properties.Author);
Console.WriteLine(workbook.Properties.Created);
Console.WriteLine(workbook.Properties.LastModifiedBy);
```

## Custom properties

Arbitrary typed key/value pairs stored in the file — useful for stamping a generated workbook
with the version, run id, or source system that produced it:

```csharp
workbook.CustomProperties.Add("GeneratorVersion", "2.4.1");
workbook.CustomProperties.Add("RunId", 48213);
workbook.CustomProperties.Add("GeneratedAt", DateTime.UtcNow);
workbook.CustomProperties.Add("IsDraft", false);
```

Four types are supported — `Text`, `Number`, `Date`, and `Boolean` — inferred from the value.

Reading them back:

```csharp
foreach (var property in workbook.CustomProperties)
{
    Console.WriteLine($"{property.Name} ({property.Type}) = {property.Value}");
}

var runId = workbook.CustomProperties.CustomProperty("RunId").GetValue<int>();

workbook.CustomProperties.Delete("IsDraft");
```

This is a good place for a provenance stamp, because unlike a cell it cannot be accidentally
deleted while editing:

```csharp
static void StampProvenance(XLWorkbook workbook, string source)
{
    workbook.CustomProperties.Add("SourceSystem", source);
    workbook.CustomProperties.Add("GeneratedAtUtc", DateTime.UtcNow);
    workbook.CustomProperties.Add("GeneratorVersion",
        typeof(Program).Assembly.GetName().Version?.ToString() ?? "unknown");
}
```

## Calculation

```csharp
workbook.CalculateMode = XLCalculateMode.Auto;         // Excel recalculates on open and edit
workbook.CalculateMode = XLCalculateMode.Manual;       // only on F9
workbook.CalculateMode = XLCalculateMode.AutoNoTable;  // auto, except data tables
```

`ReferenceStyle` controls whether Excel shows `A1` or `R1C1` addresses in its UI — it does not
change how you write formulas in code:

```csharp
workbook.ReferenceStyle = XLReferenceStyle.A1;     // or R1C1, Default
```

Recalculating with XLibur's own engine, rather than deferring to Excel:

```csharp
workbook.RecalculateAllFormulas();
workbook.Worksheet("Data").RecalculateAllFormulas();
```

See [Formulas](./formulas.md#evaluating-formulas) for when this matters.

## Date system

Excel workbooks store dates as a serial number counting from an epoch. Windows Excel uses
1900; legacy Mac files use 1904. Switching moves every date in the workbook by four years, so
this is a load-time property rather than something to toggle mid-build:

```csharp
workbook.Use1904DateSystem = true;
workbook.SetUse1904DateSystem();

Console.WriteLine(workbook.Use1904DateSystem);
```

Leave it alone unless you are reading a file that already uses the 1904 system.

## Right-to-left

```csharp
workbook.RightToLeft = true;                          // default for new sheets
workbook.Worksheet("Data").RightToLeft = true;        // per sheet
```

## Default style

`workbook.Style` is the style every new cell inherits. Setting it once is far cheaper than
styling cells individually:

```csharp
workbook.Style.Font.FontName = "Calibri";
workbook.Style.Font.FontSize = 11;
workbook.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
```

It is also the starting point for building a reusable style — see
[Styling](./styling.md#reusing-a-style).

## Protection

### Workbook structure

Workbook protection stops the *structure* changing — sheets being added, deleted, renamed, or
reordered. It does not protect cell contents; that is
[sheet protection](./worksheets.md#protecting-a-sheet).

```csharp
workbook.Protect("s3cret");

// Allow window moves but lock the structure
workbook.Protect("s3cret", XLProtectionAlgorithm.Algorithm.SHA512,
    XLWorkbookProtectionElements.Windows);

Console.WriteLine(workbook.IsProtected);
Console.WriteLine(workbook.IsPasswordProtected);

workbook.Unprotect("s3cret");
```

Elements: `Structure`, `Windows`, `Everything`, `None`.

Two hashing algorithms are available: `SimpleHash` (the default, matching Excel's legacy
scheme) and `SHA512`. Pass `SHA512` explicitly for new files unless you need the old format:

```csharp
workbook.Protect("s3cret", XLProtectionAlgorithm.Algorithm.SHA512);
```

:::warning
Neither workbook nor sheet protection encrypts anything. The data is stored in plain text
inside the `.xlsx` and any tool — including XLibur — can read it without the password. Treat it
as a guard against accidental edits, not as security. To genuinely protect contents, encrypt
the file at rest or restrict access to it.
:::

### Read-only recommendation

A softer signal: Excel prompts the user to open the file read-only.

```csharp
workbook.FileSharing.ReadOnlyRecommended = true;
workbook.FileSharing.UserName = "Reporting Service";
```

## Load options

`LoadOptions` configures a workbook as it is constructed — for both new and loaded files:

```csharp
var options = new LoadOptions
{
    RecalculateAllFormulas = true,          // recalculate during load; default false
    Dpi = new Point(120, 120),              // affects text measurement and image sizing
    FontEngine = new SkiaSharpFontEngine("Arial"),
};

using var workbook = new XLWorkbook("Report.xlsx", options);
using var fresh = new XLWorkbook(options);
```

| Option | Effect |
|---|---|
| `RecalculateAllFormulas` | Re-evaluate every formula on load rather than trusting cached values |
| `Dpi` | Resolution assumed for text measurement and images; default 96×96 |
| `FontEngine` | Per-workbook font engine — see [Fonts](./fonts.md) |
| `GraphicEngine` | Per-workbook image handling engine |

Two static members set global defaults for every workbook that does not specify its own:

```csharp
LoadOptions.DefaultFontEngine = new SkiaSharpFontEngine("Arial");
LoadOptions.DefaultGraphicEngine = customEngine;
```

## Save options

`SaveOptions` controls what happens on the way out:

```csharp
var options = new SaveOptions
{
    EvaluateFormulasBeforeSaving = true,
    ValidatePackage = true,
    ConsolidateConditionalFormatRanges = true,
    ConsolidateDataValidationRanges = true,
    GenerateCalculationChain = true,
    FilterPrivacy = true,
};

workbook.SaveAs("Report.xlsx", options);
```

| Option | Default | Effect |
|---|---|---|
| `EvaluateFormulasBeforeSaving` | `false` | Compute formula results and store them alongside the formulas |
| `ValidatePackage` | `false` | Run OpenXML schema validation before writing — slow, but catches malformed output |
| `ConsolidateConditionalFormatRanges` | `true` | Merge adjacent conditional-format ranges to shrink the file |
| `ConsolidateDataValidationRanges` | `true` | Same, for data validation |
| `GenerateCalculationChain` | `true` | Write the calc chain part Excel uses to order recalculation |
| `FilterPrivacy` | `null` | Set the privacy flag; `null` leaves it unchanged |

The shorthand overloads cover the two common cases:

```csharp
workbook.SaveAs("Report.xlsx");                                   // fast path
workbook.SaveAs("Report.xlsx", validate: true, evaluateFormulae: true);
workbook.Save();                                                   // back to where it was loaded from
workbook.Save(validate: false, evaluateFormulae: true);
```

:::note
`ValidatePackage` is worth switching on in tests and leaving off in production — it walks the
whole package against the OpenXML schema, which is expensive on large files.
:::

## Disposal

`XLWorkbook` implements `IDisposable`. Always dispose it, and prefer a `using` declaration:

```csharp
using var workbook = new XLWorkbook("Report.xlsx");
// ...
```

Long-lived workbooks held in a field or a cache keep their whole cell model in memory. For a
service generating files per request, construct, save, and dispose within the request.

## A worked example

```csharp
using System.Drawing;
using XLibur.Excel;

var loadOptions = new LoadOptions
{
    RecalculateAllFormulas = false,
    Dpi = new Point(96, 96),
};

using var workbook = new XLWorkbook(loadOptions)
{
    Author = "Reporting Service",
};

// Document metadata
workbook.Properties.Title = "Q1 Sales Report";
workbook.Properties.Subject = "Regional performance";
workbook.Properties.Company = "Example Ltd";
workbook.Properties.Category = "Finance";
workbook.Properties.Status = "Final";
workbook.Properties.Created = DateTime.UtcNow;

// Provenance, so support can tell where a file came from
workbook.CustomProperties.Add("SourceSystem", "sales-api");
workbook.CustomProperties.Add("GeneratedAtUtc", DateTime.UtcNow);
workbook.CustomProperties.Add("RunId", 48213);

// House style
workbook.Style.Font.FontName = "Calibri";
workbook.Style.Font.FontSize = 11;

var ws = workbook.Worksheets.Add("Summary");
ws.Cell("A1").Value = "Q1 Sales Report";
ws.Cell("A1").Style.Font.SetBold().Font.SetFontSize(16);
ws.Cell("A3").FormulaA1 = "=TODAY()";
ws.Cell("A3").Style.DateFormat.Format = "yyyy-MM-dd";

// Recalculate automatically when opened
workbook.CalculateMode = XLCalculateMode.Auto;

// Lock the structure so the sheet layout survives contact with users,
// and hint that the file is not meant to be edited
workbook.Protect("s3cret", XLProtectionAlgorithm.Algorithm.SHA512,
    XLWorkbookProtectionElements.Structure);
workbook.FileSharing.ReadOnlyRecommended = true;

workbook.SaveAs("Q1Sales.xlsx", new SaveOptions
{
    EvaluateFormulasBeforeSaving = true,
    ValidatePackage = false,
});
```

## Where to next

- [Worksheets](./worksheets.md) — sheet-level protection and view settings
- [Formulas](./formulas.md) — calculation mode and evaluation in detail
- [Fonts](./fonts.md) — the font engine `LoadOptions` selects
