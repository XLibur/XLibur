# XLibur

<img src="resources/logo/logo.png" alt="XLibur logo" width="512" />

[![Build and Test](https://github.com/XLibur/XLibur/actions/workflows/build-and-test.yml/badge.svg)](https://github.com/XLibur/XLibur/actions/workflows/build-and-test.yml)
[![NuGet](https://img.shields.io/nuget/v/XLibur.svg)](https://www.nuget.org/packages/XLibur)
[![NuGet Downloads](https://img.shields.io/nuget/dt/XLibur.svg)](https://www.nuget.org/packages/XLibur)
[![SonarCloud Quality Gate](https://sonarcloud.io/api/project_badges/measure?project=XLibur_XLibur&metric=alert_status)](https://sonarcloud.io/dashboard?id=XLibur_XLibur)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## About

XLibur is a .NET 8+ library for reading, manipulating, and writing Excel 2007+
(.xlsx, .xlsm) files. It provides an intuitive interface over the underlying
[OpenXML](https://github.com/OfficeDev/Open-XML-SDK) API.

XLibur forked from [ClosedXML v0.105.0](https://github.com/ClosedXML/ClosedXML/)
(May 2025), created to ship patches and improvements that didn't land upstream.
Namespaces are prefixed with `XLibur`.

## Should I use this?

**Stick with ClosedXML** if:
- A library with developers who have had experience with the library over many years.
- A longer term focus on the product.
- You need support for dotnet <8 (net standard, or NET472)

**Consider XLibur if** you want any of the following changes over ClosedXML 0.105:

- **Reduced memory usage and performance gains** - particularly for workbooks with many formatted cells
- **Bug fixes** — several outstanding community issues resolved that are pending upstream
- **Community PR/enhancements** - several community contriubtions or requests have been merged in to this codebase.

## Migration from ClosedXML

The public API surface is largely unchanged from ClosedXML 0.105. To migrate:

1. Install the NuGet package (see below)
2. Replace `using ClosedXML` namespace references with `using XLibur`

### Install XLibur via NuGet

The recommended package is **`XLibur.Bundle`**, which installs the core library together with the
default font engine and behaves like ClosedXML out of the box:

```sh
PM> Install-Package XLibur.Bundle
```

Or via the .NET CLI:
```sh
dotnet add package XLibur.Bundle
```

### Font engine configuration (different from ClosedXML)

ClosedXML bundles [SixLabors.Fonts](https://github.com/SixLabors/Fonts) directly into its core assembly for text
measurement (column auto-fit, row heights, glyph metrics). XLibur instead keeps the **core assembly
free of any font library** and ships the font engine as a **separate, swappable package**. This lets
you pick a font library with a license that suits you and avoids forcing a font dependency on library
authors who don't need one.

What this means when migrating:

- **Install `XLibur.Bundle` (or `XLibur` + `XLibur.Fonts.SkiaSharp`) and no code changes are needed.**
  The default [SkiaSharp](https://github.com/mono/SkiaSharp) engine (MIT-licensed) is **auto-registered
  by XLibur core the first time you create a workbook** — there is no startup call to add:

  ```csharp
  using var wb = new XLWorkbook(); // font engine resolved automatically
  ```

  The default resolves system fonts and falls back to an embedded, metric-only Calibri-compatible font,
  so text measurement works even in headless/serverless environments with no system fonts installed.

- **If you install the bare `XLibur` package with no font engine**, creating a workbook throws an
  `InvalidOperationException` telling you to add a font engine package. This is intentional — it's how
  the core stays font-library-agnostic.

- **To choose a different engine**, install its package and either register it at startup (it takes
  precedence over the auto-registered default) or pass it per workbook:

  | Package | Font library | License | Notes |
  |---|---|---|---|
  | `XLibur.Fonts.SkiaSharp` | SkiaSharp | MIT | **Default.** Auto-registers; ships native binaries. |
  | `XLibur.Fonts.SixLabors.V1` | SixLabors.Fonts 1.x | Apache 2.0 | Pure-managed; matches ClosedXML 0.105's engine. |
  | `XLibur.Fonts.SixLabors` | SixLabors.Fonts 2.x | Six Labors Split License | Commercial restrictions over $1M revenue. |

  ```csharp
  // Override globally at startup (e.g. keep ClosedXML's SixLabors 1.x behavior):
  SixLaborsV1FontBootstrap.Register();

  // Or override per workbook:
  var options = new LoadOptions { FontEngine = new SkiaSharpFontEngine("Arial") };
  using var wb = new XLWorkbook(options);
  ```

Resolution order for the font engine is: `LoadOptions.FontEngine` (per workbook) →
`LoadOptions.DefaultFontEngine` (explicitly registered global) → the auto-registered default engine.
See [docs/font-architecture.md](docs/font-architecture.md) for the full design.

## User Guide

XLibur documentation lives at [xlibur.github.io/XLibur](https://xlibur.github.io/XLibur/) — start with
the [Getting Started](https://xlibur.github.io/XLibur/getting-started) guide.

## Report templating

`XLibur.Report` generates reports from `.xlsx` templates: author the report in Excel — placeholder
expressions, a defined name over the rows to repeat, tag markers for totals and grouping — bind .NET
data to it, and generate the finished workbook. Charts, pivot tables and pictures survive range
expansion, which is the part comparable libraries do not do.

```csharp
using var template = new XLTemplate("SalesReport.xlsx");
template.AddVariable("Company", "Contoso");
template.AddVariable("Sales", sales);
template.Generate();
template.SaveAs("SalesReport-2026.xlsx");
```

See [docs/report-templating.md](docs/report-templating.md) for the template language and the tag
reference, and [XLibur.Report.Examples](XLibur.Report.Examples/README.md) for ten worked examples that
each write both their template and their generated report.

## Usage

XLibur lets you create and manipulate Excel files without Excel installed — a common use case is generating reports on a web server.
```csharp
using (var workbook = new XLWorkbook())
{
    var worksheet = workbook.Worksheets.Add("Sample Sheet");
    worksheet.Cell("A1").Value = "Hello World!";
    worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
    workbook.SaveAs("HelloWorld.xlsx");
}
```

## Building, Testing, and Benchmarks

Build the solution:

```sh
dotnet build XLibur.slnx
```

Run the test suite:

```sh
dotnet test XLibur.Tests/XLibur.Tests.csproj
dotnet test XLibur.Report.Tests/XLibur.Report.Tests.csproj
```

Published benchmark results are available at
[jafin.github.io/XLBench](https://jafin.github.io/XLBench/charts.html):

Benchmark snapshot 2026-07-20:

[![XLibur benchmark results](docs/benchmark_snapshot.jpg)](https://jafin.github.io/XLBench/charts.html)

[![XLibur benchmark memory results](docs/benchmark_snapshot_memory.jpg)](https://jafin.github.io/XLBench/charts.html)

## Developer guidelines

Please read the [full developer guidelines](CONTRIBUTING.md).

## Credits

* ClosedXML authors [Manuel de Leon](https://github.com/mdeleone), [Jan Havlíček](https://github.com/jahav), [Francois Botha](https://github.com/igitur), [Aleksei Pankratev](https://github.com/Pankraty)
