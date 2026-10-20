# XLibur

## About

XLibur is a .NET library for reading, manipulating and writing Excel 2007+ (.xlsx, .xlsm) files. It aims to provide an
intuitive and user-friendly interface to dealing with the underlying [OpenXML](https://github.com/OfficeDev/Open-XML-SDK) API.

## Install

The recommended package is **`XLibur.Bundle`**, which installs the core library together with the
default font engine and behaves like ClosedXML out of the box:

```sh
dotnet add package XLibur.Bundle
```

Or via the Package Manager console:

```sh
PM> Install-Package XLibur.Bundle
```

## Quick start

XLibur lets you create and manipulate Excel files without Excel installed — a common use case is
generating reports on a web server.

```csharp
using (var workbook = new XLWorkbook())
{
    var worksheet = workbook.Worksheets.Add("Sample Sheet");
    worksheet.Cell("A1").Value = "Hello World!";
    worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
    workbook.SaveAs("HelloWorld.xlsx");
}
```

More in the [Getting Started guide](https://xlibur.github.io/XLibur/getting-started).

## Ancestry

XLibur forked [ClosedXML v0.105](https://github.com/ClosedXML/ClosedXML/). Namespaces are prefixed with `XLibur`. Surface API is *mostly* compatible, but check the docs for details.

## Migration from ClosedXML

The public API surface is largely unchanged from ClosedXML 0.105. To migrate:

1. Install `XLibur.Bundle` (see [Install](#install))
2. Replace `using ClosedXML` namespace references with `using XLibur`

**One behavioural difference: font engines.** ClosedXML bundles
[SixLabors.Fonts](https://github.com/SixLabors/Fonts) into its core assembly for text measurement
(column auto-fit, row heights, glyph metrics). XLibur keeps the core assembly free of any font
library and ships the font engine as a separate, swappable package, so you can pick one whose
license suits you.

With `XLibur.Bundle`, no code changes are needed — the [SkiaSharp](https://github.com/mono/SkiaSharp)
engine (MIT) auto-registers the first time you create a workbook, and falls back to an embedded
metric-only Calibri-compatible font so measurement works in headless environments. Installing the
bare `XLibur` package with no font engine throws an `InvalidOperationException` on workbook creation
telling you which package to add.

To use a different engine — including SixLabors 1.x, which matches ClosedXML 0.105's behaviour
exactly — see [docs/font-architecture.md](https://github.com/XLibur/XLibur/blob/main/docs/font-architecture.md) for the available packages,
their licenses, and the engine resolution order.

## Report templating

`XLibur.Report` generates reports from `.xlsx` templates: author the report in Excel, bind .NET data
to it, and generate the finished workbook. Charts, pivot tables and pictures survive range expansion,
which is the part comparable libraries do not do.

```csharp
using var template = new XLTemplate("SalesReport.xlsx");
template.AddVariable("Company", "Contoso");
template.AddVariable("Sales", sales);
template.Generate();
template.SaveAs("SalesReport-2026.xlsx");
```

## License

[MIT LICENSE](https://github.com/XLibur/XLibur/blob/main/LICENSE).

## Credits

ClosedXML authors who developed the core code we build on.
[Manuel de Leon](https://github.com/mdeleone),
[Jan Havlíček](https://github.com/jahav), [Francois Botha](https://github.com/igitur),
[Aleksei Pankratev](https://github.com/Pankraty).
