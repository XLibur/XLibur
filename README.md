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

XLibur is a fork of [ClosedXML v0.105.0](https://github.com/ClosedXML/ClosedXML/)
(May 2025), created to ship patches and improvements that couldn't land upstream.
Namespaces are prefixed with `XLibur` to avoid conflicts with ClosedXML if both
are referenced in the same project.

## Should I use this?

**Stick with ClosedXML** if it meets your needs — XLibur gives you nothing extra.

**Consider XLibur if** you want any of the following improvements over ClosedXML 0.105:

- **Reduced memory usage and performance gains** — particularly for workbooks with many formatted cells
- **Nullability annotations** — full nullable reference type support throughout the API
- **Bug fixes** — several outstanding issues resolved that are pending upstream
- **No legacy .NET support** — .NET 8 and below are not supported

> ⚠️ XLibur has limited real-world production history. Use in critical systems at your own discretion.

## Migration from ClosedXML

The public API surface is largely unchanged from ClosedXML 0.105. To migrate:

1. Install the NuGet package (see below)
2. Replace `using ClosedXML` namespace references with `using XLibur`

### Install XLibur via NuGet

```sh
PM> Install-Package XLibur
```

Or via the .NET CLI:
```sh
dotnet add package XLibur
```

## User Guide

Nothing local. As the library is largely the same as ClosedXML, the [official documentation](https://closedxml.github.io/ClosedXML/) is still valid.


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
```

Run benchmarks (XLibur vs ClosedXML comparison):

```sh
# List available benchmarks
dotnet run -c Release --project XLibur.Benchmarks/XLibur.Benchmarks.csproj -- --list flat

# Run all benchmarks
dotnet run -c Release --project XLibur.Benchmarks/XLibur.Benchmarks.csproj -- --filter *

# Run a specific benchmark class
dotnet run -c Release --project XLibur.Benchmarks/XLibur.Benchmarks.csproj -- --filter '*XLiburWorkbookBenchmarks*'
dotnet run -c Release --project XLibur.Benchmarks/XLibur.Benchmarks.csproj -- --filter '*ClosedXmlWorkbookBenchmarks*'
```

## Developer guidelines

Please read the [full developer guidelines](CONTRIBUTING.md).

## Credits

* ClosedXML originally created by [Manuel de Leon](https://github.com/mdeleone)
* ClosedXML maintainer: [Jan Havlíček](https://github.com/jahav)
* Former ClosedXML maintainer and lead developer: [Francois Botha](https://github.com/igitur)
* Master of Computing Patterns: [Aleksei Pankratev](https://github.com/Pankraty)