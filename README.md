# XLibur

[![Build and Test](https://github.com/XLibur/XLibur/actions/workflows/build-and-test.yml/badge.svg)](https://github.com/XLibur/XLibur/actions/workflows/build-and-test.yml)
[![NuGet](https://img.shields.io/nuget/v/XLibur.svg)](https://www.nuget.org/packages/XLibur)
[![NuGet Downloads](https://img.shields.io/nuget/dt/XLibur.svg)](https://www.nuget.org/packages/XLibur)
[![SonarCloud Quality Gate](https://sonarcloud.io/api/project_badges/measure?project=XLibur_XLibur&metric=alert_status)](https://sonarcloud.io/dashboard?id=XLibur_XLibur)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## About

This is a fork from the [ClosedXML](https://github.com/ClosedXML/ClosedXML/) project, taken from version v0.105.0 (May 15, 2025)
Namespaces are changed to avoid conflicts with the original project.

This project goal was a desire to get some much-needed patches implemented that I use in my workflows, and the ClosedXML
project is at the time of writing not accepting community contributions. We've had some PR's up for a few months with no traction.

XLibur is a .NET library for reading, manipulating and writing Excel 2007+ (.xlsx, .xlsm) files. It aims to provide an
intuitive and user-friendly interface to dealing with the underlying [OpenXML](https://github.com/OfficeDev/Open-XML-SDK) API.

## Should I use this?

If you have no issues currently with ClosedXML, then this library likely gains you nothing.
If some of the differences introduced here are useful, then sure consider it!

### Primary differences from ClosedXML (0.105)

- Dropped support for <net8
- Enable nullability annotations.
- Leverage later c# lang features.
- Fix some outstanding bugs we wanted.
- Improve memory usage, especially with formatted cells.

### Release notes and migration guide

At present most of the surface area is the same as ClosedXML.
Import the Nuget.
Rename the namespace to `XLibur` and in most cases you should be ready to go.

### Install XLibur via NuGet

```
PM> Install-Package XLibur
```

### What can you do with this?

XLibur allows you to create Excel files without the Excel application. The typical example is creating Excel reports on
a web server.

**Example:**

```c#
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

The [OpenXML specification](https://ecma-international.org/publications-and-standards/standards/ecma-376/) is a large
and complicated beast.

Feel free to submit a PR

Please read the [full developer guidelines](CONTRIBUTING.md).

## Credits

* CloesdXML Project originally created by Manuel de Leon
* Maintainer of ClosedXML: [Jan Havlíček](https://github.com/jahav)
* Former maintainer and lead developer: [Francois Botha](https://github.com/igitur)
* Master of Computing Patterns: [Aleksei Pankratev](https://github.com/Pankraty)
* Logo design by [@Tobaloidee](https://github.com/Tobaloidee)
