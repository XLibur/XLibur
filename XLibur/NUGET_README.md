# XLibur

<img src="https://raw.githubusercontent.com/XLibur/XLibur/main/resources/logo/nuget-logo.png" alt="XLibur logo" width="128" />

[![NuGet](https://img.shields.io/nuget/v/XLibur.svg)](https://www.nuget.org/packages/XLibur)
[![NuGet Downloads](https://img.shields.io/nuget/dt/XLibur.svg)](https://www.nuget.org/packages/XLibur)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/XLibur/XLibur/blob/main/LICENSE)

## About

XLibur is a .NET library for reading, manipulating and writing Excel 2007+ (.xlsx, .xlsm) files. It aims to provide an
intuitive and user-friendly interface to dealing with the underlying [OpenXML](https://github.com/OfficeDev/Open-XML-SDK) API.

This is a fork from the [ClosedXML](https://github.com/ClosedXML/ClosedXML/) project, taken from version v0.105.0 (May 15, 2025).
Namespaces are changed to avoid conflicts with the original project.

### Primary differences from ClosedXML (0.105)

- Dropped support for <net8
- Enable nullability annotations.
- Leverage later C# lang features.
- Fix some outstanding bugs we wanted.
- Improve memory usage, especially with formatted cells.

### Migration from ClosedXML

At present most of the surface area is the same as ClosedXML.
Import the NuGet package, rename the namespace to `XLibur`, and in most cases you should be ready to go.

### Install

```
dotnet add package XLibur
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

## Documentation

For full documentation, source code, and contribution guidelines, visit the [GitHub repository](https://github.com/XLibur/XLibur).

## Credits

* ClosedXML Project originally created by Manuel de Leon
* Maintainer of ClosedXML: [Jan Havlíček](https://github.com/jahav)
* Former maintainer and lead developer: [Francois Botha](https://github.com/igitur)
* Master of Computing Patterns: [Aleksei Pankratev](https://github.com/Pankraty)
* Logo design by [@Tobaloidee](https://github.com/Tobaloidee)
