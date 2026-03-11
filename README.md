# About

This is a fork from the ClosedXML project at around v105.
For now the internal namespace is the same, but likely to change to avoid conflict with the origin project.

The project goal was a desire to get some much needed patches implemented that I use in my workflows, and the ClosedXML project was not accepting contributions.

This is a fork from the [ClosedXML](https://github.com/ClosedXML/ClosedXML/) project, taken from version v0.105.0 (May 15, 2025)
Namespaces are changed to avoid conflicts with the original project.


### Primary differences from ClosedXML (0.105)

- Dropped support for <net8
- Leverage later c# lang features.
- Fix some minor but relevant to us outstanding issues.

XLibur is a .NET library for reading, manipulating and writing Excel 2007+ (.xlsx, .xlsm) files. It aims to provide an
intuitive and user-friendly interface to dealing with the underlying [OpenXML](https://github.com/OfficeDev/Open-XML-SDK) API.


### Install ClosedXML via NuGet

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

## Developer guidelines

The [OpenXML specification](https://ecma-international.org/publications-and-standards/standards/ecma-376/) is a large
and complicated beast.

Feel free to submit a PR

Please read the [full developer guidelines](CONTRIBUTING.md).

## Credits
* Project originally created by Manuel de Leon
* Current maintainer of ClosedXML: [Jan Havlíček](https://github.com/jahav)
* Former maintainer and lead developer: [Francois Botha](https://github.com/igitur)
* Master of Computing Patterns: [Aleksei Pankratev](https://github.com/Pankraty)
* Logo design by [@Tobaloidee](https://github.com/Tobaloidee)
