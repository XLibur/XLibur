# XLibur

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
- Add a streaming write API for exports too large to hold in memory (see below).

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

### Writing very large files

`XLWorkbook` builds the whole workbook in memory before saving, which puts a ceiling on how large
an export can be. For those, `XLStreamingWorkbook` writes rows straight into the file as you append
them, so memory stays flat no matter how many rows there are — a million rows by ten columns costs
about 108 MB, or 14 MB with `Inline` string storage.

```c#
using XLibur.Excel.Streaming;

using var workbook = XLStreamingWorkbook.Create("Large.xlsx");

var sheet = workbook.AddWorksheet("Data");
sheet.Column(1).Width = 30;
sheet.FreezeRows(1);

var header = workbook.CreateStyle();
header.Font.Bold = true;
sheet.AppendRow(["Name", "Amount"], header);

for (var i = 0; i < 1_000_000; i++)
    sheet.AppendRow($"Item {i}", i * 1.5);

workbook.Finish();   // required: writes the strings, styles and workbook parts
```

The trade is that it is append-only: rows go in ascending order, one worksheet at a time, nothing
can be read back or revised, and formulas are stored verbatim rather than evaluated. Use
`XLWorkbook` whenever any of that matters.

Two notes worth knowing. `Finish()` must be called — disposing without it abandons the write.
And by default distinct strings accumulate in a shared string table until then, so if your data has
an unbounded number of distinct strings, set
`StringStorage = XLStreamingStringStorage.Inline` to keep memory flat at the cost of a larger file.

## Documentation

For full documentation, source code, and contribution guidelines, visit the [GitHub repository](https://github.com/XLibur/XLibur).

## Credits

* ClosedXML Project originally created by Manuel de Leon
* Maintainer of ClosedXML: [Jan Havlíček](https://github.com/jahav)
* Former maintainer and lead developer: [Francois Botha](https://github.com/igitur)
* Master of Computing Patterns: [Aleksei Pankratev](https://github.com/Pankraty)
* Logo design by [@Tobaloidee](https://github.com/Tobaloidee)
