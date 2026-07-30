# XLibur.Report.DynamicLinq

ClosedXML.Report expression-syntax compatibility for
[XLibur.Report](https://www.nuget.org/packages/XLibur.Report).

Install this **only** if you have templates written for ClosedXML.Report, with C# expressions inside
`{{ }}`. XLibur.Report's own engine needs nothing from here.

```csharp
using XLibur.Report;
using XLibur.Report.DynamicLinq;

using var template = new XLTemplate("LegacyReport.xlsx", new DynamicLinqExpressionEngine());
template.AddVariable("Items", sales);
template.Generate();
template.SaveAs("Report-2026.xlsx");
```

A template's structure — defined-name-bound ranges, the options row, `<<tags>>`, the `&=` formula
prefix — means the same thing whichever engine evaluates its expressions. The expression language is
the only difference, so pointing an upstream template at this engine is usually all that is needed.

Supported: property and method access (`item.Name.ToUpper()`), arithmetic and comparison, the
conditional operator, and LINQ over collections in scope
(`items.Where(x => x.Qty > 10).Sum(x => x.Total)`). Inside a bound range the names are `item`, `index`
and `items`, and workbook variables are reachable by name or with an `@` prefix.

Not supported: the Excel-function bridge (`{{ SUM(...) }}`). That is a feature of XLibur.Report's
default engine; upstream syntax never had it, and templates for this engine call .NET methods instead.

## Trusted templates only

[Dynamic LINQ](https://dynamic-linq.net) has no sandbox. An expression it parses can reach the methods
and properties of any object in scope, and the library's own history includes CVE-2023-32571 —
arbitrary method invocation. **Do not point this engine at a template a user uploaded.** For that,
use XLibur.Report's default Scriban engine, which has real execution limits and no reflection escape.

Licensed under the MIT License. Portions derived from
[ClosedXML.Report](https://github.com/ClosedXML/ClosedXML.Report) (MIT) — see `NOTICE`.
