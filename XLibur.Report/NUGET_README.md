# XLibur.Report

Report templating for [XLibur](https://github.com/XLibur/XLibur).

Author a report as an ordinary `.xlsx` template — placeholder expressions, named ranges and tag
markers — bind .NET data to it, and generate the finished workbook. Charts, pivot tables and
pictures survive range expansion.

```csharp
using var template = new XLTemplate("SalesReport.xlsx");
template.AddVariable("Company", "Contoso");
template.AddVariable("Items", sales);
var result = template.Generate();
template.SaveAs("SalesReport-2026.xlsx");
```

The template expression language is [Scriban](https://github.com/scriban/scriban), with XLibur's
Excel function library bridged in, so `{{ SUM(items.Price) }}` and `{{ ROUND(item.Price, 2) }}`
work inside `{{ }}` expressions.

To run templates written in ClosedXML.Report's C# expression syntax instead, install
`XLibur.Report.DynamicLinq` and pass its engine to the template.

Licensed under the MIT License. Portions adapted from
[ClosedXML.Report](https://github.com/ClosedXML/ClosedXML.Report) (MIT) — see `NOTICE`.
