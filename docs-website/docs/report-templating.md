---
id: report-templating
title: Report Templating
sidebar_label: Overview
description: Generate Excel reports from .xlsx templates with XLibur.Report — placeholder expressions, defined names that repeat rows, and tag markers for totals and grouping.
---

# Report Templating

`XLibur.Report` builds a report from a template that is itself an ordinary `.xlsx` file. You
author the report in Excel — the fonts, the number formats, the borders, the chart, the pivot
table — and mark the parts that come from data with `{{ }}` expressions and `<<Tag>>` markers.
At run time you bind .NET objects to it and generate the finished workbook.

The point is that none of the report's *appearance* is described in code. Whoever knows what
the report should look like opens Excel and makes it look like that.

```csharp
using XLibur.Report;

using var template = new XLTemplate("SalesReport.xlsx");
template.AddVariable("Company", "Contoso Ltd");
template.AddVariable("Sales", sales);

template.Generate();
template.SaveAs("SalesReport-2026.xlsx");
```

## Installation

```sh
dotnet add package XLibur.Report
```

The package depends on `XLibur`. If you use tags that size columns or rows to their contents
(`<<ColsFit>>`, `<<RowsFit>>`) you also need a font engine registered — see
[Getting Started](./getting-started.md#installation), or install `XLibur.Bundle` alongside it.

To run templates written for **ClosedXML.Report**'s C# expression syntax, add
`XLibur.Report.DynamicLinq` as well and pass its engine. See
[Coming from ClosedXML.Report](./report-migration.md).

## The three things a template does

Everything else in these pages is these three used harder.

### 1. An expression in a cell

`{{ … }}` anywhere in a cell's text is evaluated and replaced:

|  | A |
|---|---|
| **1** | `{{ Company }} — annual sales` |

With `AddVariable("Company", "Contoso Ltd")` bound, that cell reads `Contoso Ltd — annual sales`
in the generated workbook.

The language is [Scriban](https://github.com/scriban/scriban), and property names keep their C#
spelling — `{{ item.UnitPrice }}`, not `unit_price`. See [Expressions](./report-expressions.md).

### 2. A defined name that repeats rows

A **defined name** whose name matches a bound variable holding a *collection* makes the rows it
covers repeat, once per item. Inside those rows, `item` is the current one.

|  | A | B | C |
|---|---|---|---|
| **3** | Product | Quantity | Line total |
| **4** | `{{ item.Product }}` | `{{ item.Quantity }}` | `=B4*1.2` |
| **5** |  | `<<Sum>>` | `<<Sum>>` |

```csharp
// In the template, authored in Excel or in code:
workbook.DefinedNames.Add("Sales", sheet.Range("A4:C5"));
```

Row 3 is a heading and sits *outside* the name, so it is written once. Rows 4–5 are the name, so
they are what repeats.

### 3. The options row

The last row of a bound range is its **options row**. It carries the tags and the totals, it is
**not** repeated, and it is deleted if nothing was written into it.

In the table above, row 5 is the options row: `<<Sum>>` writes a `SUBTOTAL` there over every
generated row, so it survives. Had it been empty it would have gone, and the report would end at
the last data row.

:::note
A range **one row deep** has no options row, and repeats entirely. That is the right shape when
there is nothing to total.
:::

## A complete example

```csharp
using XLibur.Excel;
using XLibur.Report;

public sealed record Sale(string Product, string Region, int Quantity, decimal UnitPrice);

// --- Author the template (normally you would do this in Excel and ship the .xlsx) ---
using (var workbook = new XLWorkbook())
{
    var sheet = workbook.AddWorksheet("Sales");

    sheet.Cell("A1").Value = "{{ Company }} — every line";
    sheet.Cell("A1").Style.Font.SetBold().Font.SetFontSize(14);

    sheet.Cell("A3").Value = "Product";
    sheet.Cell("B3").Value = "Quantity";
    sheet.Cell("C3").Value = "Unit price";
    sheet.Range("A3:C3").Style.Font.SetBold();

    sheet.Cell("A4").Value = "{{ item.Product }}";
    sheet.Cell("B4").Value = "{{ item.Quantity }}";
    sheet.Cell("C4").Value = "{{ item.UnitPrice }}";
    sheet.Cell("C4").Style.NumberFormat.Format = "#,##0.00";

    sheet.Cell("B5").Value = "<<Sum>>";

    workbook.DefinedNames.Add("Sales", sheet.Range("A4:C5"));
    workbook.SaveAs("SalesTemplate.xlsx");
}

// --- Generate the report ---
var sales = new List<Sale>
{
    new("Widget", "North", 12, 9.99m),
    new("Gadget", "South", 4, 24.50m),
    new("Doohickey", "North", 27, 3.75m),
};

using var template = new XLTemplate("SalesTemplate.xlsx");
template.AddVariable("Company", "Contoso Ltd");
template.AddVariable("Sales", sales);

template.Generate();
template.SaveAs("SalesReport.xlsx");
```

The generated workbook has three data rows where the template had one, the number format and the
bold heading intact, and a live `=SUBTOTAL(9,B4:B6)` in the options row.

## Binding data

`AddVariable` is the only way data reaches a template.

```csharp
// By name
template.AddVariable("Company", "Contoso Ltd");
template.AddVariable("Sales", sales);              // a collection binds a range
template.AddVariable("PrintedOn", DateTime.Today);

// Every public property and field of an object, under its own name
template.AddVariable(new { Company = "Contoso Ltd", Year = 2026 });

// A dictionary contributes its entries
template.AddVariable(new Dictionary<string, object?>
{
    ["Company"] = "Contoso Ltd",
    ["Sales"] = sales,
});
```

A `DataTable` is materialised into a list of column-keyed rows, so `{{ row.Customer }}` works the
way a template author expects rather than needing an indexer:

```csharp
template.AddVariable("Orders", dataTable);   // then bind a defined name called "Orders"
```

Adding the same alias twice replaces the earlier value. Names are matched **case-sensitively**,
matching C# member semantics.

### Property paths in a defined name

A defined name may reach into a bound object with an underscore. A name `Order_Lines` binds
`Order.Lines`, where `Order` is the bound variable:

```csharp
template.AddVariable("Order", order);
// A defined name "Order_Lines" over a block of rows repeats it once per order.Lines item.
```

## Generating

`Generate()` resolves the whole workbook in place and may be called **once** per template.

```csharp
var result = template.Generate();
```

After it returns, `template.Workbook` *is* the generated workbook — you can go on editing it with
the ordinary XLibur API before saving, which is the escape hatch for anything the template
language cannot express.

```csharp
template.Generate();

// Ordinary XLibur, on the generated result
template.Workbook.Worksheet("Sales").SheetView.FreezeRows(3);

template.SaveAs("SalesReport.xlsx");
```

### Saving to a stream

```csharp
using var stream = new MemoryStream();
template.SaveAs(stream);

return File(
    stream.ToArray(),
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "report.xlsx");
```

### Working from a workbook you already have

Three constructors open a template. Passing a path or a `Stream` gives the template ownership of
the workbook, so disposing the template closes it. Passing an `IXLWorkbook` does **not** — you
keep ownership:

```csharp
using var workbook = new XLWorkbook("SalesTemplate.xlsx");
workbook.Worksheet("Sales").Cell("A2").Value = "Draft";   // amend before generating

using (var template = new XLTemplate(workbook))           // does not own `workbook`
{
    template.AddVariable("Sales", sales);
    template.Generate();
}

workbook.SaveAs("SalesReport.xlsx");                      // still open
```

## Errors

Generation **does not throw** for a bad expression or a tag it cannot apply. A hundred-page report
with one bad cell is worth having; the same report as an exception is not.

Each failure is recorded on the result, the offending cell is left showing the message in red, and
everything else is generated:

```csharp
var result = template.Generate();

if (result.HasErrors)
{
    foreach (var error in result.ParsingErrors)
    {
        logger.LogWarning("{Location}: {Message}", error.Location, error.Message);
    }
}
```

| Member | What it gives you |
|---|---|
| `XLGenerateResult.HasErrors` | Whether anything failed |
| `XLGenerateResult.ParsingErrors` | The failures, in the order they were found |
| `TemplateError.Message` | What went wrong, in the words the cell shows |
| `TemplateError.SheetName` / `.CellAddress` | Where, when known |
| `TemplateError.Location` | The two above as `Sales!B7`, or empty |
| `TemplateError.Exception` | The underlying exception, when there was one |

:::caution What is *not* an error
A variable nobody bound, and a property an object has not got, both read as **blank**. That is
what lets one template survive an optional field — a middle name, a discount that is sometimes
there — without testing for each.

The price is that a **typo in a name is silent**. A column that comes out empty is the first place
to look for one.
:::

## Where to next

- [Expressions](./report-expressions.md) — the `{{ }}` language, Excel functions, and building
  formulas with `&=`
- [Tags](./report-tags.md) — the full `<<Tag>>` reference, grouping and subtotals, horizontal
  ranges, and writing a tag of your own
- [Charts, pivot tables and pictures](./report-charts-and-pivots.md) — what survives range
  expansion, and what you have to draw
- [Coming from ClosedXML.Report](./report-migration.md) — porting a template, and the compatibility
  engine

The repository's
[`XLibur.Report.Examples`](https://github.com/XLibur/XLibur/tree/main/XLibur.Report.Examples)
project holds ten runnable examples, each writing **both** the template it authored and the report
generated from it. Opening a pair side by side is the fastest way to understand any of this.
