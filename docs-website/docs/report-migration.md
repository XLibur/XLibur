---
id: report-migration
title: Coming from ClosedXML.Report
sidebar_label: From ClosedXML.Report
description: Port a ClosedXML.Report template to XLibur.Report — what carries over unchanged, the expression syntax differences, and the compatibility engine that runs the old syntax as written.
---

# Coming from ClosedXML.Report

XLibur.Report uses ClosedXML.Report's template model, so **a template carries over**. What changes
is the expression syntax, because the default engine is
[Scriban](https://github.com/scriban/scriban) rather than System.Linq.Dynamic.Core.

If you would rather not touch the expressions at all, skip to
[keeping the old syntax](#or-keep-the-old-syntax).

## What is unchanged

- Defined names bind ranges to collections, and the range's last row is its options row.
- `<<Tag>>` markers, with the same names and the same meanings.
- `{{ }}` marks an expression.
- `&=` at the start of a cell builds a formula.
- Property paths in a defined name: `Order_Lines` binds `Order.Lines`.
- The API shape: `new XLTemplate(path)`, `AddVariable`, `Generate`, `SaveAs`.

```csharp
using XLibur.Report;

using var template = new XLTemplate("SalesReport.xlsx");
template.AddVariable("Company", "Contoso");
template.AddVariable("Sales", sales);
template.Generate();
template.SaveAs("out.xlsx");
```

## What is better

Three things ClosedXML.Report never handled, and one it handled badly:

| | |
|---|---|
| **Charts** | A series drawn over the template's repeated row plots every generated row. Upstream left it naming the one template row — its issues #123 and #351. |
| **Pivot tables** | A template pivot is re-pointed at the generated rows, refreshed, and moved if the rows grew over it. `<<Pivot>>` builds one from the template's own shape. |
| **Pictures** | A picture below a bound range ends up below the generated rows. |
| **Conditional formatting** | A rule over the repeated rows is *stretched* over the generated block. Upstream copies it per generated cell — three rules over three rows becoming nine, its issue #216. |

See [Charts, pivot tables and pictures](./report-charts-and-pivots.md).

## What changes: the expression syntax

Scriban is a template language, not C#. The common translations:

| ClosedXML.Report | XLibur.Report |
|---|---|
| `{{item.Name}}` | `{{ item.Name }}` — unchanged |
| `{{item.Qty * item.Price}}` | unchanged |
| `{{item.Name.Substring(0,3)}}` | `{{ item.Name \| string.slice 0 3 }}` |
| `{{item.Name.ToUpper()}}` | `{{ item.Name \| string.upcase }}` |
| `{{items.Sum(x => x.Total)}}` | `{{ SUM(array.map items "Total") }}` |
| `{{items.Count()}}` | `{{ items.size }}` |
| `{{items.Where(x => x.Qty > 10)}}` | `{{ items \| array.filter @(do; ret $0.Qty > 10; end) }}` |
| `{{item.SoldOn.ToString("dd MMM")}}` | `{{ item.SoldOn \| date.to_string "%d %b" }}` |
| `{{DateTime.Now}}` | bind it as a variable — the sandbox has no reflection escape |

Two things that do **not** change and often surprise people mid-port:

- Property names keep their C# spelling. Scriban would ordinarily rename `UnitPrice` to
  `unit_price`; that is turned off.
- A missing property still reads as blank rather than throwing, the same as upstream.

You also gain [Excel's own functions](./report-expressions.md#excel-functions) inside `{{ }}` —
`SUM`, `IF`, `ROUND`, `VLOOKUP` and the rest of the library, under their Excel names.

## Or keep the old syntax

Install **`XLibur.Report.DynamicLinq`** and pass its engine. Templates written for ClosedXML.Report
then run **as written** — no expression changes at all:

```sh
dotnet add package XLibur.Report.DynamicLinq
```

```csharp
using XLibur.Report;
using XLibur.Report.DynamicLinq;

using var template = new XLTemplate("LegacyReport.xlsx", new DynamicLinqExpressionEngine());
template.AddVariable("Sales", sales);
template.Generate();
template.SaveAs("out.xlsx");
```

Everything structural is unchanged — the defined names, the options row, the tags, `&=` — because
none of it goes through the engine. Property and method access, arithmetic, the conditional
operator and LINQ over collections in scope all work. `item`, `index` and `items` are the row
bindings, and a workbook variable is reachable as `Company` or as `@Company`, as upstream allowed.

Engine choice is **per template**, so a codebase can port templates one at a time. `XLibur.Report`
never references the package, so adding or removing it changes nothing for code using the default.

### Two differences worth knowing

**The Excel-function bridge is not available under this engine.** `{{ SUM(...) }}` is an unknown
name. It is a feature of the default engine, upstream syntax never had it, and templates written
for that syntax call .NET methods instead.

:::danger Trusted templates only
Dynamic LINQ has **no sandbox**. An expression it parses can reach the methods and properties of
any object in scope, and the library's history includes CVE-2023-32571 — arbitrary method
invocation.

Point this engine at your own templates, never at one a user uploaded. For that, use the default
Scriban engine, which has real execution limits and no reflection escape. That is why it is the
default and this one is opt-in.
:::

## Not implemented yet

An honest list, so a port does not discover these the hard way:

- Expressions in **comments, hyperlinks and rich text** — cell values only for now.
- **Nested vertical subranges** — a child range inside a parent's rows.
- The `Image`, `PageOptions`, `Protected`, `Height`, `OnlyValues` and `Range` tags.
- `<<Group>>` and `<<Pivot>>` are **vertical-only**, by design rather than by omission. See
  [Ranges that repeat across](./report-tags.md#ranges-that-repeat-across).

## Packaging and versioning

`XLibur.Report` and `XLibur.Report.DynamicLinq` release on their **own version stream**, separate
from the core `XLibur` package. A report release does not need a core release, and a core release
does not renumber the report packages.
