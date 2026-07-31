---
id: report-expressions
title: Report Expressions
sidebar_label: Expressions
description: The {{ }} template expression language in XLibur.Report — Scriban syntax, cell types, Excel functions inside expressions, and building formulas with &=.
---

# Expressions

`{{ … }}` anywhere in a cell's text is evaluated and replaced. The language is
[Scriban](https://github.com/scriban/scriban), with two deliberate departures from its defaults
and XLibur's Excel function library bridged in.

```
{{ Company }}                                a workbook variable
{{ item.Product }}                           a property of the current row's item
{{ item.Quantity * item.UnitPrice }}         arithmetic
{{ item.SoldOn | date.to_string "%d %b" }}   a Scriban filter
{{ SUM(array.map Sales "Total") }}           an Excel function
Sold to {{ item.Customer }} in {{ Year }}    text mixed with expressions
```

## Names in scope

| Name | Where it exists | What it is |
|---|---|---|
| Any bound alias | Everywhere | Whatever `AddVariable` bound under that name |
| `item` | Inside a bound range | The current row's item |
| `index` | Inside a bound range | The current row's zero-based position |
| `items` | Inside a bound range | The whole collection the range is bound to |

`index` and `items` are what a row uses to refer to its neighbours or to the set as a whole:

```
{{ index + 1 }}                          a 1-based row number
{{ item.Total / SUM(array.map items "Total") }}    this row's share of the report
```

Property names keep their **C# spelling**. Scriban would ordinarily rename `UnitPrice` to
`unit_price`; that is turned off, so `{{ item.UnitPrice }}` binds the way a template author
naturally writes it. Names are case-sensitive.

## Cell type is decided by what is in the cell

A cell whose text is **nothing but one expression** keeps the value's .NET type. A cell that mixes
text and expressions can only produce text — because that is what it is.

| Template cell | Result |
|---|---|
| `{{ item.UnitPrice }}` | A **number** cell — formattable, chartable, summable |
| `{{ item.SoldOn }}` | A **date** cell |
| `{{ item.IsPaid }}` | A **boolean** cell |
| `{{ item.Discount }}` where the value is `null` | A **blank** cell |
| `Price: {{ item.UnitPrice }}` | **Text** |

This matters more than it looks. Apply a number format to the template cell and the generated
cell keeps it, because the value landing in it is a number rather than a string that resembles
one.

:::note
Expressions are read in **cell values** only. Comments, hyperlinks and rich text are not evaluated.
:::

## What is not an error

A variable nobody bound, and a property an object has not got, both read as **blank** rather than
throwing. Relaxed member and target access are on deliberately: report data is routinely sparse,
and one template has to survive an optional field without testing for each one.

```
{{ item.Customer.MiddleName }}      blank when there is no middle name
{{ item.Customer.Address.City }}    blank when Address is null — no null check needed
{{ Subtitle }}                      blank when nothing bound "Subtitle"
```

The cost is that a **misspelled name is silent**. When a column comes out empty, check the
spelling before anything else.

## Excel functions

Every function XLibur's calculation engine implements is callable inside `{{ }}` under its
**upper-case Excel name** — `SUM`, `IF`, `ROUND`, `MAX`, `MIN`, `CEILING`, `TEXT`, `VLOOKUP`, and
the rest of the library. A report author who knows Excel does not have to learn a second set of
names for the same things.

```
{{ ROUND(item.UnitPrice * item.Quantity, 2) }}
{{ IF(item.Quantity > 100, "Bulk", "Standard") }}
{{ MAX(array.map items "Total") }}
{{ TEXT(item.SoldOn, "yyyy-mm-dd") }}
```

Upper case is not cosmetic: Scriban's keywords are lower case, so `if` would be parsed as a block
conditional while `IF` parses as an ordinary function call.

Results are **typed**. `{{ SUM(…) }}` over decimals lands in a number cell that Excel can format
and total, not text that looks like a number.

:::caution Functions that need a grid
`OFFSET`, `INDIRECT`, `ROW`, `COLUMN` and the other context-dependent functions are registered
like the rest, but a template expression is evaluated before there is a cell to be relative to.
Calling one reports that it was used outside a worksheet rather than returning something
misleading.

Write a real Excel formula for those and let Excel evaluate it.
:::

### Aggregating a collection

Excel's aggregate functions take a *range*; a template expression has a collection. Scriban's
`array.map` pulls one property out of every item and hands the resulting list over:

```
{{ SUM(array.map items "Total") }}          total of the bound range
{{ SUM(array.map Sales "Quantity") }}       total of a collection by name, anywhere in the sheet
{{ AVERAGE(array.map items "UnitPrice") }}
```

For a total of the *generated rows* in the options row, prefer the [`<<Sum>>`
tag](./report-tags.md#summaries) — it leaves a live `SUBTOTAL` formula that stays correct if
someone edits or filters the report afterwards.

## Scriban filters

Scriban's own filters are available and are often the better tool for shaping data:

```
{{ item.Product | string.upcase }}
{{ item.SoldOn | date.to_string "%d %b %Y" }}
{{ items | array.sort "Total" }}
{{ items | array.filter @(do; ret $0.Quantity > 10; end) }}
{{ item.Notes | string.truncate 40 }}
```

The [Scriban built-in function reference](https://github.com/scriban/scriban/blob/master/doc/builtins.md)
lists them all.

## Building formulas with `&=`

`&=` at the **start** of a cell's text builds an Excel formula at generation time rather than a
value. Everything after the prefix is interpolated, and the result becomes the cell's formula:

```
&=SUM(B{{ FirstDataRow }}:B{{ LastDataRow }})
&=VLOOKUP(A{{ index + 4 }}, Rates!$A:$B, 2, FALSE)
```

:::tip Prefer an ordinary formula where one will do
A template cell holding a plain Excel formula — `=B4*1.2` — is copied into every generated row
with its **relative references re-pointed**, so it becomes `=B5*1.2`, `=B6*1.2` and so on. That is
handled by the core library and needs no `&=` and no expression.

Reach for `&=` only when the formula's *shape* depends on the data — a range whose extent you do
not know until generation, or a reference built from a bound value.
:::

## Choosing a culture

Expressions evaluate under `CultureInfo.InvariantCulture` by default, so a generated report does
not change shape with the machine's locale — and a number interpolated into a formula never gets
a decimal comma, which would not be a formula.

To evaluate under a specific culture, construct the engine yourself and pass it to the template:

```csharp
using XLibur.Report.Expressions;

var engine = new ScribanExpressionEngine(new CultureInfo("de-DE"));

using var template = new XLTemplate("SalesReport.xlsx", engine);
```

## Registering a function of your own

Anything callable can be exposed to templates:

```csharp
var engine = new ScribanExpressionEngine();
engine.AddFunction("SLUG", (string text) => text.ToLowerInvariant().Replace(' ', '-'));

using var template = new XLTemplate("SalesReport.xlsx", engine);
// {{ SLUG(item.Product) }}
```

Register under an upper-case name for the same reason the Excel bridge does: it keeps the name
clear of Scriban's lower-case keywords.

:::note
An engine caches parsed expressions, so sharing one across templates is cheaper than constructing
one per report. An engine is **not** thread-safe, though — a template generates on one thread at a
time, and two templates generating concurrently need an engine each.
:::
