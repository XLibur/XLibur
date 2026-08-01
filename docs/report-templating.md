# Report templating (`XLibur.Report`)

Author a report as an ordinary `.xlsx` template — placeholder expressions, defined names and tag
markers — bind .NET data to it, and generate the finished workbook. The template is a real Excel file,
so its formatting, formulas, charts, pivot tables and pictures are authored in Excel by whoever knows
what the report should look like, and none of them are described in code.

```csharp
using XLibur.Report;

using var template = new XLTemplate("SalesReport.xlsx");
template.AddVariable("Company", "Contoso");
template.AddVariable("Sales", sales);

var result = template.Generate();
if (result.HasErrors)
{
    foreach (var error in result.ParsingErrors)
    {
        logger.LogWarning("{Location}: {Message}", error.Location, error.Message);
    }
}

template.SaveAs("SalesReport-2026.xlsx");
```

Ten worked examples live in [`XLibur.Report.Examples`](../XLibur.Report.Examples/README.md), each
writing both the template it authored and the report generated from it. Opening a pair is the fastest
way to understand any of what follows.

---

## Expressions

`{{ … }}` anywhere in a cell's text is evaluated and replaced. The language is
[Scriban](https://github.com/scriban/scriban).

```
{{ Company }}                             a workbook variable
{{ item.Product }}                        a property of the current row's item
{{ item.Quantity * item.UnitPrice }}      arithmetic
{{ item.SoldOn | date.to_string "%d %b" }}   a Scriban filter
{{ SUM(array.map Sales "Total") }}        an Excel function (see below)
```

A cell whose text is **nothing but** one expression keeps the value's **type**: a decimal lands in a
number cell, a `DateTime` in a date cell, `null` in a blank one. A cell that mixes text and expressions
produces text, because that is what it is.

Expressions are read in **cell values**. Comments, hyperlinks and rich text are not evaluated yet.

`&=` at the start of a cell's text builds a **formula** at generation time rather than a value:
`&=SUM(B{{ FirstRow }}:B{{ LastRow }})` writes a formula into the cell. Prefer an ordinary Excel
formula in the template where one will do — the engine's insert-and-copy re-points relative references
for you, so a template cell holding `=E2*F2` becomes `=E7*F7` in the row it lands in.

### What is not an error

A variable nobody bound, and a property an object has not got, both read as **blank**. That is what
lets one template survive an optional field — a middle name, a discount that is sometimes there —
without testing for each. The price is that a **typo in a name is silent**, so a column that comes out
empty is the first place to look for one.

### Culture

Everything the report formats or orders follows **one** culture, and it is the template's rather than
the machine's. A default `XLTemplate` is invariant end to end, so the same template over the same data
produces the same workbook wherever it is generated — the same row order, the same group labels, the
same decimal point. Pass a culture to the engine to say otherwise:

```csharp
using var template = new XLTemplate("Rapport.xlsx", new ScribanExpressionEngine(new CultureInfo("sv-SE")));
```

That one argument decides how an interpolated number or date reads, how `<<Sort>>` and `<<Group>>`
collate text — Swedish sorting `å` after `z`, Czech treating `ch` as a single letter — and how a
group's `{0} Total` label formats a date or a decimal key. A custom tag that needs the same answer
reads `context.Engine.Culture`.

## Bound ranges

A **defined name** matching the name of a variable that holds a collection makes the rows it covers
repeat, once per item. Inside those rows, `item` is the current one.

The range's **last row is its options row**. It carries the tags and the totals, it is not repeated, and
it is removed if nothing was written into it. A range one row deep has no options row and repeats
entirely.

```
        A                  B                C
1   Product            Quantity         Line total      ← heading, outside the range
2   {{ item.Product }} {{ item.Qty }}   =B2*1.2         ← repeated, once per item
3                      <<Sum>>          <<Sum>>         ← options row
        └───────────── defined name "Sales" covers A2:C3 ─────────────┘
```

Defined names may address a property path with an underscore: a name `Order_Lines` binds
`order.Lines` where `Order` is the bound variable.

### Name scope

Names bind under Excel's own scoping. A name scoped to a **sheet** may be declared once per sheet, and
every one of them binds — the natural way to write a template with a section per sheet:

```
Q1!Items → Q1!A2:C3     both bind, both read the Items variable
Q2!Items → Q2!A2:C3
```

A name scoped to the **workbook** binds everywhere except on a sheet that declares its own name of that
name, which shadows it there and nowhere else. Shadowing is silent: it is what Excel does, so it is not
reported as a template error.

The name is what selects the variable, so all the ranges above read `Items`. To bind two sheets to
different data, give the ranges different names.

## Tags

`<<Tag param=value>>` in a range's options row. Parameters may be bare flags (`desc`), assigned
(`over=D`) or quoted (`totalLabel="Grand total"`). Tag text is stripped as it is read, so it never
reaches the report.

Most tags describe the range and belong in the options row. A tag written in a **repeated** row is
describing one item — which is the difference between `<<If>>` dropping a row and dropping the range.

| Tag | What it does |
|---|---|
| `<<Sum>>` `<<Avg>>` `<<Count>>` `<<CountA>>` `<<Max>>` `<<Min>>` `<<Product>>` `<<StdDev>>` `<<StdDevP>>` `<<Var>>` `<<VarP>>` | Totals the column it sits under, as a live `SUBTOTAL` formula. `over=D` totals a different column. |
| `<<Sort>>` `<<Asc>>` `<<Desc>>` | Orders the rows by the column's own expression. `by="…"` sorts by something the range does not display. |
| `<<Group>>` | Groups by the column it sits under: a subtotal row per group, an Excel outline, and the rows ordered so groups come out contiguous. Several nest, leftmost outermost. See below. |
| `<<If test="…">>` | In a repeated row, drops that row when the test is falsy. In the options row, drops the whole range. |
| `<<Horizontal>>` | In the range's **last column**: the range repeats across, one column per item. |
| `<<AutoFilter>>` | Excel's autofilter over the generated rows and the heading above them. `noheader` for the rows alone. |
| `<<ColsFit>>` `<<RowsFit>>` | Fit the range's columns / the generated rows to their contents. Needs a font engine registered. |
| `<<Hidden>>` `<<Hide>>` | Hides the line the tag sits in — for a column a template needs in order to sort or total, but the reader should not see. |
| `<<Delete>>` | Removes the line the tag sits in, after everything else has run. `keep="{{ … }}"` makes the removal conditional. |
| `<<Pivot dest="…">>` | Builds a pivot table over the generated rows. See below. |
| `<<Row>>` `<<Column>>` `<<Col>>` `<<Page>>` `<<Data>>` | Under a column, says what `<<Pivot>>` should use it as. `<<Data>>` takes a summary name (`<<Data avg>>`) and `title=`. |
| `<<SummaryAbove>>` `<<MergeLabels>>` `<<PageBreaks>>` `<<Collapse>>` `<<DisableSubtotals>>` `<<DisableGrandTotal>>` | The range-wide form of the like-named `<<Group>>` parameter, so a template with several group levels does not repeat itself. |

Only `null` and `false` are false in a `test=`. **Zero and the empty string are true**, so a test
meaning "more than nothing" has to say so: `test="item.Quantity > 0"`.

### Grouping and subtotals

`<<Group>>` sits under the column to group by and takes that column's expression as its key, the same
way `<<Sort>>` does.

- Each group gets a **subtotal row** carrying whatever summary tags the options row declares, over that
  group's rows alone, plus a `{0} Total` label in the grouped column (`totalLabel=` to change it).
- The subtotal row **takes the options row's styling** — the only styling a template can express for a
  row that does not exist until generation, and what makes a group total look like the grand total.
- The block is **outlined**, so Excel's outline buttons collapse to subtotals and again to the grand
  total. `collapse` writes it collapsed.
- The engine **orders the rows** by the group keys, stably — so a `<<Sort>>` on another column still
  decides the order *within* a group. `nosort` groups the given order as it comes.
- The grand total does not double-count the group totals, because `SUBTOTAL` ignores nested
  `SUBTOTAL`s. `<<DisableGrandTotal>>` keeps the group totals and drops the report total.

Parameters: `by`, `desc`, `nosort`, `totalLabel`, `merge`/`mergeLabels`, `summaryAbove`, `pageBreaks`,
`collapse`, `disableSubtotals`.

### Ranges that repeat across

`<<Horizontal>>` in a range's last column turns everything ninety degrees: the last **column** is the
options column, the columns before it repeat, and a tag sits in a **row** and acts on that row. It suits
a report with few items and many measures — a region or a quarter per column.

`<<Group>>` and `<<AutoFilter>>` do not apply across, and report an error rather than doing something
surprising: a subtotal *column* labelled with a group key is not a thing report readers ask for, and
Excel filters rows.

## Excel functions in expressions

Every function XLibur's calculation engine implements is callable inside `{{ }}` under its upper-case
Excel name — `SUM`, `IF`, `ROUND`, `MAX`, `CEILING`, `TEXT`, `VLOOKUP`, and the rest of the library. A
report author who knows Excel does not have to learn a second set of names for the same things, and
`IF` works despite `if` being a keyword of the expression language.

Results are **typed**: `{{ SUM(…) }}` over decimals lands in a number cell that can be formatted and
totalled, not text that looks like a number.

Functions that need a grid to work on — `OFFSET`, `INDIRECT`, `ROW`, `COLUMN` and the other
context-dependent ones — are registered like the rest, but a template expression has no cell to be
relative to, so calling one reports that it was used outside a worksheet rather than returning something
misleading. Write a real Excel formula for those and let Excel evaluate it.

Scriban's own filters are available too and are usually the better tool for shaping data:
`array.map`, `array.filter`, `array.sort`, `string.upcase`, `date.to_string`. So
`{{ SUM(array.map Sales "Total") }}` pulls one property out of every item and hands Excel the list.

## Charts, pivot tables and pictures

These are the three things ClosedXML.Report never handled, and they need nothing from the template
author beyond drawing them on the data.

**Charts.** Draw the chart in the template over the single row the template has. After generation its
series cover every row generated. Nothing in the template says so and there is no tag.

**Pictures.** Nothing to do. A picture below a bound range ends up below the generated rows.

**Pivot tables, drawn.** The pattern to reach for. Build the pivot in Excel over the template's data
rows, lay it out however the report wants; the engine points its cache at what was generated, refreshes
it, marks it to re-read on open, and moves the pivot itself if the generated rows would have grown over
it. A pivot over a **table** or a **defined name** needs no re-pointing — the table grew on its own —
but still gets the refresh.

**Pivot tables, generated.** For a pivot whose shape belongs to the template rather than to whoever
drew it:

```
        A            B              C
3   Region       Category       Line total          ← heading row: the pivot's field names
4   {{ item… }}  {{ item… }}    {{ item… }}
5   <<Row>>      <<Column>>     <<Data title="Sold">><<Pivot dest="Summary!A3">>
```

`dest` is required and is read in **template coordinates**, like everything else a template writes: a
cell marked below the range comes out below the *generated* range. It takes a sheet-qualified reference
(`Summary!A3`), a plain one for the range's own sheet (`F1`), or the name of a defined name covering one
cell. `name=` names the pivot table.

Each field is named by its column's **heading** — the row above the range, which is where a pivot cache
reads its field names from anyway. A range starting in row 1 has no heading row, and reports that rather
than inventing names.

`<<Pivot>>` is not supported in a horizontal range (no heading row to name fields from) or in a grouped
range (its subtotal rows are inside the generated block, and the pivot would count them as data on top
of the rows they already total). Both report an error.

## Conditional formatting

A rule over the template's repeated rows is **stretched** over the generated block, not copied per row.
A template that declares three rules produces three rules over however many rows it generated. (The
engine this one is a port of produces `rows × rules`, which is its issue #216.)

## Errors

Generation does not throw for a bad expression or a tag it cannot apply. The failure is recorded on
`XLGenerateResult.ParsingErrors` with the sheet and cell it came from, the offending cell is left
showing the message, and everything else is generated. A hundred-page report with one bad cell is worth
having; the same report as an exception is not.

## Custom tags

```csharp
public sealed class TopTag : OptionTag
{
    public override IReadOnlyList<object?> TransformItems(IReadOnlyList<object?> items, ProcessingContext context)
        => items.Take((int)Token.Number("count", 10)).ToList();
}

TagsRegister.Add<TopTag>("Top", priority: 15);
```

A tag acts at one of two moments and may use both. `TransformItems` runs **before any row exists**,
which is where reordering and filtering belong — nothing downstream then has to know that anything was
dropped. `Execute` runs **once the rows are there**, which is where anything referring to the generated
block belongs: a total, a border, a column width.

`priority` is how a tag says what it has to see first — lower runs earlier. `<<Sort>>` is 10, `<<Group>>`
20, the summaries 50, the layout tags 200, `<<Delete>>` 250. `ProcessingContext` hands a tag the
generated range, the items, the expression engine, the errors list, the other tags in the range, and
`IsTrue(…)` for evaluating a parameter as a question.

## Coming from ClosedXML.Report

The template model is the same — defined names, an options row, `<<tags>>`, `{{ }}` — so a template
carries over. What changes is the expression syntax, because the default engine is Scriban rather than
System.Linq.Dynamic.Core:

| ClosedXML.Report | XLibur.Report |
|---|---|
| `{{item.Name.Substring(0,3)}}` | `{{ item.Name \| string.slice 0 3 }}` |
| `{{items.Sum(x => x.Total)}}` | `{{ SUM(array.map items "Total") }}` |
| `{{items.Where(x => x.Qty > 10)}}` | `{{ items \| array.filter @(do; ret $0.Qty > 10; end) }}` |
| `{{DateTime.Now}}` | bind it as a variable — the sandbox has no reflection escape |

### Or keep the old syntax

Install **`XLibur.Report.DynamicLinq`** and pass its engine, and upstream templates run as written:

```csharp
using XLibur.Report.DynamicLinq;

using var template = new XLTemplate("LegacyReport.xlsx", new DynamicLinqExpressionEngine());
```

Everything structural is unchanged — the defined names, the options row, the tags, `&=` — because none
of it goes through the engine. Property and method access, arithmetic, the conditional operator and LINQ
over collections in scope all work, `item`/`index`/`items` are the row bindings, and a workbook variable
is reachable as `Company` or as `@Company`.

Two differences worth knowing. The Excel-function bridge is **not** available under this engine
(`{{ SUM(...) }}` is an unknown name): it is a feature of the default engine, upstream syntax never had
it, and templates written for that syntax call .NET methods instead. And **it is for trusted templates
only** — Dynamic LINQ has no sandbox, an expression can reach any member of any object in scope, and the
library's history includes CVE-2023-32571. Point it at your own templates, never at one a user uploaded;
for that, use the default engine.

Engine choice is per template, and `XLibur.Report` never references the package, so adding or removing
it changes nothing for code using the default.

## Not implemented yet

Honest list, as of the spec-12 branch: expressions in comments, hyperlinks and rich text (cell values
only for now); nested vertical subranges (a child range inside a parent's rows); and the `Image`,
`PageOptions`, `Protected`, `Height`, `OnlyValues` and `Range` tags. Grouping is vertical-only by design,
as is `<<Pivot>>`.

See [`docs/specs/12-report-templating.md`](specs/12-report-templating.md) for the design, the findings
that changed it, and what remains.
