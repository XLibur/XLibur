---
id: report-tags
title: Report Tags
sidebar_label: Tags
description: The <<Tag>> reference for XLibur.Report — totals, sorting, filtering rows, grouping with subtotals and outlines, horizontal ranges, and writing a tag of your own.
---

# Tags

A tag is a `<<Name param=value>>` marker written in a cell of a bound range. It tells the engine
to do something to the range that an expression cannot express — total a column, sort the rows,
group them, drop one.

Tag text is **stripped as it is read**, so it never reaches the generated report. A cell holding
nothing but tags is cleared entirely, which is how an options row that carried only tags gets
removed.

## Where a tag goes

Most tags describe the range and belong in its **options row** — the last row of the bound range,
which is not repeated.

A tag written in a **repeated** row is describing *one item* rather than the range. That is the
difference between `<<If>>` dropping a row and `<<If>>` dropping the whole range.

Several tags act on the **line** they sit in — the column for a range that repeats downwards, the
row for one that repeats across. `<<Sum>>` totals the column above it; `<<Hidden>>` hides the
column it is in.

## Parameters

```
<<Sort desc>>                       a bare flag
<<Sum over=D>>                      assigned
<<Group totalLabel="Grand total">>  quoted, so the value may contain spaces or >
```

Parameter names are matched case-insensitively. A bare flag means "on".

## Reference

| Tag | What it does |
|---|---|
| `<<Sum>>` `<<Avg>>` `<<Average>>` `<<Count>>` `<<CountA>>` `<<Max>>` `<<Min>>` `<<Product>>` `<<StdDev>>` `<<StdDevP>>` `<<Var>>` `<<VarP>>` | Totals the column it sits under, as a live `SUBTOTAL` formula. `over=D` totals a different column. |
| `<<Sort>>` `<<Asc>>` `<<Desc>>` | Orders the rows by the column's own expression. `by="…"` sorts by something the range does not display. |
| `<<Group>>` | Groups by the column it sits under: a subtotal row per group, an Excel outline, and the rows ordered so groups come out contiguous. Several nest, leftmost outermost. |
| `<<If test="…">>` | In a repeated row, drops that row when the test is falsy. In the options row, drops the whole range. |
| `<<Horizontal>>` | In the range's **last column**: the range repeats across, one column per item. |
| `<<AutoFilter>>` | Excel's autofilter over the generated rows and the heading above them. `noheader` for the rows alone. |
| `<<ColsFit>>` `<<RowsFit>>` | Fit the range's columns, or the generated rows, to their contents. Needs a font engine registered. |
| `<<Hidden>>` `<<Hide>>` | Hides the line the tag sits in. |
| `<<Delete>>` | Removes the line the tag sits in, after everything else has run. `keep="{{ … }}"` makes the removal conditional. |
| `<<Pivot dest="…">>` | Builds a pivot table over the generated rows. See [Charts, pivot tables and pictures](./report-charts-and-pivots.md#pivot-tables-generated-by-the-template). |
| `<<Row>>` `<<Column>>` `<<Col>>` `<<Page>>` `<<Data>>` | Under a column, says what `<<Pivot>>` should use it as. |
| `<<SummaryAbove>>` `<<MergeLabels>>` `<<PageBreaks>>` `<<Collapse>>` `<<DisableSubtotals>>` `<<DisableGrandTotal>>` | The range-wide form of the like-named `<<Group>>` parameter. |

## Summaries

A summary tag sits in the options row under the column to total, and leaves a `SUBTOTAL` formula
rather than a computed number:

|  | A | B | C |
|---|---|---|---|
| **4** | `{{ item.Product }}` | `{{ item.Quantity }}` | `{{ item.Total }}` |
| **5** | Total | `<<Sum>>` | `<<Sum>>` |

Two reasons it is a formula and not a value: the total stays live if someone edits a generated
row, and `SUBTOTAL` ignores rows that are filtered out or that are themselves subtotals — so a
grand total does not double-count the group totals beneath it.

`over` totals a **different** column from the one the tag sits in, which is how a label column can
carry the total of a value column:

```
<<Sum over=D>>      total column D, print it here
```

A summary over **no rows** — a range bound to an empty collection — writes `0` rather than a
broken reference.

| Tag | `SUBTOTAL` function |
|---|---|
| `<<Avg>>` `<<Average>>` | 1 |
| `<<Count>>` | 2 |
| `<<CountA>>` | 3 |
| `<<Max>>` | 4 |
| `<<Min>>` | 5 |
| `<<Product>>` | 6 |
| `<<StdDev>>` | 7 |
| `<<StdDevP>>` | 8 |
| `<<Sum>>` | 9 |
| `<<Var>>` | 10 |
| `<<VarP>>` | 11 |

## Sorting

`<<Sort>>` takes the **sort key from the column it sits under**, so a column already showing
`{{ item.SoldOn }}` needs no second mention of it:

|  | A | B | C |
|---|---|---|---|
| **4** | `{{ item.Product }}` | `{{ item.SoldOn }}` | `{{ item.Total }}` |
| **5** |  | `<<Sort desc>>` |  |

To sort by something the range does not display, give `by`:

```
<<Sort by="item.Customer.Name">>
```

`<<Desc>>` is `<<Sort desc>>` written another way. Sorting happens **before any row is written**,
so everything downstream — grouping, totals, the autofilter — sees the sorted order and nothing
has to know that a sort happened.

Blank keys are kept together at the end of an ascending sort.

## Filtering rows

Where `<<If>>` is written decides what it drops.

In a **repeated row**, it asks the question of each item and keeps the ones that answer yes:

|  | A | B | C |
|---|---|---|---|
| **4** | `{{ item.Product }}` | `{{ item.Quantity }}` | `<<If test="item.Quantity > 0">>` |

In the **options row**, it asks once, of the range as a whole. A no renders the range exactly as
an empty collection does — the rows gone, the headings and any options-row total behaving as they
would over no data:

```
<<If test="items.size > 3">>
```

The test runs before any row exists, so what survives it is what everything else sees.

:::caution Zero is true
Only `null` and `false` are false. **Zero and the empty string are true** — Scriban's rule — so a
test meaning "more than nothing" has to say so:

```
<<If test="item.Quantity > 0">>     correct
<<If test="item.Quantity">>         keeps every row, including the empty ones
```
:::

## Grouping and subtotals

`<<Group>>` sits under the column to group by and takes that column's expression as its key, the
same way `<<Sort>>` does.

|  | A | B | C |
|---|---|---|---|
| **3** | Region | Product | Total |
| **4** | `{{ item.Region }}` | `{{ item.Product }}` | `{{ item.Total }}` |
| **5** | `<<Group>>` |  | `<<Sum>>` |

What that produces:

- A **subtotal row per group**, carrying whatever summary tags the options row declares, over that
  group's rows alone, plus a `{0} Total` label in the grouped column (`totalLabel=` to change it,
  where `{0}` is the group key).
- The subtotal row **takes the options row's styling**. That is the only styling a template can
  express for a row that does not exist until generation, and it is what makes a group total look
  like the grand total.
- An **Excel outline**, so the outline buttons collapse to subtotals and again to the grand total.
- The rows **ordered by the group keys**, stably — so a `<<Sort>>` on another column still decides
  the order *within* a group.
- A grand total that does not double-count the group totals, because `SUBTOTAL` ignores nested
  `SUBTOTAL`s.

Several `<<Group>>` tags **nest, leftmost outermost**.

### Group parameters

| Parameter | Effect |
|---|---|
| `by="…"` | Group by an expression the range does not display |
| `desc` | Largest group key first |
| `nosort` | Leave the row order alone; group runs of equal keys as they come |
| `totalLabel="…"` | The subtotal row's label; `{0}` is the group key. Default `"{0} Total"` |
| `merge` / `mergeLabels` | Merge the group's repeated label cells into one |
| `summaryAbove` | Put the subtotal row above the group rather than below |
| `pageBreaks` | A page break after each group |
| `collapse` | Write the outline collapsed |
| `disableSubtotals` / `noSubtotals` | Outline the group without a subtotal row |

Each has a **range-wide form** written as its own tag anywhere in the options row —
`<<SummaryAbove>>`, `<<MergeLabels>>`, `<<PageBreaks>>`, `<<Collapse>>`, `<<DisableSubtotals>>` —
so a template with several group levels does not repeat itself. A level may still turn one on for
itself alone.

`<<DisableGrandTotal>>` is the exception, having no per-level form: it leaves the options row's own
summaries unwritten, which is how a report shows a total per group and none for the report.

## Ranges that repeat across

`<<Horizontal>>` in a range's **last column** turns everything ninety degrees. The last *column*
becomes the options column, the columns before it repeat, and a tag sits in a **row** and acts on
that row.

|  | B | C |
|---|---|---|
| **4** | `{{ item.Region }}` | `<<Horizontal>>` |
| **5** | `{{ item.Quantity }}` | `<<Sum>>` |
| **6** | `{{ item.Total }}` | `<<Sum>>` |

The defined name covers `B4:C6`: column B repeats once per item, column C stays for the options.
It suits a report with few items and many measures — a region or a quarter per column, with the
labels a reader reads down the side sitting in a column *outside* the range.

The tag has to be found before the range can be read at all, which is why it is looked for in the
last column specifically.

:::note Not supported across
`<<Group>>` and `<<AutoFilter>>` report an error rather than doing something surprising. Excel
filters rows, not columns; and a subtotal *column* labelled with a group key is not a thing report
readers ask for. `<<Pivot>>` is out too — a pivot names its fields from a heading row, which a
horizontal range does not have.
:::

## Layout tags

```
<<AutoFilter>>            filter over the generated rows and the heading above them
<<AutoFilter noheader>>   the generated rows alone
<<ColsFit>>               size the range's columns to their contents
<<RowsFit>>               size the generated rows to their contents
<<Hidden>>                hide the line the tag sits in
<<Delete>>                remove the line the tag sits in
```

`<<Hidden>>` and `<<Delete>>` are for a column a template needs in order to sort or total by, but
that the reader should not see. `<<Delete>>` runs last of all, so the column may be sorted and
totalled by and then removed.

`keep` makes the removal conditional, which is how one template serves a summary and a detailed
version:

```
<<Delete keep="{{ ShowWorkings }}">>    keep the column when the variable is truthy
<<Delete keep>>                         keep it outright
```

:::note
`<<ColsFit>>` and `<<RowsFit>>` measure text, so they need a font engine registered. See
[Getting Started](./getting-started.md#installation).
:::

## Writing a tag of your own

A tag is a class deriving from `OptionTag`, registered by name:

```csharp
using XLibur.Report.Tags;

/// <summary>&lt;&lt;Top count=5&gt;&gt; — keep only the first N rows.</summary>
public sealed class TopTag : OptionTag
{
    public override IReadOnlyList<object?> TransformItems(
        IReadOnlyList<object?> items, ProcessingContext context)
        => items.Take((int)Token.Number("count", 10)).ToList();
}

TagsRegister.Add<TopTag>("Top", priority: 15);
```

Registration is global and lasts for the process; registering over an existing name replaces it.

### A tag's two moments

| Override | When it runs | What belongs there |
|---|---|---|
| `TransformItems` | **Before any row exists** | Reordering, filtering — anything that changes *what* is generated. Nothing downstream then has to know something was dropped. |
| `Execute` | **Once the rows are there** | Anything referring to the generated block: a total, a border, a column width, an autofilter. |

A tag may use both.

### Priority

`priority` is how a tag says what it has to see first — **lower runs earlier**. Tags of equal
priority run in the order they were read, left to right.

| Priority | Built-in tags |
|---|---|
| 0 | `<<Horizontal>>` |
| 1 | `<<If>>` |
| 5 | The range-wide group options |
| 10 | `<<Sort>>` `<<Asc>>` `<<Desc>>` |
| 20 | `<<Group>>` |
| 50 | The summaries |
| 200 | The layout and pivot-field tags |
| 250 | `<<Delete>>` |
| 255 | `<<Pivot>>` |

So a tag that filters rows wants a priority below 10; one that measures the finished block wants
one above 50.

### Reading parameters

```csharp
Token.Name                          // "Top", as written
Token.Has("count")                  // was it given at all?
Token.Value("count")                // "5", or "" when absent
Token.Value("label", "Total")       // with a fallback
Token.Number("count", 10)           // parsed as a double, invariant culture
Token.Flag("desc")                  // bare, or =true, or =1
```

### What `ProcessingContext` gives you

| Member | What it is |
|---|---|
| `Worksheet` | The sheet being generated |
| `GeneratedRange` | What the range produced, excluding its options row |
| `OptionsRow` | The options row, or `null` once it has been removed |
| `Items` | The data the range was generated from, in the order written |
| `Engine` | The expression engine, for evaluating a parameter |
| `Scope` | The names in scope for the range |
| `Errors` | Where a tag reports a problem instead of throwing |
| `LineExpressions` | The expression each line of the range held — how `<<Sort>>` knows what to sort by |
| `Tags` | Every tag the range declared, including this one |
| `IsHorizontal` | Whether the range repeats across |
| `IsTrue(expression, scope)` | Evaluates a parameter as a question, recording a failure as an error and answering no |

`OptionTag` itself carries `Token`, `Row`, `Column`, `Line` (the column for a vertical range, the
row for a horizontal one) and `InRepeatedRow`.

:::tip Report, do not throw
A tag that hits a template mistake should add to `context.Errors` and return, the way the built-in
tags do. Throwing from a tag aborts the whole report; recording costs the tag and nothing else.
:::
