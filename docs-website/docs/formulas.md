---
id: formulas
title: Formulas
sidebar_label: Formulas
description: Write normal, array, and dynamic array formulas in A1 or R1C1 notation, evaluate them, and clear them.
---

# Formulas

XLibur can both *write* formulas into a workbook and *evaluate* them with its own calculation
engine, so you can read a computed result without opening Excel.

There are four kinds of formula in a worksheet:

| Kind | How you set it | Notes |
|---|---|---|
| Normal | `cell.FormulaA1` / `cell.FormulaR1C1` | One cell, one result |
| Array | `range.FormulaArrayA1` | Legacy CSE formula over a fixed range |
| Dynamic array | `cell.SetDynamicFormulaA1(...)` | Excel 365 spilling formula |
| Data table | *(read-only)* | What-if tables — preserved on round-trip |

## Normal formulas

Assign to `FormulaA1`. The leading `=` is optional:

```csharp
using XLibur.Excel;

var ws = workbook.Worksheet("Data");

ws.Cell("A2").Value = 1;
ws.Cell("B2").Value = 2;

ws.Cell("C2").FormulaA1 = "=A2+$B$2";
ws.Cell("C3").FormulaA1 = "SUM(A2:B2)";        // the = is optional
ws.Cell("C4").SetFormulaA1("=AVERAGE(A2:B2)"); // fluent form, returns the cell
```

### R1C1 notation

`FormulaR1C1` is the relative-offset notation. `RC[-2]` means "same row, two columns left";
`R3C2` is an absolute reference to `B3`:

```csharp
ws.Cell("C3").FormulaR1C1 = "RC[-2]+R3C2";
ws.Cell("C5").FormulaR1C1 = "=SUM(R[-3]:R[-1])";
```

Both properties address the *same* underlying formula, so you can set one and read the other:

```csharp
var cell = ws.Cell("C2");
cell.FormulaA1 = "=A2+B2";

Console.WriteLine(cell.FormulaA1);     // "A2+B2"
Console.WriteLine(cell.FormulaR1C1);   // "RC[-2]+RC[-1]"
```

To change what Excel *displays* in its UI, set the workbook reference style:

```csharp
workbook.ReferenceStyle = XLReferenceStyle.R1C1;   // or A1, Default
```

### Filling a formula down a range

Setting `FormulaR1C1` on a range writes the formula into every cell, with relative references
resolved per cell — the equivalent of dragging the fill handle:

```csharp
// Every cell in D2:D100 gets "=B{row}*C{row}"
ws.Range("D2:D100").FormulaR1C1 = "=RC[-2]*RC[-1]";
```

With A1 notation you build the string per row instead:

```csharp
for (var row = 2; row <= 100; row++)
{
    ws.Cell(row, 4).FormulaA1 = $"=B{row}*C{row}";
}
```

Or copy a cell, which shifts its references:

```csharp
var seed = ws.Cell("D2");
seed.FormulaA1 = "=B2*C2";
seed.CopyTo(ws.Cell("D3"));    // becomes "=B3*C3"
```

### Formulas over named ranges and tables

```csharp
ws.Range("A2:A10").AddToNamed("SalesFigures");
ws.Cell("C1").FormulaA1 = "=SUM(SalesFigures)";

// Structured table references
ws.Cell("C2").FormulaA1 = "=SUM(SalesTable[Amount])";
ws.Cell("C3").FormulaA1 = "=SUMIF(SalesTable[Region], \"North\", SalesTable[Amount])";

// Cross-sheet
ws.Cell("C4").FormulaA1 = "=Summary!B2";
ws.Cell("C5").FormulaA1 = "='Q1 Sales'!B2";   // quote names containing spaces
```

## Array formulas

A legacy array (CSE) formula occupies a fixed range and produces one result per cell in it.
Set `FormulaArrayA1` on the target range — no curly braces, XLibur adds them:

```csharp
// Single-cell array formula
ws.Range("B6").FormulaArrayA1 = "A2+A3";

// Multi-cell: transpose A2:A3 into two horizontal cells
ws.Range("C6:D6").FormulaArrayA1 = "TRANSPOSE(A2:A3)";

// The classic SUMPRODUCT-style aggregate
ws.Range("F1").FormulaArrayA1 = "SUM(IF(A2:A100>100, B2:B100, 0))";
```

The range you set it on *is* the extent of the array — size it to match the result the formula
produces, exactly as you would with Ctrl+Shift+Enter in Excel.

```csharp
var cell = ws.Cell("C6");
Console.WriteLine(cell.HasArrayFormula);   // true
Console.WriteLine(cell.FormulaReference);  // the array's range address
```

## Dynamic array formulas

Excel 365 functions such as `FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`, and `XLOOKUP` *spill*:
they return a result of whatever size the data implies, filling the cells below and to the
right. These need `SetDynamicFormulaA1` rather than `FormulaA1`, so that Excel does not prepend
the implicit-intersection operator `@`:

```csharp
ws.Cell("E1").SetDynamicFormulaA1("=SORT(UNIQUE(A2:A100))");
ws.Cell("F1").SetDynamicFormulaA1("=FILTER(A2:C100, C2:C100>1000)");
ws.Cell("G1").SetDynamicFormulaA1("=SEQUENCE(10, 1, 1, 5)");
```

:::warning
Using plain `FormulaA1` for a dynamic array function writes `=@FILTER(...)`, which Excel
interprets as a single-cell intersection and will not spill. Use `SetDynamicFormulaA1` for any
of the functions listed under [Dynamic array](./functions.md#dynamic-array) on the Functions
page.
:::

## Implicit intersection

A legacy formula that applies an *operator* to a range intersects that range against its own cell,
the way Excel does. With `A1 = 42`, `B1 = 100` and `B3 = 5`, a formula in `C3`:

```csharp
ws.Cell("C3").FormulaA1 = "=A1+B1:B3";
var value = ws.Cell("C3").Value;   // 47 — that is A1 + B3, the cell of B1:B3 on row 3
```

`B1:B3` is reduced to the one cell sharing the formula's row before the `+` runs. That is what
Excel shows for the very file XLibur writes, because Excel reads the stored `A1+B1:B3` back as
`A1+@B1:B3`. It applies to every operator kind:

| Formula in `C3` | Result |
|---|---|
| `=A1*B:B` | `210` |
| `=A1&B1:B3` | `425` |
| `=A1>B1:B3` | `TRUE` |
| `=-B1:B3` | `-5` |
| `=B1:B3%` | `0.05` |

A range that spans neither the formula's row nor its column is `#VALUE!`, again matching Excel.

Three things deliberately do **not** intersect:

- **A dynamic-array formula still spills.** `SetDynamicFormulaA1("A1+B1:B3")` gives `142, 49, 47`.
- **`Evaluate(expression)` with no address keeps array semantics** — with no cell, there is nothing
  to intersect against. Pass an address (`ws.Evaluate(expr, "C3")`) to get the cell behaviour.
- **An operator inside a function argument does not intersect**, so `MIN(A1:A2-B1)` and
  `SUMPRODUCT((A1:C2-E1:G2)^2/E1:G2)` are unaffected.

### The `@` and space operators {#at-and-space-operators}

You can also ask for an intersection yourself. XLibur calculates both operators:

| Formula | What it gives |
|---|---|
| `=@B1:B10` in row 3 | `B3`, the cell of the range on the formula's row |
| `=SUM(B7:D7 C6:C8)` | the sum of `C7`, the cell that both ranges contain |

The `@` operator gives the cell on the formula's row or column. For an array, it gives the first
element. The space operator gives the cells that both references have in common. If they have no
cell in common, the result is `#NULL!`.

:::caution
Formulas of this shape used to answer with the range's *first* element — `142` for the example
above — so a cached value computed by an earlier version can differ from the same formula
recalculated now. That is the point of the fix, but worth knowing before a diff of recalculated
workbooks surprises you. There is no compile-time signal.
:::

## Clearing a formula

Assigning an empty (or whitespace) string removes the formula. The cell keeps whatever value
was last computed:

```csharp
ws.Cell("C2").FormulaA1 = "";      // formula removed, cached value retained
```

To remove the formula *and* the value, clear the contents:

```csharp
ws.Cell("C2").Clear(XLClearOptions.Contents);
```

To replace a formula with its result — "paste values" — read the value first:

```csharp
var cell = ws.Cell("C2");
var result = cell.Value;      // evaluates the formula
cell.FormulaA1 = "";
cell.Value = result;
```

Applied to a whole sheet:

```csharp
foreach (var cell in ws.CellsUsed(c => c.HasFormula).ToList())
{
    var value = cell.Value;
    cell.FormulaA1 = "";
    cell.Value = value;
}
```

Checking before you act:

```csharp
if (cell.HasFormula)
{
    Console.WriteLine(cell.FormulaA1);
}
```

## Evaluating formulas

Reading `Value` on a formula cell evaluates it — XLibur's calculation engine runs the formula
and caches the result:

```csharp
ws.Cell("A1").Value = 10;
ws.Cell("A2").Value = 32;
ws.Cell("A3").FormulaA1 = "=SUM(A1:A2)";

Console.WriteLine(ws.Cell("A3").Value);       // 42 — evaluated on demand
Console.WriteLine(ws.Cell("A3").GetDouble()); // 42
```

`CachedValue` returns the stored result *without* triggering a recalculation. It may be stale —
check `NeedsRecalculation` first:

```csharp
var cell = ws.Cell("A3");

if (!cell.NeedsRecalculation)
{
    Console.WriteLine(cell.CachedValue);
}

cell.InvalidateFormula();       // force re-evaluation on next read
```

Recalculating in bulk:

```csharp
ws.RecalculateAllFormulas();
workbook.RecalculateAllFormulas();
```

### When a formula cannot be calculated

Some formulas have no answer XLibur can give. There are three kinds:

- **A circular reference.** The formula depends on its own value, directly or through other
  formulas or defined names.
- **A formula XLibur does not evaluate.** For example, a DDE link such as
  `Sdemo123|tik!'id1?req?AAPL'`.
- **A formula the parser cannot read.** For example, `'[Book2.xlsx]Sheet1'!A1`, or a function
  called with the wrong number of arguments, such as `ABS(1,2)`.

Reading such a cell throws. Each kind has its own exception, so you can tell them apart from a
bug in XLibur:

```csharp
using XLibur.Excel.CalcEngine;
using XLibur.Excel.CalcEngine.Exceptions;

try
{
    var value = ws.Cell("A1").Value;
}
catch (XLCircularReferenceException)
{
    // the cell is in a cycle, or depends on one
}
catch (NotImplementedException)
{
    // the formula uses something XLibur does not evaluate
}
catch (ExpressionParseException)
{
    // the parser cannot read the formula
}
```

`XLCircularReferenceException` derives from `InvalidOperationException`, so older code that
catches that type still works.

The problem stays with that cell and the cells that depend on it. The rest of the workbook still
works:

- You can still edit any cell, and reading a cell that does not depend on the problem cell
  gives its value.
- `RecalculateAllFormulas` (and `LoadOptions.RecalculateAllFormulas`) does not throw. It
  calculates every other formula and leaves the problem cells with `NeedsRecalculation` set to
  `true`.

So, to find the problem cells after a recalculation, check `NeedsRecalculation`:

```csharp
workbook.RecalculateAllFormulas();

var notCalculated = ws.CellsUsed(c => c.HasFormula && c.NeedsRecalculation).ToList();
```

:::caution
In earlier versions `RecalculateAllFormulas` threw at the first circular reference or unsupported
formula. If your code used that exception to find such a cell, it no longer sees one. Check
`NeedsRecalculation` instead.
:::

### Evaluating an expression directly

You can run a formula without writing it into a cell:

```csharp
var value = workbook.Evaluate("=SUM(Data!A1:A10)");

// Sheet-scoped, so unqualified references resolve against that sheet
var local = ws.Evaluate("=SUM(A1:A10)");

// With an address so relative references have an anchor
var relative = ws.Evaluate("=A1+B1", "C1");
```

An expression that needs to know *which cell* it is being evaluated in — `ROW()`, `COLUMN()`, or
anything reaching implicit intersection such as `VLOOKUP(A1:B341,,1,FALSE)` — throws
`XLNoWorksheetContextException` when no address was supplied:

```csharp
try
{
    var row = ws.Evaluate("=ROW()");
}
catch (XLNoWorksheetContextException)
{
    var row = ws.Evaluate("=ROW()", "B7");   // give it a cell to be relative to
}
```

:::caution
This used to be an `InvalidOperationException` — an internal exception type a caller outside the
assembly could not name, so a broken expression could not be told apart from a library bug. The
new type derives from `XLiburException`, so **a `catch (InvalidOperationException)` around
`Evaluate` no longer runs** and the exception escapes. There is no compile-time signal; catch
`XLNoWorksheetContextException` instead.
:::

### Saving computed values

By default XLibur writes formulas without their results, and Excel computes them on open.
Other consumers — a CSV exporter, a headless parser, LibreOffice in some configurations —
may need the values present in the file. Ask for them at save time:

```csharp
workbook.SaveAs("Report.xlsx", validate: false, evaluateFormulae: true);

// Equivalent, with the options object
workbook.SaveAs("Report.xlsx", new SaveOptions { EvaluateFormulasBeforeSaving = true });
```

:::note
If a formula cannot be calculated — a circular reference, a function XLibur does not implement,
or text the parser cannot read — the save still succeeds. XLibur writes that cell without a
value, and Excel calculates it when it opens the file.

Any other failure is a bug in XLibur, and the save throws it. Earlier versions hid these bugs and
wrote the cell without a value.
:::

### Calculation mode

```csharp
workbook.CalculateMode = XLCalculateMode.Auto;         // Excel recalculates automatically
workbook.CalculateMode = XLCalculateMode.Manual;       // user presses F9
workbook.CalculateMode = XLCalculateMode.AutoNoTable;  // auto, except data tables
```

## Formulas and structural edits

Inserting or deleting rows and columns rewrites affected references across the workbook, and
renaming a sheet rewrites the formulas that mention it — the same behaviour as Excel:

```csharp
ws.Cell("C1").FormulaA1 = "=SUM(B2:B10)";
ws.Row(5).InsertRowsAbove(1);
Console.WriteLine(ws.Cell("C1").FormulaA1);   // "SUM(B2:B11)"

ws.Name = "Renamed";
// A formula elsewhere reading "=Data!A1" now reads "=Renamed!A1"
```

Deleting a sheet also rewrites the formulas that refer to it, as Excel does. Each reference to
the deleted sheet becomes `#REF!`:

```csharp
// Sheet2!A1 holds "=Sheet1!A1*2"
workbook.Worksheet("Sheet1").Delete();
Console.WriteLine(workbook.Worksheet("Sheet2").Cell("A1").FormulaA1);   // "#REF!*2"
```

A 3D reference that has the deleted sheet at one end gets smaller instead. Deleting `Sheet1`
changes `SUM(Sheet1:Sheet3!A1)` to `SUM(Sheet2:Sheet3!A1)`. If you delete a sheet from the middle
of the span, the reference does not change.

Two kinds of text are not rewritten:

- Text inside a string, such as `INDIRECT("Sheet1!A1")`.
- A formula the parser cannot read. XLibur does not know which references it holds, so it keeps
  the text as it is.

## Data tables

Excel's *Data Table* feature (Data → What-If Analysis → Data Table) produces a special
`{=TABLE(row_input, col_input)}` formula. XLibur reads these, preserves them across a
load/save cycle, and keeps their cached values — but there is **no public API to create one**.

If a workbook needs a data table, build it in Excel as a template and let XLibur populate the
input cells around it:

```csharp
using var workbook = new XLWorkbook("WhatIfTemplate.xlsx");
var ws = workbook.Worksheet("Scenarios");

// The data table formula in B2:F10 is preserved; only the inputs change
ws.Cell("B1").Value = 0.05;    // interest rate
ws.Cell("A2").Value = 250_000; // principal

workbook.Save();
```

For scenario grids built entirely in code, generate the cross-product yourself with ordinary
formulas — the result is a plain range Excel and every other reader understands:

```csharp
double[] rates = [0.03, 0.04, 0.05, 0.06];
int[] terms = [10, 15, 20, 25, 30];

ws.Cell("A1").Value = "Principal";
ws.Cell("B1").Value = 250_000;

for (var c = 0; c < rates.Length; c++)
{
    ws.Cell(3, c + 2).Value = rates[c];
}

for (var r = 0; r < terms.Length; r++)
{
    ws.Cell(r + 4, 1).Value = terms[r];

    for (var c = 0; c < rates.Length; c++)
    {
        ws.Cell(r + 4, c + 2).FormulaA1 =
            $"=PMT({ws.Cell(3, c + 2).Address.ToStringFixed()}/12, " +
            $"{ws.Cell(r + 4, 1).Address.ToStringFixed()}*12, $B$1)";
    }
}
```

## Where to next

- [Functions](./functions.md) — what the calculation engine supports, with examples
- [Cells and Ranges](./cells-and-ranges.md) — reading typed results back out
- [Workbook Settings](./workbook-settings.md#calculation) — calculation mode and save options
