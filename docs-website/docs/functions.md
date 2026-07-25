---
id: functions
title: Functions
sidebar_label: Functions
description: Using Excel functions in XLibur formulas, and the full list of functions the built-in calculation engine can evaluate.
---

# Functions

Any function name you write into a formula is stored in the file and evaluated by Excel when
the workbook is opened — XLibur does not restrict what you can write. The list on this page is
about a narrower question: which functions **XLibur's own calculation engine** can evaluate,
so that reading `cell.Value` (or saving with `evaluateFormulae: true`) produces a result.

## Using a function

Functions are just formula text:

```csharp
using XLibur.Excel;

var ws = workbook.Worksheet("Data");

ws.Cell("A1").Value = "Ada";
ws.Cell("B1").Value = "Lovelace";

ws.Cell("C1").FormulaA1 = "=CONCAT(A1, \" \", B1)";
Console.WriteLine(ws.Cell("C1").GetString());     // "Ada Lovelace"
```

:::tip
Formula strings are just C# strings, so `"` inside them must be escaped. A raw string literal
avoids the backslashes entirely and is much easier to read for anything non-trivial:

```csharp
ws.Cell("C1").FormulaA1 = """=CONCAT(A1, " ", B1)""";
ws.Cell("C2").FormulaA1 = """=IF(B2>100, "High", "Low")""";
```
:::

### Text

```csharp
ws.Cell("D1").FormulaA1 = """=TEXTJOIN(", ", TRUE, A1:A10)""";
ws.Cell("D2").FormulaA1 = """=UPPER(TRIM(A1))""";
ws.Cell("D3").FormulaA1 = """=LEFT(A1, FIND(" ", A1) - 1)""";
ws.Cell("D4").FormulaA1 = """=SUBSTITUTE(A1, "-", "/")""";
ws.Cell("D5").FormulaA1 = """=TEXT(B1, "$ #,##0.00")""";
```

### Lookup

```csharp
ws.Cell("E1").FormulaA1 = "=VLOOKUP(A1, Products!A:C, 3, FALSE)";
ws.Cell("E2").FormulaA1 = "=INDEX(C:C, MATCH(A1, A:A, 0))";
ws.Cell("E3").FormulaA1 = """=XLOOKUP(A1, Products[Code], Products[Price], "not found")""";
```

### Conditional aggregation

```csharp
ws.Cell("F1").FormulaA1 = """=SUMIF(B:B, ">100", C:C)""";
ws.Cell("F2").FormulaA1 = """=SUMIFS(C:C, B:B, ">100", A:A, "North")""";
ws.Cell("F3").FormulaA1 = """=COUNTIF(A:A, "North")""";
ws.Cell("F4").FormulaA1 = """=AVERAGEIF(B:B, ">0")""";
```

### Logic and error handling

```csharp
ws.Cell("G1").FormulaA1 = """=IF(B1 > 100, "High", "Low")""";
ws.Cell("G2").FormulaA1 = """=IFS(B1>1000, "A", B1>100, "B", TRUE, "C")""";
ws.Cell("G3").FormulaA1 = """=IFERROR(A1/B1, 0)""";
ws.Cell("G4").FormulaA1 = """=SWITCH(A1, "N", "North", "S", "South", "Other")""";
```

### Dates

```csharp
ws.Cell("H1").FormulaA1 = "=TODAY()";
ws.Cell("H2").FormulaA1 = "=EOMONTH(A1, 0)";
ws.Cell("H3").FormulaA1 = "=NETWORKDAYS(A1, B1)";
ws.Cell("H4").FormulaA1 = """=DATEDIF(A1, TODAY(), "y")""";
```

### Dynamic arrays

These spill across multiple cells, so they need `SetDynamicFormulaA1` — see
[Formulas](./formulas.md#dynamic-array-formulas):

```csharp
ws.Cell("I1").SetDynamicFormulaA1("=SORT(UNIQUE(A2:A100))");
ws.Cell("J1").SetDynamicFormulaA1("=FILTER(A2:C100, C2:C100>1000)");
ws.Cell("K1").SetDynamicFormulaA1("=SEQUENCE(12, 1, 1, 1)");
```

## Evaluating without a workbook cell

`Evaluate` runs an expression and returns the result directly — handy for checking a formula,
or for using Excel's function library as a calculator:

```csharp
var total = workbook.Evaluate("=SUM(Data!A1:A10)");
var text = workbook.Evaluate("""=CONCAT("Hello", " ", "World")""");
var rounded = workbook.Evaluate("=ROUND(PI() * 2, 4)");

// Sheet-scoped: unqualified references resolve against this sheet
var localTotal = ws.Evaluate("=SUM(A1:A10)");
```

## When a function isn't supported

If the engine does not recognise a function, evaluation raises an error rather than returning
a value. The formula itself is still written to the file correctly, and Excel will compute it
on open — the only thing you lose is XLibur's ability to give you the answer:

```csharp
ws.Cell("A1").FormulaA1 = "=BAHTTEXT(1234)";   // written fine, not evaluable by XLibur

try
{
    var value = ws.Cell("A1").Value;
}
catch (Exception ex)
{
    // Not in the supported list below — let Excel calculate it
}
```

Save without evaluation (the default) when a workbook uses functions outside the list:

```csharp
workbook.SaveAs("Report.xlsx");   // evaluateFormulae defaults to false
```

## Supported functions

The following functions are implemented by XLibur's calculation engine. Function names are
case-insensitive in formulas.

### Math and trigonometry

`ABS` · `ACOS` · `ACOSH` · `ACOT` · `ACOTH` · `ARABIC` · `ASIN` · `ASINH` · `ATAN` · `ATAN2` ·
`ATANH` · `BASE` · `CEILING` · `CEILING.MATH` · `COMBIN` · `COMBINA` · `COS` · `COSH` · `COT` ·
`COTH` · `CSC` · `CSCH` · `DECIMAL` · `DEGREES` · `EVEN` · `EXP` · `FACT` · `FACTDOUBLE` ·
`FLOOR` · `FLOOR.MATH` · `GCD` · `INT` · `LCM` · `LN` · `LOG` · `LOG10` · `MDETERM` ·
`MINVERSE` · `MMULT` · `MOD` · `MROUND` · `MULTINOMIAL` · `ODD` · `PI` · `POWER` · `PRODUCT` ·
`QUOTIENT` · `RADIANS` · `RAND` · `RANDBETWEEN` · `ROMAN` · `ROUND` · `ROUNDDOWN` · `ROUNDUP` ·
`SEC` · `SECH` · `SERIESSUM` · `SIGN` · `SIN` · `SINH` · `SQRT` · `SQRTPI` · `SUBTOTAL` ·
`SUM` · `SUMIF` · `SUMIFS` · `SUMPRODUCT` · `SUMSQ` · `SUMX2MY2` · `SUMX2PY2` · `SUMXMY2` ·
`TAN` · `TANH` · `TRUNC`

### Statistical

`AVEDEV` · `AVERAGE` · `AVERAGEA` · `AVERAGEIF` · `AVERAGEIFS` · `BINOM.DIST` · `BINOMDIST` ·
`COUNT` · `COUNTA` · `COUNTBLANK` · `COUNTIF` · `COUNTIFS` · `DEVSQ` · `FISHER` · `GEOMEAN` ·
`LARGE` · `MAX` · `MAXA` · `MAXIFS` · `MEDIAN` · `MIN` · `MINA` · `MINIFS` · `MODE` ·
`MODE.SNGL` · `PERCENTILE` · `PERCENTILE.INC` · `QUARTILE` · `QUARTILE.INC` · `RANK` ·
`RANK.EQ` · `SMALL` · `STDEV` · `STDEV.P` · `STDEV.S` · `STDEVA` · `STDEVP` · `STDEVPA` ·
`T.INV` · `T.INV.2T` · `TINV` · `VAR` · `VAR.P` · `VAR.S` · `VARA` · `VARP` · `VARPA`

### Text

`ASC` · `CHAR` · `CLEAN` · `CODE` · `CONCAT` · `CONCATENATE` · `DOLLAR` · `EXACT` · `FIND` ·
`FIXED` · `LEFT` · `LEFTB` · `LEN` · `LOWER` · `MID` · `NUMBERVALUE` · `PROPER` · `REPLACE` ·
`REPT` · `RIGHT` · `SEARCH` · `SUBSTITUTE` · `T` · `TEXT` · `TEXTJOIN` · `TRIM` · `UPPER` ·
`VALUE`

### Logical

`AND` · `FALSE` · `IF` · `IFERROR` · `IFS` · `NOT` · `OR` · `SWITCH` · `TRUE`

### Lookup and reference

`ADDRESS` · `AREAS` · `CHOOSE` · `COLUMN` · `COLUMNS` · `FORMULATEXT` · `GETPIVOTDATA` ·
`HLOOKUP` · `HYPERLINK` · `INDEX` · `INDIRECT` · `LOOKUP` · `MATCH` · `OFFSET` · `ROW` ·
`ROWS` · `RTD` · `TRANSPOSE` · `VLOOKUP`

### Dynamic array

`FILTER` · `SEQUENCE` · `SORT` · `SORTBY` · `UNIQUE` · `XLOOKUP` · `XMATCH`

### Date and time

`DATE` · `DATEDIF` · `DATEVALUE` · `DAY` · `DAYS` · `DAYS360` · `EDATE` · `EOMONTH` · `HOUR` ·
`ISOWEEKNUM` · `MINUTE` · `MONTH` · `NETWORKDAYS` · `NOW` · `SECOND` · `TIME` · `TIMEVALUE` ·
`TODAY` · `WEEKDAY` · `WEEKNUM` · `WORKDAY` · `YEAR` · `YEARFRAC`

### Financial

`FV` · `IPMT` · `IRR` · `NPER` · `NPV` · `PMT` · `PPMT` · `PV` · `RATE`

### Information

`CELL` · `ERROR.TYPE` · `INFO` · `ISBLANK` · `ISERR` · `ISERROR` · `ISEVEN` · `ISLOGICAL` ·
`ISNA` · `ISNONTEXT` · `ISNUMBER` · `ISODD` · `ISREF` · `ISTEXT` · `N` · `NA` · `TYPE`

### Engineering

`BIN2DEC` · `BIN2HEX` · `BIN2OCT` · `DEC2BIN` · `DEC2HEX` · `DEC2OCT` · `HEX2BIN` · `HEX2DEC` ·
`HEX2OCT` · `OCT2BIN` · `OCT2DEC` · `OCT2HEX`

### Database

`DAVERAGE` · `DCOUNT` · `DCOUNTA` · `DGET` · `DMAX` · `DMIN` · `DPRODUCT` · `DSTDEV` ·
`DSTDEVP` · `DSUM` · `DVAR` · `DVARP`

## Where to next

- [Formulas](./formulas.md) — writing, evaluating, and clearing formulas
- [Tables](./tables.md) — structured references like `SalesTable[Amount]`
