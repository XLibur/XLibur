# Formula conformance: pilot results

Spec 59, Task 0. 25 functions, 549 cases, recorded from Excel 16.0.20430 on 2026-11-03 and compared with
XLibur on the same day. How the corpus works and how to refresh it is in
`tools/excel-conformance/README.md`.

## Result

XLibur disagrees with Excel on 20 of 549 cases (3.6%). 16 of the 25 functions have at least one
disagreement. The 20 fall into 11 distinct causes. The sample is weighted toward edge cases on purpose
(coercion, errors, boundaries, wrong argument counts), so this overstates the rate for ordinary use.

The first run found 24 disagreements in 13 causes. The three `MOD` cases (`MOD(1E15,7)`, `MOD(1E20,7)`,
`MOD(3,1E-20)`) were fixed on `main` by #889 while this work was in review, and the runner reported their
known-difference rows as stale. That is the mechanism working as designed.

A fourth case, `COUNTIF(1,1)`, stopped counting when the comparison rules changed: Excel refuses that formula
at entry, and a refused formula now matches XLibur returning any error. A workbook Excel wrote cannot contain
such a formula, so an error value is the same outcome.

The spec's gate sets two outcomes: under about 2%, stop at Task 3 scope; far higher, prioritise Task 4. This
result is between them. It is high enough that the unit tests were agreeing with the code and not with
Excel in several places, and each cause below was invisible to the existing suite. It is not so high that
the cost of the full corpus is in doubt, since it runs 565 tests in about 1.6 seconds.

## Results table

| Function | Family | Cases run | Disagreements |
|---|---|---|---|
| SUM | math | 20 | 2 |
| PRODUCT | math | 17 | 1 |
| ROUND | math | 20 | 1 |
| MOD | math | 20 | 0 |
| INT | math | 20 | 0 |
| POWER | math | 20 | 1 |
| ABS | math | 20 | 0 |
| AVERAGE | statistical | 20 | 1 |
| COUNTIF | statistical | 23 | 0 |
| MEDIAN | statistical | 20 | 1 |
| LEFT | text | 20 | 1 |
| MID | text | 20 | 1 |
| LEN | text | 20 | 4 |
| FIND | text | 23 | 0 |
| SUBSTITUTE | text | 20 | 0 |
| VALUE | text | 28 | 1 |
| TRIM | text | 20 | 1 |
| IF | logical | 23 | 0 |
| AND | logical | 20 | 1 |
| IFERROR | logical | 20 | 1 |
| ISNUMBER | information | 20 | 1 |
| N | information | 20 | 0 |
| VLOOKUP | lookup | 28 | 1 |
| MATCH | lookup | 29 | 0 |
| DATE | date | 38 | 1 |
| **Total** | | **549** | **20** |

Recorded answer types: 271 number, 126 error, 86 text, 33 bool, 33 rejected (Excel refused the formula).

## Disagreements by cause

Ranked by number of cases. Every case is a row in `known-diffs.tsv` with its reason.

| Cause | Cases | Examples |
|---|---|---|
| Double overflow: Excel returns `#NUM!`, XLibur throws `ArgumentException` or passes infinity through | 5 | `SUM(9E307,9E307)`, `PRODUCT(1E200,1E200)`, `IFERROR(9E307*10,"x")`, `ISNUMBER(9E307*10)` |
| Number to text: Excel uses 15 significant digits in plain decimal form | 5 | `LEN(1/3)` is 17 in Excel and 18 in XLibur; `LEN(1E-5)` is 7 and 5; `TRIM(0.1+0.2)` is `0.3` and `0.30000000000000004` |
| Surrogate pair counted as one character by Excel | 2 | `LEN("😀")` is 1 and 2; `MID("😀ab",2,1)` is `a` and half an emoji. XLibur's `LEFT` and `FIND` already agree with Excel |
| `ROUND(1.005,2)` | 1 | 1.01 in Excel, 1 in XLibur |
| `SUM(0.1,0.2)-0.3` | 1 | exactly 0 in Excel, 5.55E-17 in XLibur |
| `MEDIAN(9E307,9E307)` | 1 | 9E307 in Excel, `ArgumentException` in XLibur |
| `POWER(-8,1/3)` | 1 | -1.9999999999999998 in Excel, `#NUM!` in XLibur |
| `VALUE("5-")` | 1 | `#VALUE!` in Excel, -5 in XLibur |
| `AND("")` | 1 | `#VALUE!` in Excel, TRUE in XLibur |
| `DATE(1900,0,31)` | 1 | `#NUM!` in Excel, 0 in XLibur |
| `VLOOKUP` approximate on unsorted data | 1 | `c` in Excel, `#N/A` in XLibur |

By family: math 8 of 137 cases, text 8 of 151, statistical 3 of 63, logical 2 of 63, information 1 of 40,
lookup 1 of 57, date 1 of 38.

## Task 3: generated cases for math, text, logical and information

`tools/excel-conformance/New-CaseFile.ps1` expanded 56 more functions into 1,517 cases, recorded from the
same Excel build. XLibur disagrees on 46 (3.0%), recorded in `known-diffs.tsv`. The whole conformance class
(2,082 tests including the pilot) runs in about 2.5 s. The expansion is mechanical, so most cases are
cheap and most agree. The disagreements come mostly from the argument variants.

The first run found 52. Two `MROUND` cases (`MROUND(0,3)` and `MROUND` of an empty cell) were fixed on `main`
by #885 while this was in review, and the runner reported their rows as stale.

Four `SIN`/`COS` cases failed CI on Linux: XLibur's answer for `COS(1E15)` is a number that differs between
Windows and Linux, so a known-difference row pinning it could not be stable. Excel returns `#NUM!` for an
angle of 2^27 (134,217,728) or more and a value just below it, for `SIN`, `COS` and `TAN` alike (measured
over COM). `TAN` already had that cutoff; `SIN` and `COS` now do too, and the four rows are gone.

| Cause | Cases | Examples |
|---|---|---|
| Array constant as an argument: XLibur throws `UnsupportedFeatureException`, Excel uses the first element | 15 | `LOWER({1,2,3})`, `CONCATENATE("a",{1,2,3})`, `REPT({1,2,3},3)` |
| `FACTDOUBLE`, `MROUND`, `QUOTIENT` take a boolean or multi-cell range: Excel `#VALUE!`, XLibur coerces | 10 | `QUOTIENT(7,TRUE)`, `MROUND(A1:A3,3)` |
| Huge value or digits count: XLibur throws `ArgumentException` | 7 | `ROUNDUP(9E307,2)`, `ROUNDDOWN(2.345,1000000)`, `DEGREES(9E307)` |
| `OR` with a direct text argument: Excel `#VALUE!`, XLibur FALSE | 4 | `OR("x")`, `OR("")` |
| `GCD`/`LCM` with an empty cell reference | 3 | `GCD(A1)` is `#VALUE!` in Excel and 0 in XLibur |
| Text or fractional digits argument: XLibur returns float noise | 3 | `ROUNDUP(2.345,"5")` is 2.34501, `TRUNC(8.9,2.5)` is 8.898649335713818 |
| `ROMAN(1999,TRUE)`, `REPLACE` with an emoji, `SUMPRODUCT` with an error, `TYPE` of a range | 1 each | |

By family: math 25 of 890, text 16 of 383, logical 4 of 50, information 1 of 194. The information functions
agree almost everywhere. The text functions fail on arrays. Several causes share the pilot's themes
(overflow, surrogate pairs, text direct arguments) and could join #894.

## Task 5: date, lookup, statistical, financial and engineering

74 more functions, 2,744 cases, same Excel build. XLibur disagrees on 306 (11.2%), in 48 of the 74 functions.
This is more than three times the rate of the first four families. The whole conformance class (about 4,830
tests) runs in about 4.8 s.

The first run found 359. Fifty-three were formulas Excel refuses at entry where XLibur returned an error value.
Once a refused formula matched any error, those stopped counting.

| Family | Cases | Disagreements |
|---|---|---|
| engineering | 463 | 109 |
| financial | 923 | 95 |
| statistical | 617 | 52 |
| lookup | 149 | 19 |
| date | 592 | 31 |

**Three functions are not implemented.** `ADDRESS`, `CHOOSE` and `LOOKUP` return `#NAME?` for every call
(their registrations in `Lookup.cs` are commented out). 134 generated cases measured only that, so they were
removed from the corpus and the signature table, and are listed here instead. They belong to spec 07's
coverage count and need no known-difference rows.

How the 306 split, by the rule that classified them:

| Mechanism | Cases | Examples |
|---|---|---|
| Boolean or multi-cell range where a number is expected: Excel `#VALUE!`, XLibur coerces | 61 | `WEEKNUM(TRUE)`, `EDATE(43831,TRUE)`, `DELTA(5,TRUE)`, `ERF(1,A1:A3)` |
| Excel refuses the formula at entry, XLibur evaluates it to a value | 1 | `RANK(3,{1,2,3})` is 1 in XLibur |
| Two error arguments: a different argument's error wins | 28 | `GESTEP(#N/A,#DIV/0!)`, `SMALL(#N/A,#DIV/0!)` |
| Huge magnitude or count: XLibur throws `ArgumentException` | 27 | `PMT(0.05,1000000,1000)`, `ERF(9E307)`, `PV(-1,10,-100)` |
| Last digits differ: RATE and IRR converge differently | 22 | `RATE(10,-100,800)` is 0.042774978035587 in Excel and 0.0427749780351 in XLibur |
| Array constant as an argument: `UnsupportedFeatureException` | 9 | `HEX2BIN({1,2,3})`, `CONVERT(1,{1,2,3},"ft")` |
| Other, each recorded as the pair of answers | 158 | below |

Largest groups inside "other":

- **Engineering base conversion.** `BIN2DEC`, `HEX2DEC`, `OCT2DEC` return 0 for an empty cell or empty string in
  Excel and `#NUM!` in XLibur. `DEC2HEX(255,A1)` with an empty cell for the digits argument is `#NUM!` in Excel
  and `FF` in XLibur. `BIN2DEC("-1.5")` and similar throw `ArgumentException`.
- **Statistical text coercion.** `RANK` with a cell holding numeric text is `#VALUE!` in XLibur and a rank in
  Excel. `SMALL(A1:A5,2.5)` is 1 in Excel and 3 in XLibur. `CORREL` and `SLOPE` swap `#N/A` and `#VALUE!` for
  an empty cell, a text argument or an empty string. `MODE` differs in how it treats a text direct argument.
- **Financial edge cases.** `PMT(1E-10,10,1000)` is -100.000000055 in Excel and -99.9999918 in XLibur, which
  looks like cancellation in the closed form for a tiny rate. `NPER` and `PV` differ on a negative or huge rate.
- **Lookup.** `INDEX(A1:B3,2)` is `#REF!` in Excel and `#VALUE!` in XLibur. `HLOOKUP` with a text or empty
  table is `#VALUE!` in Excel and `#N/A` in XLibur.
- **Others.** `YEARFRAC` basis 0 differs near the 1900 leap-day boundary. `TIME(12,-1,0)` is 0.4993 in Excel and
  `#NUM!` in XLibur. `WEEKDAY(43831,A1)` with an empty return type is `#NUM!` in Excel and 4 in XLibur.

These rows are observations, not diagnoses. The reasons in `known-diffs.tsv` for the 158 name the two answers
and nothing more. Task 4 is the step that turns them into causes and fixes.

## Task 4: triage and first fixes

Across the three recordings the corpus has 4,810 cases for 164 functions. Before the fixes below, XLibur
disagreed on 372. After them it disagrees on 329 (6.8%), in 78 of the 164 functions.

Ranked by family, after the fixes:

| Family | Disagreements | Rate |
|---|---|---|
| engineering | 93 of 463 | 20.1% |
| lookup | 20 of 206 | 9.7% |
| financial | 84 of 923 | 9.1% |
| statistical | 52 of 680 | 7.6% |
| date | 32 of 630 | 5.1% |
| text | 24 of 534 | 4.5% |
| logical | 5 of 113 | 4.4% |
| math | 18 of 1,027 | 1.8% |
| information | 1 of 234 | 0.4% |

The functions with most disagreements: `RATE` 28, `INDEX` 15, `SLOPE` 14, `CORREL` 14, `YEARFRAC` 10,
`HEX2BIN` 9, `BIN2HEX` 9, `RANK` 9.

**Fixed here (43 cases removed, with regression tests in `ExcelConformanceFixesTests`).**

- A non-finite number is `#NUM!`. `ScalarValue.From(double)` now turns NaN and infinity into `#NUM!`, which is
  what Excel has (it holds no infinity). Before, `SUM(9E307,9E307)` and every other overflow threw
  `ArgumentException` from `XLCellValue`, and `ISNUMBER(9E307*10)` was TRUE. One change at the single factory
  cleared overflows in `SUM`, `PRODUCT`, `AVERAGE`, `PMT`, `FV`, `ERF` and others.
- `ROUNDUP`, `ROUNDDOWN` and `TRUNC` work on the 15-significant-digit decimal form, like Excel. They
  returned `2.34501` for `ROUNDUP(2.345,"5")` and `8.898649335713818` for `TRUNC(8.9,2.5)`, and they threw for
  a huge value or a digits count of 1000000, where Excel returns the value unchanged.
- `BIN2DEC`, `HEX2DEC`, `OCT2DEC`, `BIN2HEX`, `HEX2BIN` and the other base conversions: an empty string reads
  as 0, a string of only spaces is `#NUM!`, and a minus sign is `#NUM!` and no longer throws.
- `RADIANS` and `MEDIAN` no longer overflow in an intermediate step when the result fits.

Ten rows stayed as differences with a new XLibur answer: they returned a crash before and `#NUM!` now,
where Excel returns a number (for example `PMT(0.05,1000000,1000)` is -50 in Excel).

**Remaining, in the order I would take them.**

1. A boolean or multi-cell range argument where Excel returns `#VALUE!` (71 cases). Mostly the Analysis
   ToolPak functions. A shared argument adapter would clear most of it.
2. An array constant as an argument throws `UnsupportedFeatureException` (24).
3. Which of two error arguments wins (28).
4. `RATE` and `IRR` convergence (22). Excel and XLibur stop at different digits; this needs a decision on
   tolerance and iteration count.
5. The 150 rows that only state the two answers. The biggest are `INDEX` (15), `SLOPE` and `CORREL` (26),
   `YEARFRAC` (10) and the engineering conversions.
6. Singletons: surrogate pairs (3), number to text (5), `OR` with a text argument (4), `GCD`/`LCM` with an empty cell (3), `ROUND(1.005,2)`, `VALUE("5-")`, `AND("")`,
   `DATE(1900,0,31)`, `ROMAN` with a boolean form, `SUMPRODUCT` with an error, `TYPE` of a range, the
   unsorted approximate `VLOOKUP`.

## What the pilot does not cover

- Array and spilling results. The recorder enters formulas with `Range.Formula`, which does not spill, and
  XLibur's public `Evaluate` returns one value. Arrays need the hash form the spec describes.
- Only one Excel build, and one machine locale (en-AU, with `.` and `,` separators).
- Fixing any of the 24. Each is its own change; none was attempted here and no issues are filed yet.

## Findings about the method

- The recorder is deterministic: re-running it against the same Excel reproduced the committed file byte
  for byte.
- Excel refuses a formula holding a literal above about 9.99E307. The first recording measured that limit
  and not the function, so the 13 occurrences of `1E308` were changed to `9E307` and re-recorded.
- XLibur's parser refuses a formula when the value is read, not when the formula is set, so the runner maps
  `ExpressionParseException` from either step to `rejected`.
