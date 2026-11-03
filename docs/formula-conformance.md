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
