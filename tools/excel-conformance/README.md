# Excel formula conformance corpus

Spec 59. The corpus holds formulas and the value Excel computed for each. `FormulaConformanceTests`
checks that XLibur agrees. CI never needs Excel, because Excel's answers are stored in the files.

```
src/XLibur.Tests/Resource/Other/FormulaConformance/
    pilot.tsv         cases and Excel's recorded answers
    known-diffs.tsv   cases XLibur is known to answer differently, each with a reason
tools/excel-conformance/Record-Corpus.ps1   the recorder (Windows, desktop Excel, PowerShell 7)
```

## Case file format

One TSV row per case, seven columns:

| column | meaning |
|---|---|
| function | function under test; the case id is `function#n`, n counting rows of that function in file order |
| family | math, text, logical, information, lookup, statistical, date |
| setup | cells to fill, `A1=5;A2="x";A3=#N/A;A4=TRUE;A5==SUM(1,2)`. Nothing after `=` is an empty cell. Use cells A to J |
| formula | entered in `K1` |
| excel_value | recorded. Numbers as a round-trip string, text with `\\ \t \n \r` escaped (a lone surrogate as `\uXXXX`), `TRUE`/`FALSE`, or an error code |
| excel_type | recorded: `number`, `text`, `bool`, `error`, `empty`, `rejected` (Excel refused the formula at entry) |
| note | free text |

The first lines are `# excel_build:` and `# recorded:`, written by the recorder. The runner rejects a file
without the build stamp and a row without a recorded type.

Ids are positional. Append new cases at the end of a function's rows and never reorder or delete, or the
ids in `known-diffs.tsv` stop matching.

## Add or refresh cases

1. Add rows with empty `excel_value` and `excel_type` (leave the tabs in place), or edit nothing to refresh.
2. Close nothing of yours. The recorder starts its own Excel instance and closes only its own workbook.
3. Run `pwsh tools/excel-conformance/Record-Corpus.ps1`. It fills every row from Excel and re-stamps the file.
4. Run `dotnet test src/XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/FormulaConformanceTests/*"`.
5. For each failure, fix XLibur, or add a row to `known-diffs.tsv` with a reason that states what was observed.

The recorder enters formulas with `Range.Formula` (legacy semantics, no spilling), so array results are not
covered yet. It refuses to run if Excel's decimal and thousands separators are not `.` and `,`.

## Generate cases for a function family

`signatures.psd1` lists each function with its argument kinds, typical values and minimum argument count.
`New-CaseFile.ps1` expands it into `<family>.tsv` with empty answers. No Excel is needed for this step.

Per function it writes the typical call, then each argument replaced by an empty cell, numeric text,
non-numeric text, the empty string, a boolean, an error, a range, an array constant and a cell holding
numeric text. It adds the kind's boundary values, which error wins when two arguments are errors, and calls
with too few or too many arguments. Functions already in an existing file are skipped.

Optional table `S` gives a function the cells every case starts from (lookup and statistical functions); a
case's own setup is written after it. With `S`, the empty-cell variant uses `J10` instead of `A1`. Kind `date`
adds serial-date boundaries: 0, 1, 60, 61, 2958465, 2958466. A function XLibur has not implemented yet (it
returns `#NAME?`) is left out of the table, because its cases would measure only that.

Case ids are positional, so the script refuses to overwrite an existing file. `-Force` discards the recorded
answers. To add a function, add it to the signature table, move the existing family file aside, and put the
new cases in a new file. Then run the recorder on that file.

## Comparison rules

Numbers match within relative 1e-12 (absolute 1e-300 near zero). Text is ordinal. An error must be the same
code. Empty, zero and the empty string are different types. Dates are compared as serial numbers.

One exception: a formula Excel refuses at entry (`rejected`) matches XLibur refusing it or evaluating it to any
error value, because such a formula cannot exist in a workbook Excel wrote. XLibur evaluating it to a number,
text or boolean is still a difference.

A known-differences row fails the run when the case now agrees with Excel (the row is stale), when XLibur's
answer changed, or when Excel's stored answer no longer matches the row.

## Limits found while recording

- Excel refuses a formula holding a literal above about 9.99E307 (`1E308`), so cases use `9E307`.
- Keep locale-dependent text out of cases (`VALUE("1/2/2020")`). The corpus is en-US.
- Volatile and environment functions (`TODAY`, `NOW`, `RAND`, `INFO`, `CELL("filename")`) are not for this corpus.
