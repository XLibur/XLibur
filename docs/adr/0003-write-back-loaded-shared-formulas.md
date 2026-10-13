# A save writes back the shared formulas a load brought in, and finds no new ones

A save writes each shared formula that a load brought in back as a shared formula, limited to the cells that still hold it. It does not look for new shared formulas among the formulas a caller wrote. Until this decision, a save wrote a formula of its own in every cell, so a load and a save of any file from Excel expanded every shared formula (#692).

A cell still holds its shared formula while its formula text and its position are the ones the load gave it (`XLCellFormula.TryGetShared`). Formula text in R1C1 depends on nothing else, so such a cell holds the same formula as every other cell of the group, and the save neither parses nor converts any formula to find out. An edit, a move, a row or column shift, a copy and a rewrite by a sheet rename each drop the group from the cell, and the cell is then written with its own text. Evaluation already relies on the same rule to read each cell of a shared formula through one tree (#686).

Three alternatives were rejected:

- **Write no shared formulas.** The measurements below show that the files are 10–12% smaller, Excel opens them 28% faster, and the next XLibur load recalculates with 24–25% less allocation. The save costs the same.
- **Also find new shared formulas at save.** Excel itself does not do this. It writes a shared formula only for a block it filled or copied, and writes formulas entered one by one as formulas of their own, even when their R1C1 text is the same. To find new groups, XLibur would have to convert every formula to R1C1, which is about one parse per formula (232 B for `SUM(D5:H5)`, #686), to save bytes in files that Excel never writes that way.
- **A public `SaveOptions` switch.** Every reader that was tested reads the output. A save writes back only the groups that its input file already had, so the output is no harder to read than the input. `SaveOptions.WriteSharedFormulas` exists, but it is internal, so that tests and benchmarks can compare the two forms. It can become public if a reader is found that needs it.

## The rules a written shared formula keeps

ECMA-376 Part 1 §18.3.1.40 says only this. The first cell of a group is the master, and it holds the text. A later cell holds `t="shared"` and the `si`, and any text in it is ignored. A cell inside `ref` that is not in the group keeps its own formula. A cell outside `ref` is implementation-defined. The `ref` ranges of one sheet shall not overlap.

What Excel 16 does was measured over COM, with hand-made files opened by `Workbooks.Open`:

| Case | Excel |
|---|---|
| A follower before its master in the sheet | refuses the file |
| A follower outside its master's `ref` | refuses the file |
| A master with no `ref` | refuses the file |
| A shared formula with a sheet reference (`Alpha!A1`, also `Alpha!$A$1`), a 3D reference or a structured reference | refuses the file |
| The master is not the top-left cell of `ref` | opens it. A cell left of the master gets the master's text moved left, and the reference wraps (`XFD2*2`) |
| A follower whose reference runs off the sheet | opens it, and wraps the reference (row 1048577 becomes row 1) |
| `ref` holds cells that are not in the group, or is larger than the group | opens it. Excel writes this itself when a cell inside a filled block is overwritten |
| `si` not from 0, with gaps, out of order, or 4,000,000 | opens it |
| Two groups whose `ref` ranges overlap | opens it |
| A follower with text, even text that disagrees with the master | opens it, and reads the master's formula |
| A defined name, a mixed range such as `SUM($A$1:A1)`, `ROW()` or `#REF!` in a shared formula | opens it |

What Excel writes as shared, measured from files that Excel saved:

- A block that is filled (`Range.Formula` on a range, or `FillDown`) or pasted. For `FillDown` and a paste, the source cell stays a formula of its own, and the shared formula starts at the next cell.
- At most 64 rows in each shared formula. A column of 50,000 filled-down formulas is 782 shared formulas, not one.
- `si` numbered from 0 on each sheet, in the order the masters appear in the sheet.
- Never a formula with a sheet, 3D or external reference, a structured reference, a defined name, a range whose two corners are anchored differently (`$A$1:A1`, `A$1:A1`, `$A1:A1`), or `ROW()`. Never an array formula, or a dynamic array formula that spills (`cm`). A dynamic array function with a single result (`SUM(SEQUENCE(A1))`) and a spill reference (`SUM(Z1#)`) are shared.
- A cell whose formula wraps or becomes `#REF!` is a formula of its own.

From those, the writer (`SharedFormulaPlan`) writes a group as follows:

- The master is the first cell of the group in row-major order, and is written first. A cell left of the master's column is written on its own, so the master is also the top-left cell of `ref`, as Excel 16 writes it. Files with such cells exist: `functions-excel-2010.xlsx` in the test resources has a master at D17 with `ref="C17:N19"`, and C18:C19 in the group.
- `ref` is the smallest area that holds the master and every cell written as a follower. Each follower is checked to be inside it.
- `si` is numbered from 0 on each sheet, in the order the masters are written.
- A cell whose text has `#REF!` is written on its own. Excel would wrap a reference there instead.
- A group whose text has `!` or `[` is written on its own in every cell, because Excel refuses a sheet reference or a structured reference in a shared formula. The test is on characters, so a string literal with either character also turns sharing off. That costs bytes, not correctness.
- A group that has only one cell left is written on its own.

## Measurements

These were measured on 2026-10-13, on a Ryzen 9 5950X, with .NET 10.0.12 and Excel 16 (Microsoft 365). Other agents ran builds and Excel on the same machine at the same time, so timings carry more noise than usual. Sizes and allocations are exact. XLibur timings come from BenchmarkDotNet (`SharedFormulaWriteBenchmarks`). Exact allocations come from `profile sharedwrite` and from probes that counted only the measuring thread.

Two inputs were saved by Excel. The first, `filldown-50k`, has numbers in A1:E50000 and `=SUM(A1:E1)` filled down column F. The second, `orders-10k`, is an order list of 10,000 rows: two `VLOOKUP` columns into another sheet, four arithmetic or `IF` columns, a running `SUM($H$2:H2)`, and a summary sheet. The repository's `TemplateWithTableSourcePivotTables.xlsx` was measured for size and allocation only. Excel did not finish opening it under automation (`RPC_E_CALL_REJECTED`), with or without this change. Each input was loaded by XLibur and saved twice, once with shared formulas and once without.

| File | Form | File bytes | Sheet XML bytes | Shared formulas / followers |
|---|---|---:|---:|---:|
| filldown-50k | Excel's input | 2,324,206 | 12,986,446 | 782 / 49,218 |
| | XLibur, shared | 3,358,194 | 15,737,252 | 782 / 49,218 |
| | XLibur, expanded | 3,715,924 | 15,789,936 | 0 |
| orders-10k | Excel's input | 1,126,109 | 6,112,342 | 628 / 39,372 |
| | XLibur, shared | 1,693,486 | 7,075,995 | 628 / 39,372 |
| | XLibur, expanded | 1,895,559 | 7,415,917 | 0 |
| pivot template | Excel's input | 235,954 | 630,667 | 133 / 1,263 |
| | XLibur, shared | 199,256 | 263,303 | 133 / 1,263 |
| | XLibur, expanded | 200,976 | 308,386 | 0 |

The compressed file is 9.6% smaller for filldown-50k and 10.7% smaller for orders-10k. The XML is barely smaller, because a follower `<x:f t="shared" si="12" />` is about as long as a short formula such as `<x:f>SUM(A12:E12)</x:f>`. The formula texts, each with different numbers, are what compress poorly.

Excel `Workbooks.Open`, with 10 opens of each file after one warm-up open, given as the median and the range:

| File | Excel's input | XLibur, shared | XLibur, expanded |
|---|---:|---:|---:|
| filldown-50k | 536 ms (509–590) | 517 ms (503–690) | 714 ms (685–777) |
| orders-10k | 626 ms (534–726) | 582 ms (537–646) | 810 ms (750–1,298) |

Exact XLibur allocations on the Excel files, as the second of two passes:

| File | Form | Save | Load of the output | First recalculation after that load | Retained by it |
|---|---|---:|---:|---:|---:|
| filldown-50k | shared | 79,100,232 | 56,738,792 | 39,295,568 | −815,304 |
| | expanded | 78,961,664 | 55,529,208 | 52,339,352 | 12,436,184 |
| orders-10k | shared | 72,790,600 | 34,705,976 | 62,519,840 | 28,202,752 |
| | expanded | 72,672,232 | 33,397,952 | 81,782,360 | 48,812,648 |
| pivot template | shared | 24,394,816 | 14,633,640 | 1,589,984 | 701,976 |
| | expanded | 24,369,200 | 14,144,760 | 2,844,352 | 1,907,784 |

The save allocates 26–139 KB more, for the plan of the sheet's shared formulas. A load of a shared formula allocates 2–4% more than a load of the same formulas written out. The loader converts the master to R1C1 and writes each later cell's A1 text from a template, where a written-out formula is kept as the text read. The first recalculation allocates 24–25% less, and keeps 12–20 MB less, because every cell of a group evaluates the group's one tree (#686).

BenchmarkDotNet, on fixtures that `SharedFormulaWriteBenchmarks` builds in the shape Excel saves (64-row shared formulas, cached values):

| Benchmark | Fixture | Shared | Mean | Error | Allocated |
|---|---|---|---:|---:|---:|
| RoundTrip (load, save) | FillDown50k | no | 513.85 ms | 10.21 ms | 129.38 MB |
| | | yes | 514.63 ms | 10.26 ms | 129.24 MB |
| | Orders10k | no | 375.91 ms | 7.29 ms | 103.50 MB |
| | | yes | 376.65 ms | 6.55 ms | 103.33 MB |
| LoadOutput | FillDown50k | no | 138.86 ms | 2.76 ms | 30.17 MB |
| | | yes | 159.77 ms | 3.46 ms | 31.34 MB |
| | Orders10k | no | 77.31 ms | 1.39 ms | 22.60 MB |
| | | yes | 81.73 ms | 1.60 ms | 23.85 MB |

| LoadOutputAndRecalculate | FillDown50k | no | 352.5 ms | 7.00 ms | 91.60 MB |
| | | yes | 269.4 ms | 5.38 ms | 80.33 MB |
| | Orders10k | no | 3,879.8 ms | 76.70 ms | 107.35 MB |
| | | yes | 3,868.5 ms | 73.54 ms | 90.23 MB |

The round trip takes the same time either way. A load of the file with shared formulas takes 6–15% longer. A load followed by a recalculation takes 24% less time on FillDown50k, and the same time on Orders10k, where the running total and the lookups use most of the time. It allocates 12–16% less on both.

A file that XLibur itself wrote with shared formulas is now saved back unchanged, byte for byte, when its cells are unchanged. The writer writes what the file had, so the save keeps the sheet as it was (#702). On FillDown50k, that saves 29 MB of allocation on a save.

Readers of the XLibur output:

| Reader | Result |
|---|---|
| Excel 16 | Opens every output, and every edge case from the tests: an edited follower, an edited or cleared master, a row inserted or deleted inside a group, a copy, a copied sheet, `#REF!` in the first cell, a cell left of the master, and several groups numbered in order. The formulas it shows match what XLibur reads. |
| openpyxl 3.1.5 | Reads the same formula in every cell of each output as in Excel's input (50,000, 70,004 and 1,450 formula cells), and as in the written-out form of each edge case. |
| Open XML SDK 3.4.1 `OpenXmlValidator` | No errors in the outputs of the three inputs. |
| LibreOffice | Not installed on the machine, so not tested. |

## Consequences

A load and a save of a file from Excel keep its shared formulas, and only the cells that a caller changed lose them. XLibur's own next load of the file reads every shared formula through one tree.

Formulas that a caller writes are never shared. A file that XLibur creates from nothing still has a formula of its own in every cell.

A rename of a sheet that a shared formula refers to drops the group, but such a group is never written as shared anyway. A row inserted above a group moves every cell of it and writes each on its own, although the moved block would still be a valid shared formula. Keeping a group through a uniform move would need the group's anchor to move with it, and is left for later.

If #720 keeps an unchanged formula sheet as the file had it, that sheet's shared formulas are kept with it. This decision covers the sheets that the save writes again.
