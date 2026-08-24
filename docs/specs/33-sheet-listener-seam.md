# Spec 33 — Every sheet feature reacts to a structural edit through one seam

**Area:** Architecture · Refactor · **Defect (4 features do not move)**
**Effort:** M–L (~6–8 days)
**Dependencies:** **Hard — spec 26 (grid axis) must land first.** Shares
`XLWorksheetRangeShifter.cs` and `XLWorksheet.cs`. See Conflicts.
**Status:** Proposed.

## Goal

Make registering an `ISheetListener` the one thing that makes a sheet feature edit-aware, so that
adding a feature that must survive an insert or a delete means writing one adapter instead of
editing three files — and so that the four features that react to nothing today start reacting.

## Why this spec exists

`ISheetListener` (`XLibur/Excel/Cells/ISheetListener.cs`) is a declared seam with four methods:

```csharp
void OnInsertAreaAndShiftDown(XLWorksheet sheet, Area area);      // :15
void OnInsertAreaAndShiftRight(XLWorksheet sheet, Area area);     // :22
void OnDeleteAreaAndShiftLeft(XLWorksheet sheet, Area deletedRange);  // :29
void OnDeleteAreaAndShiftUp(XLWorksheet sheet, Area deletedRange);    // :36
```

**Two types implement it.** `grep -rn 'ISheetListener' XLibur --include=*.cs` returns 12 lines, of
which the implementations are `XLCalcEngine.cs:19` and `XLHyperlinks.cs:10`. Neither is found by
enumeration — both are reached by name, from four hardcoded lines in the shifter.

Seventeen sheet-scoped features must survive a row or column insert. They do it four different ways:

| # | Feature | How it reacts | Moves? |
|---|---|---|---|
| 1 | Cell formulas, dependency tree | `ISheetListener` — `XLCalcEngine.cs:19` | yes |
| 2 | Hyperlinks | `ISheetListener` — `XLHyperlinks.cs:10` | yes |
| 3 | Merged ranges (straddle split) | hardcoded — `XLWorksheetRangeShifter.cs:24-33` / `:67-76` | yes |
| 4 | Conditional formats | hardcoded — `:127` / `:141` | yes |
| 5 | Data validations (sqref) | hardcoded — `:176` / `:192` | yes |
| 6 | Data-validation criteria formulas | hardcoded, **workbook-wide** — `:239` / `:259` | yes |
| 7 | Defined names | hardcoded, **workbook-wide** — `:35-36` / `:78-79` | yes |
| 8 | Page breaks | hardcoded — `:103` / `:115` | yes |
| 9 | Sparklines | hardcoded cleanup at `:273`, **plus** a shift called from `XLRangeInsertHelper.cs:21` / `:128` | yes |
| 10 | Tables | range repository (`XLTable : XLRange`, `XLTable.cs:16`) | yes |
| 11 | AutoFilter | range repository (`XLAutoFilter.cs:30` holds an `IXLRange`) | yes |
| 12 | Selected ranges | range repository | yes |
| 13 | Pictures | range repository **via a fake 1-cell range** — `XLMarker.cs:10-12` | yes |
| 14 | Chart anchors | raw `int` — `XLDrawingPosition.cs:5,10` | **no** |
| 15 | Note (comment) anchors | raw `int` — same type, `XLComment.cs:103` | **no** |
| 16 | Freeze / split panes | raw `int` — `XLSheetView.cs:34,36,38` | **no** |
| 17 | Pivot table `Area` | raw `Area` — `XLPivotTable.cs:845` | **no** |

The brief for this spec estimated "roughly 16". Seventeen is the count that falls out of the tree at
`1b41cadd`; the difference is that sparklines are dispatched from two places, not one.

### The four that do not move — measured, not argued

A probe was written against the unmodified tree, run, and deleted. Five features, one workbook each,
`ws.Row(1).InsertRowsAbove(3)` then `ws.Column(1).InsertColumnsBefore(2)`:

| Feature | Anchored at | After the edit | Correct if it shifted |
|---|---|---|---|
| Chart `Position` | row 10, col 3 | **row 10, col 3** | row 13, col 5 |
| Chart `SecondPosition` | row 20, col 10 | **row 20, col 10** | row 23, col 12 |
| Note `Position` on `C10` | row 9, col 4 | **row 9**, col 4 | row 12, col 4 |
| `SheetView.SplitRow` / `SplitColumn` | 5 / 4 | **5 / 4** | 8 / 6 |
| `XLPivotTable.Area` | `D10` | **`D10`** | `D13` |
| Picture `TopLeftCell` (control) | `C10` | `C13` ✅ | `C13` |

Four of the five fail. The one that passes is the picture — the feature that gets itself shifted by
allocating a range it does not otherwise want (row 13 of the table above). That contrast is the
whole spec in one run.

The note case is the sharpest, because it is *half* right. The note itself moves: after the insert,
`ws.Cell("C13").HasComment` is `true` and `ws.Cell("C10").HasComment` is `false`, because notes live
in the misc slice and the slice shifts with the cells. Its **drawing anchor** does not move, so the
callout box stays pinned three rows above where the note now is. `XLComment.cs:8-13` already
documents the mechanism that causes this:

> *"Only a hint for `Delete`: shifting rows or columns moves the note's entry within the misc slice
> **without telling the note**, so the address can name a cell the note has since moved off."*

The library knows the note is not told. It works around the consequence for `Delete` and leaves the
anchor wrong.

### The documented workaround

`XLibur/Excel/Drawings/XLMarker.cs:10-12`, verbatim:

```csharp
    // Using a range to store the location so that it gets added to the range repository
    // and hence will be adjusted when there are insertions / deletions
    private readonly IXLRange rangeCell;
```

A picture anchor is a cell and a pixel offset. It stores an `IXLRange` — validated at `:26-27` to be
exactly one cell, and immediately unwrapped again by `Cell`, `RowNumber` and `ColumnNumber`
(`:33`, `:35`, `:39`) — for no reason except that a range gets shifted and a `Point` does not. This
is the seam being smuggled past, in a comment, in shipped code. Every listed feature that works
today works either through the port or through this trick.

### The rest of the evidence

`XLRowBatchDelete.cs:19-22` enumerates what shifts, and the omissions are the point:

> *"the live ranges, conditional formats, data validations, merges, page breaks and hyperlinks all
> still shift a block at a time"*

Six features named. No charts, no notes, no panes, no pivot tables.

Structural-edit knowledge is spread over **20 files and 76 methods**. The brief estimated 13 and 62;
that is not what the tree shows. Reproduce with:

```
grep -rn --include=*.cs -E '(private|internal|public|protected)[^(]*\b(Shift[A-Za-z]*|[A-Za-z]*AndShift[A-Za-z]*|OnInsertArea[A-Za-z]*|OnDeleteArea[A-Za-z]*|RelocateRange|Move[A-Za-z]*Rows|Move[A-Za-z]*Columns|InsertRows|InsertColumns|DeleteRows|DeleteColumns)\s*(<[^>]*>)?\s*\(' XLibur
```

77 hits, one of which (`Engineering.cs:154`, a bit-shift helper) is a false positive.

### The pattern already exists one level up

`XLWorksheets.cs:229`:

```csharp
    private IEnumerable<IWorkbookListener> GetWorkbookListeners()
    {
        // All components that should be updated when sheet is added/removed or renamed should
        // be enumerated here.
        yield return _workbook.CalcEngine;

        foreach (var sheet in _worksheets.Values)
            yield return sheet.Internals.CellsCollection;

        foreach (var definedName in _workbook.DefinedNamesInternal)
            yield return definedName;

        foreach (var sheet in _worksheets.Values)
            foreach (var definedName in sheet.DefinedNames)
                yield return definedName;
    }
```

One method, one comment telling the next author where to add themselves, consumed by a `foreach` at
`:223-224`. Three types implement `IWorkbookListener` and all three are reached this way. The
workbook has the registry. The worksheet does not.

Note what that method also settles: it yields listeners belonging to **other sheets**. Sheet scope is
a property of what the enumeration yields, not a limit on it. That is what makes features 6 and 7 —
which visit every worksheet in the workbook — expressible without widening the port.

## Non-goals

- **Not touching the range repository or the spatial index.** Spec 05 declined that on measurement:
  filtering unreachable ranges made them free, so the enumeration an index would remove was never
  the expense, and an index would add a second source of truth over a weak-reference store that is
  mutated during the iteration it would serve. Read spec 05's Results §"A2: the index was declined"
  before proposing otherwise.
- **Not a performance spec.** Spec 05 measured formula shifting at 68% of the insert workload and
  the range-shift pass at 8%. This spec reorganises the 8% and must not make the 68% worse.
  Task 7 is the gate.
- **Not touching `XLCellFormulaShifter*.cs`.** Spec 25 owns those.
- **Not re-litigating spec 26's ordering decision** (see task 1).
- **No public API change.** `PublicAPI.Unshipped.txt` untouched.

## Current state

Verified against the tree at `1b41cadd` (2026-08-24). Every line number below was read, not carried
over. **Two corrections to the brief this spec was written from:** the listener dispatch lines are
`:43,49,55` for **columns** and `:86,92,98` for **rows** (the brief had them swapped), and the
feature count is 17, not 16.

- `ISheetListener` — `XLibur/Excel/Cells/ISheetListener.cs:8`, four methods at `:15`, `:22`, `:29`, `:36`
- `XLCalcEngine : ISheetListener, IWorkbookListener` — `XLibur/Excel/CalcEngine/XLCalcEngine.cs:19`
- `XLHyperlinks : IXLHyperlinks, ISheetListener` — `XLibur/Excel/Hyperlinks/XLHyperlinks.cs:10`;
  the `sheet != _worksheet` guard idiom at `:64-66`
- `XLWorksheetRangeShifter.ShiftColumns` — `XLibur/Excel/XLWorksheetRangeShifter.cs:17-58`
- `XLWorksheetRangeShifter.ShiftRows` — `:60-101`
- Listener dispatch — `:43`, `:49`, `:55` (columns) and `:86`, `:92`, `:98` (rows)
- `XLWorksheet._rangeShifter` — `XLibur/Excel/XLWorksheet.cs:31`, constructed `:81`, called `:1240`, `:1245`
- `NotifyRangeShiftedRows` / `NotifyRangeShiftedColumns` — `XLWorksheet.cs:1248`, `:1280`
- Insert entry — `XLibur/Excel/Ranges/XLRangeInsertHelper.cs:36`, `:142`
- Delete entry — `XLibur/Excel/Ranges/XLRangeBase.cs:1156-1158`; `rowModifier = RowCount()` at `:1147`
- `XLWorksheets.GetWorkbookListeners` — `XLibur/Excel/XLWorksheets.cs:229-250`, consumed `:223`
- `XLMarker` — `XLibur/Excel/Drawings/XLMarker.cs`, comment `:10-12`, one-cell check `:26-27`
- `XLDrawingPosition` — `XLibur/Excel/Drawings/XLDrawingPosition.cs:5`, `:10` (raw `int`)
- `XLDrawing<T>.Position` — `XLibur/Excel/Drawings/XLDrawing.cs:9`, `:40`; `XLChart.SecondPosition`
  — `XLibur/Excel/Charts/XLChart.cs:44`, `:85`
- `XLComment.Position` — `XLibur/Excel/Comments/XLComment.cs:103`, seeded `:193-199`
- `XLSheetView` — `XLibur/Excel/XLSheetView.cs:34` (`FreezePanes`), `:36` (`SplitColumn`), `:38` (`SplitRow`)
- `XLPivotTable.Area` — `XLibur/Excel/PivotTables/XLPivotTable.cs:845`; `TargetCell` derives from it, `:59-70`
- `Area.ExtendBelow` / `ExtendRight` — `XLibur/Excel/Coordinates/Area.cs:238`, `:248`
- `Area.TryInsertAreaAndShiftDown` etc. — `Area.cs:443`, `:497`, `:558`, `:624`
- Structural-edit profiler — `XLibur.Benchmarks/StructuralEditProfile.cs`,
  `dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile structural`

### The two orderings differ

`ShiftColumns` and `ShiftRows` do the same seven things in **different orders**:

| Step | `ShiftColumns` (`:17-58`) | `ShiftRows` (`:60-101`) |
|---|---|---|
| 1 | merged-range straddle split | merged-range straddle split |
| 2 | defined names (all sheets, then workbook) | defined names (all sheets, then workbook) |
| 3 | conditional formats | conditional formats |
| 4 | data validations (sqref) | data validations (sqref) |
| 5 | data-validation criteria formulas | data-validation criteria formulas |
| 6 | **page breaks** | **sparkline cleanup** |
| 7 | **sparkline cleanup** | **page breaks** |
| 8 | `CalcEngine` | `CalcEngine` |
| 9 | `Hyperlinks` | `Hyperlinks` |

Steps 6 and 7 are swapped. **Spec 26 task 8 owns reconciling that**, and its analysis is that the two
commute — `RemoveInvalidSparklines` (`:273-283`) reads only sparkline address validity, and
`ShiftPageBreaks*` (`:103-125`) touches only `PageSetup.*Breaks`, so they share no state. 26 forces a
single order and pins it with a test. **This spec inherits that decision and does not re-open it** —
but task 1 still pins the whole order, because enumerating listeners must not change any of it by
accident.

## File structure

```
XLibur/Excel/Cells/ISheetListener.cs                    modified — SheetEdit argument, ordering contract
XLibur/Excel/Cells/SheetEdit.cs                         new — the readonly struct the four methods take
XLibur/Excel/XLWorksheet.cs                             modified — GetSheetListeners()
XLibur/Excel/XLWorksheetRangeShifter.cs                 modified — ~300 lines out, enumeration in
XLibur/Excel/Ranges/MergedRangeSplitListener.cs         new — adapter (owner is a general-purpose XLRanges)
XLibur/Excel/Drawings/DrawingAnchorListener.cs          new — adapter for charts and notes
XLibur/Excel/ConditionalFormats/XLConditionalFormats.cs modified — implements ISheetListener
XLibur/Excel/DataValidations/XLDataValidations.cs       modified — implements ISheetListener
XLibur/Excel/DefinedNames/XLDefinedNames.cs             modified — implements ISheetListener
XLibur/Excel/PageSetup/XLPageSetup.cs                   modified — implements ISheetListener
XLibur/Excel/Sparkline/XLSparkLineGroups.cs             modified — implements ISheetListener
XLibur/Excel/XLSheetView.cs                             modified — implements ISheetListener
XLibur/Excel/PivotTables/XLPivotTable.cs                modified — implements ISheetListener
XLibur/Excel/Drawings/XLMarker.cs                       modified — holds a Point, not an IXLRange
XLibur/Excel/CalcEngine/XLCalcEngine.cs                 modified — signature only
XLibur/Excel/Hyperlinks/XLHyperlinks.cs                 modified — signature only
XLibur.Tests/Excel/Ranges/SheetListenerCharacterizationTests.cs  new — task 1
XLibur.Tests/Excel/Ranges/SheetListenerOrderTests.cs             new — task 2
```

The exact filenames of the collection types must be confirmed with
`grep -rln 'class XLConditionalFormats\|class XLDataValidations\|class XLDefinedNames\|class XLPageSetup' XLibur`
before starting; the list above is the shape, not a promise about paths.

## The design

### 1. The registry, modelled on the one that already exists

```csharp
    /// <summary>
    /// Every component that must react when an area is inserted into or deleted from this sheet.
    /// The workbook-level counterpart is <see cref="XLWorksheets.GetWorkbookListeners"/>.
    /// </summary>
    /// <remarks>
    /// All components that should be updated when rows or columns are inserted or deleted should be
    /// enumerated here — and nowhere else. Adding a sheet feature that must survive a structural
    /// edit is one adapter plus one <c>yield return</c>.
    /// <para>
    /// <b>Order is part of the contract.</b> It is pinned by
    /// <c>SheetListenerOrderTests</c>; changing it is a behaviour change and needs that test
    /// updated deliberately. Listeners belonging to other sheets are yielded too — defined names
    /// and data-validation criteria formulas are workbook-scoped and must see an edit on any sheet.
    /// Such a listener guards on the sheet it is given, the way
    /// <see cref="XLHyperlinks"/> does.
    /// </para>
    /// </remarks>
    internal IEnumerable<ISheetListener> GetSheetListeners()
    {
        yield return _mergedRangeSplitter;

        foreach (var sheet in Workbook.WorksheetsInternal)
            yield return sheet.DefinedNames;
        yield return Workbook.DefinedNamesInternal;

        yield return ConditionalFormats;
        yield return DataValidations;

        foreach (var sheet in Workbook.WorksheetsInternal)
            yield return sheet.DataValidations;   // criteria formulas, workbook-wide

        yield return (XLPageSetup)PageSetup;
        yield return SparklineGroupsInternal;

        yield return Workbook.CalcEngine;
        yield return Hyperlinks;

        // Features that did not react before spec 33.
        yield return _drawingAnchors;
        yield return SheetView;
        foreach (var pivotTable in PivotTables.Cast<XLPivotTable>())
            yield return pivotTable;
    }
```

The shifter's two methods collapse to:

```csharp
    public void ShiftRows(XLRange range, int rowsShifted)
        => Dispatch(range, rowsShifted, Axis.Row);

    public void ShiftColumns(XLRange range, int columnsShifted)
        => Dispatch(range, columnsShifted, Axis.Column);
```

`Axis` is spec 26's type. If 26 has not landed, this spec has not started.

### 2. The port keeps four methods and gains one argument

The four callbacks take `(XLWorksheet sheet, Area area)`. That is **not enough information** for
three of the features being converted, and the arithmetic says so:

- `area` for an insert is built at `XLWorksheetRangeShifter.cs:89-91` as
  `Area.FromRangeAddress(range.RangeAddress).ExtendBelow(rowsShifted - 1)`, and `ExtendBelow`
  (`Area.cs:238-243`) extends from `LastPoint.Row`. So `area.Height == range.Height + shift - 1`.
  For `ws.Row(1).InsertRowsAbove(3)` that is `1 + 3 - 1 = 3` and the shift is recoverable. For
  `ws.Range("A1:A5").InsertRowsAbove(3)` it is `5 + 3 - 1 = 7`, and **the shift is not recoverable
  from the area alone**.
- Page breaks need the signed shift (`:110`, `:122`), not an area.
- Defined names and data-validation criteria formulas need the `XLRange` itself, because they hand
  it to `XLCellFormulaShifter.ShiftFormula*` (`:246`, `:266`, `:294`, `:312`).

So the argument widens, but the interface does not gain a method:

```csharp
/// <summary>
/// One structural edit, as the shifter sees it. Passed to every <see cref="ISheetListener"/>.
/// </summary>
/// <remarks>
/// <see cref="Area"/> alone is not sufficient: for an insert it is the edited range extended by
/// <c>Shift - 1</c> rows or columns (<c>XLWorksheetRangeShifter.ShiftRows</c>), so the shift
/// magnitude cannot be recovered from it when the edited range is taller than one row. Listeners
/// that only need the area — the calc engine and hyperlinks — read <see cref="Area"/> and ignore
/// the rest.
/// </remarks>
internal readonly struct SheetEdit
{
    /// <summary>The sheet the edit happened on. Not necessarily the listener's own sheet.</summary>
    internal required XLWorksheet Sheet { get; init; }

    /// <summary>The area inserted, or the area deleted.</summary>
    internal required Area Area { get; init; }

    /// <summary>The edited range, as the caller passed it. Not extended by <see cref="Shift"/>.</summary>
    internal required XLRange Range { get; init; }

    /// <summary>Rows or columns shifted. Positive for an insert, negative for a delete.</summary>
    internal required int Shift { get; init; }
}
```

```csharp
void OnInsertAreaAndShiftDown(in SheetEdit edit);
void OnInsertAreaAndShiftRight(in SheetEdit edit);
void OnDeleteAreaAndShiftLeft(in SheetEdit edit);
void OnDeleteAreaAndShiftUp(in SheetEdit edit);
```

Four methods before, four after. `in` so a 20-odd-byte struct is not copied twelve times per edit.

**This widening is a premise, and task 3 step 1 is what would disprove it.** If the arithmetic above
turns out to be wrong — if `ExtendBelow(shift - 1)` on a multi-row range is itself a latent defect
and the area is always exactly `shift` rows tall — then `Shift` is recoverable, `Range` is
reconstructible, and the port should stay at `(sheet, area)`. Establish which before widening
anything. **A disproved premise here is a better result than a wider port.**

### 3. What each converted feature becomes

Where a feature already has a collection type, that type implements the port — the precedent is
`XLHyperlinks`, which is both `IXLHyperlinks` and `ISheetListener`. Where the owner is a
general-purpose type (merged ranges are an `XLRanges`), a dedicated adapter carries it. Charts and
notes share one adapter, because they share one anchor type (`XLDrawingPosition`) and want the same
answer.

### 4. `XLMarker` goes back to being what it is

```csharp
[DebuggerDisplay("R{RowNumber}C{ColumnNumber} {Offset}")]
internal sealed class XLMarker
{
    // Anchors are shifted by DrawingAnchorListener, registered in XLWorksheet.GetSheetListeners.
    // Before spec 33 this field was an IXLRange, allocated purely so the range repository would
    // move it — the seam smuggled past, in a comment.
    private Point cell;   // XLibur.Excel.Coordinates.Point: (Row, Column)
    ...
}
```

Note the type collision: `XLMarker.Offset` is a `System.Drawing.Point` (pixels) and the anchor cell
is an `XLibur.Excel.Coordinates.Point` (row/column). Alias one of them; do not let the file import
both unqualified.

## Global constraints

- Warnings are errors (`TreatWarningsAsErrors=true`); nullable enabled — new code must be
  null-annotated.
- Branch per spec, `refactor/33-sheet-listener-seam`; never commit to main. Commit prefixes
  `refactor:` / `fix:` / `test:` / `perf:`.
- No compound shell commands (`&&`, `||`, `;`) in agent tool calls.
- **Do not use `sed -i` on tracked files.** `.gitattributes` checks out CRLF and Git Bash's `sed -i`
  rewrites the file as LF, turning a one-line change into a whole-file diff. Use the Edit/Write
  tools; verify with `git diff --numstat` — a file whose changed-line count is near its total line
  count was rewritten, not edited.
- Test filtering uses `--treenode-filter`, never `--filter`. Exit 5 = invalid option; exit 8 = zero
  tests matched. Never filter at solution level — name the `.csproj`.
- Pass `-f net10.0` for iteration; run without it before opening the PR.
- Build: `dotnet build XLibur/XLibur.csproj -c Release -v q`
- Tests: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
- Tests are TUnit and **assertions are awaitable**: `await Assert.That(actual).IsEqualTo(expected)`.
  A missing `await` silently passes. `[Test]`, `[Arguments(...)]`, `[MethodDataSource(...)]`. The
  suite is serial (`[assembly: NotInParallel]`).
- `required` members need C# 11+; the repo targets net8.0/net9.0/net10.0, so this is available.

## Work plan

| # | Task | Size | Gate |
|---|------|------|------|
| 1 | Characterization tests for all 17 features, including the 4 that do not move | M | New tests green on unmodified code; four of them assert the wrong answer, on purpose |
| 2 | `SheetEdit`, `GetSheetListeners()`, the two existing adapters routed through it; order pinned | M | Suite green; `SheetListenerOrderTests` green |
| 3 | Convert the six hardcoded features | M | Suite green; shifter names no feature |
| 4 | Chart and note anchors become an adapter — **behaviour change** | M | Task 1's chart/note tests re-pointed and green |
| 5 | Panes and pivot `Area` become adapters — **behaviour change** | M | Task 1's pane/pivot tests re-pointed and green |
| 6 | Delete the `XLMarker` workaround | S | Suite green; no `IXLRange` in `XLMarker.cs` |
| 7 | Confirm structural-edit cost unchanged | S | Within spec 05's profile noise |

Tasks 4 and 5 change output for existing documents. Each carries its own Excel-behaviour decision,
its own changelog entry, and the commit message that says an assertion was deliberately reversed.

---

### Task 1 — Pin today's behaviour for all seventeen features

Everything after this task is gated by it, including the four features whose current behaviour is
wrong. **Those four tests assert the wrong answer deliberately**, so that tasks 4 and 5 have to
change them and cannot change them silently.

**Files:**
- Create: `XLibur.Tests/Excel/Ranges/SheetListenerCharacterizationTests.cs`

**Interfaces:**
- Produces: `A_row_insert_moves_every_sheet_feature` and its column twin, plus four
  `*_does_not_move_yet` tests that tasks 4 and 5 re-point.

- [ ] **Step 1: One workbook carrying all seventeen features, edited once**

```csharp
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Ranges;

/// <summary>
/// Every sheet-scoped feature that must survive a structural edit, in one workbook, edited once.
/// Spec 33 moves the dispatch for all of them behind <c>ISheetListener</c>; this is what proves no
/// feature was dropped on the way.
/// <para>
/// The four <c>*_does_not_move_yet</c> tests assert the <b>current, wrong</b> behaviour. They exist
/// so that spec 33 tasks 4 and 5 must change them explicitly. Do not "fix" them here.
/// </para>
/// </summary>
public class SheetListenerCharacterizationTests
{
    [Test]
    public async Task A_row_insert_moves_every_sheet_feature()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("S");

        ws.Range("B10:D10").Merge();                                  // merged ranges
        ws.Range("B12:D14").AddConditionalFormat().WhenGreaterThan(5)
            .Fill.SetBackgroundColor(XLColor.Red);                    // conditional formats
        ws.Range("B16:B17").CreateDataValidation().WholeNumber.Between(1, 10);  // DV sqref
        ws.Range("B18:B19").CreateDataValidation().List(ws.Range("F20:F22"));   // DV criteria formula
        wb.DefinedNames.Add("Block", ws.Range("B22:C24"));            // defined names
        ws.PageSetup.AddHorizontalPageBreak(30);                      // page breaks
        ws.Cell("B26").SetValue("x").SetHyperlink(new XLHyperlink("https://example.invalid/"));
        ws.Cell("B28").FormulaA1 = "=B10+1";                          // calc engine
        ws.Range("B30:C30").AsTable();                                // tables
        ws.Range("B32:D33").SetAutoFilter();                          // autofilter

        ws.Row(1).InsertRowsAbove(3);

        await Assert.That(ws.MergedRanges.Single().RangeAddress.ToString()).IsEqualTo("B13:D13");
        await Assert.That(ws.ConditionalFormats.Single().Ranges.Single()
            .RangeAddress.ToString()).IsEqualTo("B15:D17");
        await Assert.That(ws.DataValidations.First().Ranges.Single()
            .RangeAddress.ToString()).IsEqualTo("B19:B20");
        await Assert.That(wb.DefinedNames.First().RefersTo).Contains("B25:C27");
        await Assert.That(ws.PageSetup.RowBreaks.Single()).IsEqualTo(33);
        await Assert.That(ws.Cell("B29").HasHyperlink).IsTrue();
        await Assert.That(ws.Cell("B31").FormulaA1).IsEqualTo("B13+1");
        await Assert.That(ws.Tables.Single().RangeAddress.ToString()).IsEqualTo("B33:C33");
        await Assert.That(ws.AutoFilter.Range.RangeAddress.ToString()).IsEqualTo("B35:D36");
    }
```

Write the column twin the same way with `ws.Column(1).InsertColumnsBefore(2)`, and a delete twin for
each axis. Four tests, one per operation the port has a method for.

**The expected strings above are predictions, not readings.** Run the test, and where it disagrees,
**record the actual value and use it** — this task pins what the library does, not what it should do.
Any value that looks wrong is a finding to report, not a value to correct here.

- [ ] **Step 2: The four that do not move, asserted as they are**

```csharp
    /// <summary>
    /// A chart anchor is a pair of raw <c>int</c>s on <see cref="XLDrawingPosition"/> and nothing
    /// notifies it, so a chart anchored below an insert stays where it was. Spec 33 task 4 fixes
    /// this and re-points this test.
    /// </summary>
    [Test]
    public async Task A_chart_anchor_does_not_move_yet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = "Q1";
        ws.Cell("B1").Value = 100;
        var chart = ws.Charts.Add(XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$1", "Data!$A$1:$A$1");
        chart.Position.SetColumn(3).SetRow(10);
        chart.SecondPosition.SetColumn(10).SetRow(20);

        ws.Row(1).InsertRowsAbove(3);

        // WRONG on purpose. Correct is 13 / 23 once spec 33 task 4 lands.
        await Assert.That(chart.Position.Row).IsEqualTo(10);
        await Assert.That(chart.SecondPosition.Row).IsEqualTo(20);
    }
```

The other three, measured on the unmodified tree at `1b41cadd` and reproduced by the probe described
in "Why this spec exists":

| Test | Asserts now | Asserts after |
|---|---|---|
| `A_chart_anchor_does_not_move_yet` | `Position.Row == 10`, `SecondPosition.Row == 20` | 13, 23 (task 4) |
| `A_note_anchor_does_not_move_yet` | note is on `C13`; `Position.Row == 9` | note on `C13`; `Position.Row == 12` (task 4) |
| `Freeze_panes_do_not_move_yet` | `SplitRow == 5`, `SplitColumn == 4` | 8, 6 (task 5) |
| `A_pivot_area_does_not_move_yet` | `pt.Area.FirstPoint.Row == 10` | 13 (task 5) |

The note test must assert **both halves**: the note moved (`ws.Cell("C13").HasComment` is `true`)
and its anchor did not (`Position.Row == 9`). That split is the defect, and a test that only checks
one half will not notice when the other is fixed.

- [ ] **Step 3: Verify the gate bites**

Comment out `ShiftPageBreaksRows` at `XLWorksheetRangeShifter.cs:84` and re-run.
Expected: FAIL on the `RowBreaks` assertion. Restore the line.

- [ ] **Step 4: Run**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/SheetListenerCharacterizationTests/*"`
Expected: PASS, all of them, on unmodified production code.

- [ ] **Step 5: Commit**

```bash
git add XLibur.Tests/Excel/Ranges/SheetListenerCharacterizationTests.cs
git commit -m 'test(shift): pin every sheet feature against a structural edit (spec 33 task 1)'
```

Note in the commit body that four tests assert the current **wrong** behaviour on purpose and are
re-pointed by tasks 4 and 5.

---

### Task 2 — `SheetEdit`, the registry, and the order

**Files:**
- Create: `XLibur/Excel/Cells/SheetEdit.cs`
- Create: `XLibur.Tests/Excel/Ranges/SheetListenerOrderTests.cs`
- Modify: `XLibur/Excel/Cells/ISheetListener.cs`
- Modify: `XLibur/Excel/XLWorksheet.cs`
- Modify: `XLibur/Excel/XLWorksheetRangeShifter.cs:43-57`, `:86-100`
- Modify: `XLibur/Excel/CalcEngine/XLCalcEngine.cs`, `XLibur/Excel/Hyperlinks/XLHyperlinks.cs` (signatures)

**Interfaces:**
- Produces: `SheetEdit`, `XLWorksheet.GetSheetListeners() → IEnumerable<ISheetListener>`, and the
  four `ISheetListener` methods taking `in SheetEdit`.

- [ ] **Step 1: Pin the order before anything moves**

```csharp
/// <summary>
/// The order sheet listeners run in. Spec 33 replaces nine hardcoded calls with an enumeration;
/// this is what proves the enumeration did not reorder them. Spec 26 task 4 reconciled the
/// row/column discrepancy in steps 6 and 7 — this test records the outcome, it does not decide it.
/// </summary>
[Test]
public async Task Sheet_listeners_run_in_the_pinned_order()
{
    using var wb = new XLWorkbook();
    var ws = (XLWorksheet)wb.AddWorksheet("S");
    var names = ws.GetSheetListeners().Select(l => l.GetType().Name).ToList();
    await Assert.That(names).IsEquivalentTo(new[] { /* filled from the first run */ });
}
```

Fill the array from the first run. Before that, read `ShiftRows` and `ShiftColumns` on the branch
point and write the current order into the test's summary as a comment, so a reviewer can check the
enumeration against the code it replaced without re-deriving it.

**If spec 26 task 4 left the two axes in different orders**, this test needs an axis parameter and
two expected sequences. Do not quietly unify them here — that decision belongs to 26.

- [ ] **Step 2: `SheetEdit` and the widened signatures**

Create `SheetEdit` as given in "The design". Change the four `ISheetListener` methods to take
`in SheetEdit`. `XLCalcEngine` and `XLHyperlinks` change signature only — their bodies read
`edit.Sheet` and `edit.Area` and are otherwise untouched.

`XLHyperlinks.RepositionOnChange` (`:62-81`) already guards with `if (sheet != _worksheet) return;`.
Keep that guard verbatim; it is the idiom every cross-sheet listener in task 3 will use.

- [ ] **Step 3: `GetSheetListeners()` with only the two existing adapters**

Add the method to `XLWorksheet` with the doc comment from "The design", but yielding **only**
`Workbook.CalcEngine` and `Hyperlinks` for now. Replace `XLWorksheetRangeShifter.cs:43-57` and
`:86-100` with the enumeration:

```csharp
        var edit = new SheetEdit
        {
            Sheet = range.Worksheet,
            Area = shift > 0
                ? Area.FromRangeAddress(range.RangeAddress).ExtendBelow(shift - 1)
                : Area.FromRangeAddress(range.RangeAddress),
            Range = range,
            Shift = shift,
        };

        foreach (var listener in worksheet.GetSheetListeners())
        {
            if (shift > 0)
                listener.OnInsertAreaAndShiftDown(in edit);
            else if (shift < 0)
                listener.OnDeleteAreaAndShiftUp(in edit);
        }
```

The `shift == 0` case does nothing today — `:87`/`:95` are `if`/`else if` with no `else`. Preserve
that.

- [ ] **Step 4: Build and run**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS, including all of task 1.

- [ ] **Step 5: Commit**

```bash
git add XLibur/Excel/Cells/SheetEdit.cs XLibur/Excel/Cells/ISheetListener.cs XLibur/Excel/XLWorksheet.cs XLibur/Excel/XLWorksheetRangeShifter.cs XLibur/Excel/CalcEngine/XLCalcEngine.cs XLibur/Excel/Hyperlinks/XLHyperlinks.cs XLibur.Tests/Excel/Ranges/SheetListenerOrderTests.cs
git commit -m 'refactor(shift): give the worksheet a sheet-listener registry (spec 33 task 2)'
```

---

### Task 3 — Convert the six hardcoded features

**Files:**
- Create: `XLibur/Excel/Ranges/MergedRangeSplitListener.cs`
- Modify: the five feature collections listed in "File structure"
- Modify: `XLibur/Excel/XLWorksheetRangeShifter.cs` — `:24-33`, `:35-36`, `:67-76`, `:78-79`,
  `:103-125`, `:127-155`, `:176-206`, `:239-271`, `:273-283`, `:285-319` all leave
- Modify: `XLibur/Excel/XLWorksheet.cs` — the enumeration grows

- [ ] **Step 1: Settle the `SheetEdit` premise before moving anything**

Write one test that asserts the relationship the design depends on:

```csharp
/// <summary>
/// The listener's Area is the edited range extended by Shift-1, so the shift magnitude is not
/// recoverable from the Area when the edited range is more than one row tall. This is why SheetEdit
/// carries Range and Shift. If this test fails, the premise is wrong and the port should be narrowed
/// back to (sheet, area) — record that and stop.
/// </summary>
[Test]
[Arguments("A1:A1", 3, 3)]
[Arguments("A1:A5", 3, 7)]
public async Task The_inserted_area_is_the_range_extended_by_shift_minus_one(
    string rangeAddress, int shift, int expectedAreaHeight)
```

If `A1:A5` with a shift of 3 produces an area 3 rows tall rather than 7, `ExtendBelow(shift - 1)` is
not doing what `Area.cs:238-243` says, `Shift` is recoverable, and **`SheetEdit` should be cut back
to `Sheet` and `Area`**. Record the finding and revise task 2 before continuing. That outcome is a
result, not a setback.

- [ ] **Step 2: Convert the sheet-scoped four**

Merged ranges, conditional formats, data-validation sqref areas, and page breaks are all
sheet-scoped. Each moves verbatim into an `ISheetListener` implementation on its owner, guarded on
`edit.Sheet`:

```csharp
    void ISheetListener.OnInsertAreaAndShiftDown(in SheetEdit edit)
    {
        if (edit.Sheet != _worksheet)
            return;

        for (var i = 0; i < RowBreaks.Count; i++)
        {
            var br = RowBreaks[i];
            if (edit.Range.RangeAddress.FirstAddress.RowNumber <= br)
                RowBreaks[i] = br + edit.Shift;
        }
    }
```

The conditional-format and data-validation bodies (`:127-155`, `:176-206`) each compute an
`affected` area. **Check whether `affected` equals `edit.Area`** — the arithmetic at `:149-151` is
`new Area(first.Row, first.Col, first.Row + rowsShifted - 1, last.Col)`, which is *not* the same as
`ExtendBelow(shift - 1)` when the range is more than one row tall. If they differ, keep the
existing computation in the adapter and say so in a comment; do not substitute `edit.Area` because
it looks equivalent. **A silent substitution here is exactly the class of defect this spec family
keeps finding.**

- [ ] **Step 3: Convert the workbook-scoped two**

Defined names (`:285-319`) and data-validation criteria formulas (`:239-271`) visit every worksheet.
They become listeners that are *yielded for every sheet* and do **not** guard on `edit.Sheet` — the
formula shifter already ignores references to sheets other than the edited one
(`XLCellFormulaShifter.cs:109` passes `shiftedRange.Worksheet.Name`; the remarks at
`XLWorksheetRangeShifter.cs:234-237` document the same thing). Copy that reasoning into the
adapters' remarks so the missing guard reads as deliberate.

- [ ] **Step 4: Convert sparkline cleanup, and leave the sparkline *shift* alone**

`RemoveInvalidSparklines` (`:273-283`) moves onto `XLSparklineGroups`. The other half —
`XLSparklineGroups.ShiftRows` / `ShiftColumns`, called from `XLRangeInsertHelper.cs:21`, `:128` and
`XLRangeBase.cs:1138`, `:1146` — is a **different dispatch point, upstream of the shifter**, and is
out of scope. Note it in the adapter's remarks as the one place a sheet feature is still notified
twice from two different layers, and leave it for a follow-on.

- [ ] **Step 5: Confirm the shifter names no feature**

Run: `grep -nE 'MergedRanges|ConditionalFormats|DataValidations|PageSetup|Sparkline|DefinedNames|Hyperlinks|CalcEngine' XLibur/Excel/XLWorksheetRangeShifter.cs`
Expected: no output.

Run: `wc -l XLibur/Excel/XLWorksheetRangeShifter.cs`
Expected: under 60 lines. The file is 320 at `1b41cadd`; spec 26 task 8 leaves it at ~210 by
collapsing the six mirror pairs; this task removes what is left except the `SheetEdit` construction
and the `foreach`.

- [ ] **Step 6: Build and run the full suite**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS, with **no assertion weakened**, including all four `*_does_not_move_yet` tests —
this task changes no behaviour.

- [ ] **Step 7: Commit**

```bash
git add XLibur/Excel XLibur.Tests/Excel/Ranges
git commit -m 'refactor(shift): the six hardcoded sheet features become listeners (spec 33 task 3)'
```

---

### Task 4 — Chart and note anchors (behaviour change)

**Files:**
- Create: `XLibur/Excel/Drawings/DrawingAnchorListener.cs`
- Modify: `XLibur/Excel/XLWorksheet.cs` — yield it
- Modify: `XLibur.Tests/Excel/Ranges/SheetListenerCharacterizationTests.cs` — re-point two tests
- Modify: `CHANGELOG.md`

- [ ] **Step 1: Answer the Excel question before writing code**

Charts and notes have two anchor points each (`Position` and `SecondPosition`; a note's `Position`
plus its cell). The correct behaviour is not obvious and **must be established against Excel, not
assumed**. Open Excel, build each case, do the edit, read the result back:

| Question | Recorded answer |
|---|---|
| Chart two-cell anchored `C4:J20`, insert 3 rows at row 1 — does it move, grow, or stay? | |
| Same chart, insert 3 rows at row 10 (inside it) — does it grow, or move? | |
| Same chart with `MoveWithCells` (not `MoveAndSizeWithCells`) — different answer? | |
| Chart one-cell anchored, insert above — move? | |
| Chart absolutely anchored — does anything move it? | |
| Delete rows that contain the whole chart anchor — is the chart deleted, or clamped? | |
| Note on `C10`, insert 3 rows at row 1 — where is the callout box? | |

`XLDrawingAnchor` already models `MoveAndSizeWithCells` (`XLChart.cs:87`), so the answer to row 3
determines whether the adapter needs to read it. **Record every answer in this spec's Results
section.** An adapter that guesses is worse than no adapter, because it produces confidently wrong
output.

- [ ] **Step 2: Write the adapter**

```csharp
/// <summary>
/// Moves the anchors of every chart and every note on the sheet when rows or columns are inserted or
/// deleted.
/// </summary>
/// <remarks>
/// Before spec 33 nothing notified these: <see cref="XLDrawingPosition"/> holds raw <c>int</c>s
/// (<c>XLDrawingPosition.cs:5,10</c>) and neither the chart collection nor the note slice
/// implemented <see cref="ISheetListener"/>, so a chart anchored at row 10 stayed at row 10 when
/// three rows were inserted above it. Notes were half-broken rather than wholly: the note moved with
/// the misc slice while its callout box did not, which <c>XLComment.cs:8-13</c> already documented.
/// <para>
/// Anchor semantics follow Excel; the cases were established by hand and recorded in
/// <c>docs/specs/33-sheet-listener-seam.md</c> task 4.
/// </para>
/// </remarks>
internal sealed class DrawingAnchorListener(XLWorksheet worksheet) : ISheetListener
```

Both anchor points are single cells, so both are `Area.At(point)` transforms and the existing
`Area.TryInsertAreaAndShiftDown` family (`Area.cs:443`, `:497`, `:558`, `:624`) applies. Use it
rather than writing new arithmetic — that is what keeps this adapter narrow.

- [ ] **Step 3: Re-point task 1's two tests**

`A_chart_anchor_does_not_move_yet` → `A_chart_anchor_moves`, asserting 13 and 23.
`A_note_anchor_does_not_move_yet` → `A_note_anchor_moves`, asserting the note on `C13` **and**
`Position.Row == 12`.

Rename the tests. A test whose name still says `does_not_move_yet` while asserting that it moves is
worse than no test.

- [ ] **Step 4: Round-trip, not just in-memory**

An in-memory assertion does not prove the file is right. Add one test that saves and reloads:

```csharp
[Test]
public async Task A_chart_anchor_survives_a_row_insert_through_a_round_trip()
```

Anchor a chart, insert rows, `SaveAs(ms, validate: true)`, reopen, assert the anchor. Follow
`ChartAnchorTests.cs` — it already has `SaveValidated` and `DrawingOf` helpers for reading the
`xdr:twoCellAnchor` back out of the package.

- [ ] **Step 5: Changelog**

This changes output for existing documents. Under `### Fixed`:

```
- Chart and note anchors now move when rows or columns are inserted or deleted above or to the left
  of them. Previously a chart anchored at row 10 stayed at row 10 after an insert at row 1, and a
  note's callout box detached from the note. (spec 33)
```

- [ ] **Step 6: Build and run the full suite**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`

Expected: PASS. **If an existing test fails**, it may have been pinning the old wrong behaviour.
Read it before changing it: if it was asserting an anchor position after an edit, re-point it and
say so in the commit body. If it was asserting something else, you have a real regression.

- [ ] **Step 7: Commit**

```bash
git add XLibur/Excel/Drawings/DrawingAnchorListener.cs XLibur/Excel/XLWorksheet.cs XLibur.Tests/Excel/Ranges/SheetListenerCharacterizationTests.cs CHANGELOG.md
git commit -m 'fix(shift): chart and note anchors move with structural edits (spec 33 task 4)'
```

Commit body must name every assertion that was deliberately reversed.

---

### Task 5 — Panes and pivot `Area` (behaviour change)

**Files:**
- Modify: `XLibur/Excel/XLSheetView.cs`, `XLibur/Excel/PivotTables/XLPivotTable.cs`
- Modify: `XLibur/Excel/XLWorksheet.cs`
- Modify: `XLibur.Tests/Excel/Ranges/SheetListenerCharacterizationTests.cs`, `CHANGELOG.md`

- [ ] **Step 1: Answer the Excel questions**

| Question | Recorded answer |
|---|---|
| Panes frozen at row 5, insert 3 rows at row 1 — is the freeze at row 5 or row 8? | |
| Panes frozen at row 5, insert 3 rows at row 10 (below the split) — does the split move? | |
| Panes frozen at row 5, delete rows 2–3 — row 3, or unchanged? | |
| Delete every row above the split — does the freeze disappear? | |
| Pivot table at `D10`, insert 3 rows at row 1 — does the whole table move to `D13`? | |
| Pivot table at `D10`, insert a row *inside* the table — does Excel allow it, and what happens? | |
| Pivot table whose *source* range is edited — separate concern, or the same one? | |

`SplitRow` is a **count**, not an address, which makes it the one feature in this spec whose
transform is not an area transform. Model the frozen region as `Area(1, 1, SplitRow, SplitColumn)`
and note that `Area.TryInsertAreaAndShiftDown` returns `false` for a partially covering insert
("Partial cover, don't move" — `XLHyperlinks.cs:73`), which is the **wrong** answer for a pane that
should grow. Do the arithmetic in the adapter. That is the correct place for it and is why the port
does not widen further.

- [ ] **Step 2: Implement both**

`XLSheetView` implements `ISheetListener` directly — it is already sheet-owned
(`XLSheetView.cs:80`). `XLPivotTable` likewise; its `Area` setter already exists (`:845`) and
`TargetCell` derives from it (`:59-70`), so moving `Area` moves the table.

The enumeration yields one listener per pivot table. `XLPivotTables` is a collection, so
`foreach (var pt in PivotTables.Cast<XLPivotTable>()) yield return pt;` — matching how
`GetWorkbookListeners` yields one listener per defined name (`XLWorksheets.cs:240-249`).

- [ ] **Step 3: Re-point task 1's two tests, and rename them**

`Freeze_panes_do_not_move_yet` → `Freeze_panes_move`, asserting 8 and 6.
`A_pivot_area_does_not_move_yet` → `A_pivot_area_moves`, asserting row 13.

- [ ] **Step 4: Changelog**

```
- Frozen/split panes and pivot table positions now move when rows or columns are inserted or deleted
  above or to the left of them. (spec 33)
```

- [ ] **Step 5: Build and run the full suite on both frameworks**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj`
Expected: PASS on net8.0 and net10.0.

Pay attention to the pivot suites, which are the most likely to have pinned the old behaviour:

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/*Pivot*/*"`

- [ ] **Step 6: Commit**

```bash
git add XLibur/Excel XLibur.Tests/Excel/Ranges CHANGELOG.md
git commit -m 'fix(shift): panes and pivot tables move with structural edits (spec 33 task 5)'
```

---

### Task 6 — Delete the `XLMarker` workaround

With task 4's adapter live, a picture anchor no longer needs to be a range to get moved. This task
is what proves the seam replaced the trick rather than sitting beside it.

**Files:**
- Modify: `XLibur/Excel/Drawings/XLMarker.cs`
- Modify: `XLibur/Excel/Drawings/DrawingAnchorListener.cs` — pictures join charts and notes
- Modify: `XLibur/Excel/Drawings/XLPicture.cs` — `Markers` construction, `:42-45`, `:63`, `:104`,
  `:143`, `:156`

- [ ] **Step 1: `XLMarker` holds a `Point`**

Delete `rangeCell` and the `IXLRange` constructor. Keep the one-cell invariant as an assertion on
the incoming cell if it still means anything; the `RowCount() != 1` check at `:26-27` becomes
vacuous once the field is a point and should go.

`Cell` (`:33`) currently returns `rangeCell.FirstCell()`. It becomes
`_worksheet.Internals.CellsCollection.GetCell(_point)` or equivalent — `XLMarker` needs the
worksheet, which it does not hold today. Add it as a constructor parameter; every call site in
`XLPicture` already has one.

- [ ] **Step 2: Register pictures with the anchor listener**

Both markers per picture, both transformed the same way as a chart's two positions. Answer the
Excel question the same way: does a picture anchored `C4:J20` grow or move when a row is inserted at
row 10? Record it.

- [ ] **Step 3: The picture tests must still pass unchanged**

The picture was the control in this spec's evidence — it moved correctly before. It must still move
correctly, through a different mechanism.

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/*Picture*/*"`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/*Image*/*"`
Expected: PASS, **no assertion weakened**. Unlike tasks 4 and 5, this task changes no behaviour, so
any red test is a regression.

- [ ] **Step 4: Confirm the workaround is gone**

Run: `grep -n 'IXLRange\|rangeCell\|range repository' XLibur/Excel/Drawings/XLMarker.cs`
Expected: no output.

- [ ] **Step 5: Full suite, then commit**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`

```bash
git add XLibur/Excel/Drawings
git commit -m 'refactor(drawings): XLMarker holds a point, not a smuggled range (spec 33 task 6)'
```

---

### Task 7 — Confirm the structural-edit cost is unchanged

Spec 05 measured the range-shift pass at **8%** of a 1,000-insert workload (359 ms of 4,753 ms) and
formula shifting at **68%** (3,244 ms). This spec restructures the 8%. It must not move the total.

- [ ] **Step 1: Measure the merge-base**

```
dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile structural
```

Record all four probes from spec 05's Results table: empty sheet, formula shift, range-shift pass,
and the unreachable-range sub-probe.

- [ ] **Step 2: Measure the branch**

Same command, same fixture, three runs. **The machine has ~40% run-to-run timing variance** — see
`benchmark-machine-noise`. Compare medians, never single runs.

- [ ] **Step 3: Decide**

Expected: range-shift pass within noise of 359 ms; total within noise of 4,753 ms.

The enumeration replaces nine direct calls with an iterator yielding twelve-plus listeners per edit,
plus one `SheetEdit` construction. The struct is passed `in` and does not allocate. The iterator
does: one state machine per `ShiftRows`/`ShiftColumns` call, which on a 1,000-insert workload is
1,000 allocations. **If that shows up, the fix is to build the listener list once per worksheet and
cache it, invalidating on pivot-table add/remove** — not to unwind the seam. Do not pre-emptively
cache; measure first.

A regression above the noise floor must be explained before this spec lands, not after.

- [ ] **Step 4: Record the numbers**

```bash
git add docs/specs/33-sheet-listener-seam.md
git commit -m 'docs(specs): record structural-edit numbers and the Excel anchor decisions for spec 33'
```

---

## Acceptance criteria

1. `XLWorksheetRangeShifter.cs` names no feature. Gate:
   `grep -nE 'MergedRanges|ConditionalFormats|DataValidations|PageSetup|Sparkline|DefinedNames|Hyperlinks|CalcEngine' XLibur/Excel/XLWorksheetRangeShifter.cs`
   returns nothing.
2. At least **12** types implement the port. Gate:
   `grep -rc 'ISheetListener$\|ISheetListener,\|, ISheetListener' XLibur --include=*.cs` summed over
   files, or simply that the type list in `GetSheetListeners()` has 12 or more distinct types.
3. `ISheetListener` still has **exactly four** methods. Gate:
   `grep -c '    void On' XLibur/Excel/Cells/ISheetListener.cs` returns 4.
4. `GetSheetListeners()` is the only place a sheet listener is named. Gate:
   `grep -rn 'ISheetListener' XLibur --include=*.cs` shows no call site outside
   `XLWorksheetRangeShifter.cs`'s single `foreach` and the declarations themselves.
5. A chart anchored below an inserted row moves — in memory **and** through a save/reload round trip.
6. A note's callout box and the note itself end up on the same cell after an edit.
7. `SplitRow`/`SplitColumn` and `XLPivotTable.Area` move.
8. `XLMarker.cs` contains no `IXLRange`. Gate: `grep -c 'IXLRange' XLibur/Excel/Drawings/XLMarker.cs`
   returns 0.
9. Listener order is pinned by a test, and matches the order recorded on the branch point (or spec
   26's reconciled order, whichever applies).
10. Every Excel-behaviour question in tasks 4 step 1, 5 step 1 and 6 step 2 has a recorded answer in
    a Results section.
11. Full suite green on net8.0 and net10.0, **with no assertion weakened** — except the four
    `*_does_not_move_yet` tests re-pointed by tasks 4 and 5, each named in its commit body.
12. Structural-edit profile within spec 05's noise of its pre-spec value.
13. No public API change. `PublicAPI.Unshipped.txt` untouched.
14. Two changelog entries under `### Fixed`, one for tasks 4 and 6, one for task 5.

## Conflicts

- **Spec 26 (grid axis) — hard, and 26 runs first.** Both specs own
  `XLibur/Excel/XLWorksheetRangeShifter.cs` and `XLibur/Excel/XLWorksheet.cs`. **26 collapses the
  row/column duplication in both files before 33 reorganises what is left.** The shifter has six
  mirror pairs today (`ShiftPageBreaksColumns`/`Rows`, `ShiftConditionalFormattingColumns`/`Rows`,
  `ShiftDataValidationColumns`/`Rows`, `ShiftDataValidationFormulaColumns`/`Rows`,
  `MoveDefinedNamesColumns`/`Rows`, and the two merged-range blocks). Running 33 first means writing
  every adapter twice, once per axis, and then deleting half of each when 26 lands. Spec 26 task 9
  is explicitly "`XLWorksheetRangeShifter`, and the ordering question" and takes the file from 320
  lines to ~210; this spec starts from that output. 26 task 8 also settles the
  page-break/sparkline ordering discrepancy documented above, having established that the two
  commute — **33 inherits that decision and does not re-open it.** 26 states the dependency from its
  own side (`docs/specs/26-grid-axis.md:5-6`) and explicitly leaves `ISheetListener` alone: it
  collapses the *callers* at `:43-57` and `:86-100`, not the interface.
- **Spec 05 (structural-edit scalability) — done, and 33's direct ancestor.** Read its Results
  before starting. Two things matter here: the range-repository spatial index was **declined on
  measurement** (a filter made unreachable ranges free, so the enumeration an index would remove was
  never the expense), and formula shifting is **68%** of the insert workload against the range-shift
  pass's **8%**. 33 does not re-enter the index decision and does not touch the 68%. Task 7's
  numbers come from 05's own profiler.
- **Spec 14 (`Clear`/`CopyTo` scalability) — adjacent, soft, either order.** 14's fix is inside
  `XLRangeBase.Clear` (a data validation created and deleted on every call). 33 touches
  `XLRangeBase` not at all, and the `XLDataValidations` type only to add a listener implementation.
  The one thing to watch: if 14 lands first and changes when data validations exist, task 1's
  characterization test may see a different validation count. Re-run task 1 after rebasing rather
  than assuming.
- **Spec 25 (formula shifter seam) — no conflict.** 25 owns `XLCellFormulaShifter*.cs`, which this
  spec calls (through the defined-name and criteria-formula adapters) and does not modify.
- **Spec 24 — no conflict.** Load path only.
- **Spec 21 — done, and one constraint survives it:** `Point`, `Area` and `XLRangeAddress` are
  structs, and `SheetEdit` must not undo that. It is a `readonly struct` passed `in`; task 7 is what
  confirms it did not become an allocation.
- Shared documentation files: `docs/specs/README.md` and
  `docs/specs/TASKLIST-architecture-deepening-2.md`. Expect trivial merge conflicts with the other
  round-2 specs; resolve by keeping both edits.
