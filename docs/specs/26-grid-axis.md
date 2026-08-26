# Spec 26 — Give the grid one axis

**Area:** Architecture · Refactor · **Defect (×2 shipped, ×1 latent)**
**Effort:** M (~5–6 days)
**Dependencies:** None hard. **Must run before spec 33** (sheet-listener seam) — it halves
`XLWorksheetRangeShifter.cs` and `XLWorksheet.cs` before 33 reorganises them. See Conflicts.
**Status:** ✅ Merged — PR #409 as `2b244064`, 2026-08-26. See Results.

## Goal

Every row-wise algorithm in the grid has a column-wise twin written out longhand. Introduce an
internal `IGridAxis` so there is **one** implementation and two thin adapters, and fix the three
defects the second copy has already produced.

## Why this spec exists

The duplication is not stylistic. It is line-for-line, and it has drifted three times in shipped
code.

### The mirror, measured

| File | Lines | Column copy | Row copy | Second copy costs |
|---|---:|---|---|---:|
| `XLibur/Excel/Ranges/XLRangeInsertHelper.cs` | 226 | `:13-118` | `:120-225` | 106 |
| `XLibur/Excel/Ranges/XLRangeBase.cs` (insert block) | 1,384 | `:884-980` | `:982-1078` | 97 |
| `XLibur/Excel/Ranges/XLRangeShiftHelper.cs` | 144 | `:12-76` | `:78-143` | 66 |
| `XLibur/Excel/XLWorksheetRangeShifter.cs` | 320 | 6 pairs | 6 pairs | ~115 |
| `XLibur/Excel/XLWorksheet.cs` (shift block) | 1,939 | `:1280-1308`, `:1373-1403` | `:1248-1278`, `:1333-1367` | ~60 |

**≈444 lines exist only because the same algorithm is written twice.**

`XLRangeInsertHelper.cs` is the clearest case: a 226-line file that is two 106-line copies plus a
14-line header. Six method pairs, in the same order, with the same shape:

| Column | Row |
|---|---|
| `InsertColumnsBefore` (`:13`) | `InsertRowsAbove` (`:120`) |
| `ShiftFormulasForColumns` (`:52`) | `ShiftFormulasForRows` (`:159`) |
| `ShiftColumnWidths` (`:57`) | `ShiftRowHeights` (`:164`) |
| `ApplyColumnFormatting` (`:75`) | `ApplyRowFormatting` (`:182`) |
| `ApplyColumnFormattingFromLeft` (`:83`) | `ApplyRowFormattingFromAbove` (`:190`) |
| `ApplyColumnFormattingFromExistingRows` (`:100`) | `ApplyRowFormattingFromExistingColumns` (`:207`) |

`XLRangeBase.cs:884-980` and `:982-1078` are 97 lines each, and their ten methods have identical
line counts in identical order: **4 / 21 / 4 / 4 / 19 / 4 / 21 / 4 / 4 / 3**.

`XLWorksheet.cs` says the quiet part out loud. The doc comment at `:1369-1372` reads:

```csharp
    /// <summary>
    /// Column-wise counterpart of <see cref="CollectRangesShiftedByRows"/>; see it for why filtering
    /// here is safe.
    /// </summary>
```

Thirty-five lines of reasoning live on `CollectRangesShiftedByRows` (`:1333-1367`) and its twin
(`:1373-1403`) carries a pointer instead. The reasoning is correct in one place and unstated in the
other, which is the failure mode this spec removes.

`XLWorksheetRangeShifter.cs` is 6 mirror pairs — `:17`↔`:60`, `:103`↔`:115`, `:127`↔`:141`,
`:176`↔`:192`, `:239`↔`:259`, `:285`↔`:303` — against only 3 shared methods (`:164`, `:216`,
`:273`).

Across `XLibur/Excel` as a whole, **100 declared member names have a Row/Column mirror twin; 91 of
those pairs are declared inside a single file, spread over 65 files.**

Reproduce:

```
python -c "
import re,os,collections
decl=re.compile(r'^\s*(?:\[[^\]]*\]\s*)*(?:(?:public|private|internal|protected|private protected|protected internal)\s+)?(?:(?:static|sealed|override|virtual|new|readonly|async|partial|abstract|extern|unsafe|required|const|ref|event)\s+)*[\w<>\[\],\.\?\(\) ]+?\s+([A-Za-z_][A-Za-z0-9_]*)\s*(\(|=>|\{)')
names=collections.defaultdict(set)
for root,d,fs in os.walk('XLibur/Excel'):
    for f in fs:
        if f.endswith('.cs'):
            p=os.path.join(root,f)
            for line in open(p,encoding='utf-8-sig',errors='replace'):
                m=decl.match(line)
                if m: names[m.group(1)].add(p)
pairs={tuple(sorted((n,n.replace('Rows','Q_').replace('Row','Column').replace('Q_','Columns'))))
       for n in names if 'Row' in n and n.replace('Rows','Q_').replace('Row','Column').replace('Q_','Columns') in names}
pairs={p for p in pairs if p[0]!=p[1]}
same=[p for p in pairs if names[p[0]]&names[p[1]]]
print(len(pairs),'pairs;',len(same),'same-file;',len({f for p in same for f in names[p[0]]&names[p[1]]}),'files')"
```

**Correction to the survey that seeded this spec.** It recorded "101 pairs across 34 files". The
pair count is right within one; the file count is not — 65 files declare both halves of at least one
pair, not 34. The larger number is the honest one and it strengthens the case, so it stands.

The top offenders by same-file pair count are `XLWorksheet.cs` (18), `XLPivotTable.cs` (12),
`IXLWorksheet.cs` (11), `IXLPivotTable.cs` (11), `IXLRange.cs` (10),
`XLCellFormulaShifter.Legacy.cs` (9), `XLRange.cs` (9), `XLRangeBase.cs` (7),
`XLCellsCollection.cs` (6), `XLWorksheetRangeShifter.cs` (6). This spec takes four of those ten.

### Character-identical drift

`XLRange.cs:365-378` (`FirstColumn`) and `:481-494` (`FirstRow`) are the same eleven lines with the
words swapped, and their accessibility has already diverged:

```csharp
    internal XLRangeColumn? FirstColumn(Func<IXLRangeColumn, bool>? predicate = null)   // :365
    public   XLRangeRow?    FirstRow(Func<IXLRangeRow, bool>? predicate = null)         // :481
```

`XLRange` is `internal sealed`, so `public` there is `internal` in effect. Nobody chose this; one
copy was edited and the other was not.

### Three defects the second copy has already shipped

All three are confirmed by direct read at `1b41cadd`.

---

**Defect 1 — row outline levels are written into the column counter.**

`XLibur/Excel/Rows/XLRow.cs:416-428`:

```csharp
    public int OutlineLevel
    {
        get => _outlineLevel;
        set
        {
            if (value is < 0 or > 8)
                throw new ArgumentOutOfRangeException(nameof(value), "Outline level must be between 0 and 8.");

            Worksheet.IncrementColumnOutline(value);      // :424  <-- Column, on a row
            Worksheet.DecrementColumnOutline(_outlineLevel); // :425
            _outlineLevel = value;
        }
    }
```

Character-identical to `XLibur/Excel/Columns/XLColumn.cs:334-346`, `Increment**Column**Outline`
included.

`XLOutlineTracker.IncrementRowOutline` (`XLibur/Excel/XLOutlineTracker.cs:41`) and
`DecrementRowOutline` (`:50`) therefore have **zero callers repo-wide**. Gate:

```
grep -rn "IncrementRowOutline\|DecrementRowOutline" --include=*.cs .
```

returns only the two definitions and the two `XLWorksheet` pass-throughs (`XLWorksheet.cs:1112`,
`:1114`), which themselves have no callers.

Three shipped consequences:

1. `GetMaxRowOutline()` always returns 0, so `sheetFormatPr/@outlineLevelRow`
   (`XLibur/Excel/IO/SheetViewWriter.cs:275-278`) is never emitted.
2. Grouping rows inflates `@outlineLevelCol` (`SheetViewWriter.cs:270-273`).
3. It is a **round-trip corruption**, not just a create-path one:
   `XLibur/Excel/IO/WorksheetSheetDataReader.cs:1087-1088` sets `xlRow.OutlineLevel` on load, so
   opening any file with row groups and saving it inflates that file's `@outlineLevelCol`.

Per-row `r/@outlineLevel` is unaffected — it is written directly from `xlRow.OutlineLevel`
(`XLibur/Excel/IO/SheetDataWriter.cs:580-581`) and read back at `WorksheetSheetDataReader.cs:1088`.
Only the sheet-level summary attribute is wrong.

---

**Defect 1b (latent, unlocked by fixing 1) — `GetMaxRowOutline` throws.**

The two tracker accessors are also not the same shape (`XLOutlineTracker.cs:35-39` vs `:62-65`):

```csharp
    public int GetMaxColumnOutline()
    {
        var list = _columnOutlineCount.Where(kp => kp.Value > 0).ToList();
        return list.Count == 0 ? 0 : list.Max(kp => kp.Key);      // guards the filtered set
    }

    public int GetMaxRowOutline()
    {
        return _rowOutlineCount.Count == 0 ? 0 : _rowOutlineCount.Where(kp => kp.Value > 0).Max(kp => kp.Key);
    }                          // ^ guards the unfiltered dictionary
```

`GetMaxRowOutline` guards the dictionary's size, not the filtered sequence's. A dictionary holding
only zero counts is non-empty, so `.Max()` runs on an empty sequence and throws
`InvalidOperationException`. Today `_rowOutlineCount` is always empty, so this is unreachable.
**Fixing defect 1 makes it reachable on the very sequence an existing test already performs** —
`XLibur.Tests/Excel/Rows/RowTests.cs:274-283`:

```csharp
        ws.Rows(1, 2).Group();      // _rowOutlineCount -> {1: 2}
        ws.Rows(1, 2).Ungroup(true); // _rowOutlineCount -> {1: 0}  (non-empty, all zero)
```

That test never saves, so it would not catch it. A save would:
`WorksheetPartWriter.cs:174-176` guards the call with `xlWorksheet.RowCount() > 0`, and
`XLRangeBase.RowCount()` (`:669-672`) on a worksheet returns 1,048,576, so `GetMaxRowOutline()` is
called on **every** save.

Defect 1 and 1b must land in the same commit or the fix ships a crash.

---

**Defect 2 — `XLColumn.CellCount()` always returns 1.**

`XLibur/Excel/Columns/XLColumn.cs:404-407` is character-identical to
`XLibur/Excel/Rows/XLRow.cs:486-489`:

```csharp
    public int CellCount()
    {
        return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
    }
```

`XLRangeAddress.EntireColumn` (`XLibur/Excel/Ranges/XLRangeAddress.cs:15-20`) builds
`(1, c)`–`(MaxRow, c)`. Both addresses carry the same column, so the expression is `c - c + 1 = 1`.
`XLRangeAddress.EntireRow` (`:22-27`) builds `(r, 1)`–`(r, MaxCol)`, where the same expression is
correctly 16,384.

`IXLColumn.CellCount()` should return 1,048,576. `XLRangeBase.RowCount()` already computes exactly
that, so the fix is one line.

**Premise that could be wrong.** Neither `IXLColumn.cs:162` nor `IXLRow.cs:169` documents what
`CellCount` means; both are bare `int CellCount();`. The reading above — a column has as many cells
as the sheet has rows, symmetric with the row — is the only one consistent with `IXLRow`, but it is
an inference. **Task 2 is what settles it**; if the suite disagrees, record that and stop.

`IXLColumn.CellCount() -> int` is `PublicAPI.Shipped.txt:385`, so this changes a shipped return
value. No test asserts it: the only `CellCount` assertion in the suite is
`XLibur.Tests/Excel/Ranges/RangeRowCopyToTests.cs:112`, on `IXLRangeRow`.

---

**Drift 3 — inert, not a defect. The survey overstated this one.**

Two smaller divergences were reported as shipped defects. Neither is:

- **`Cells(string)` flags.** `XLRow.cs:249` passes `XLCellsUsedOptions.AllContents`;
  `XLColumn.cs:102` passes `.All`. Both pass `usedCellsOnly: false`, and `XLCells` consults
  `_options` only from the `GetUsedCells*` paths — `XLCells.cs:149`, `:179`, `:200-217` — which
  `GetAllCells()` (`:36-50`) never enters. **The drift produces no observable difference today.** It
  is a live trap, not a live bug: the day someone flips that flag the two axes behave differently
  for no stated reason.
- **Page-break / sparkline ordering.** `XLWorksheetRangeShifter.ShiftColumns` runs
  `ShiftPageBreaksColumns` then `RemoveInvalidSparklines` (`:40-41`); `ShiftRows` runs them in the
  opposite order (`:83-84`). Nothing in the file says whether the order matters.
  `RemoveInvalidSparklines` (`:273-283`) reads only sparkline location validity and
  `ShiftPageBreaks*` (`:103-125`) touches only `PageSetup.*Breaks`, so on inspection they commute —
  but "on inspection" is exactly what one implementation would make unnecessary. Collapsing the pair
  forces a single order and a comment saying so.

### Why an axis, and why it will work here

The codebase already contains the shape. `XLRangeBase.Delete` (`:1104-1161`) handles both axes by
branching **once** on `XLShiftDeletedCells` at `:1132-1149` and sharing everything else — formula
shift, merges, notification, repository delete. It is 58 lines for both axes. The insert path, doing
the same job in the other direction, is 195 lines in `XLRangeBase` plus 226 in the helper.

The axis abstraction is the `Delete` pattern promoted from a local `switch` to a named type.

## Non-goals

- **`XLibur/Excel/Cells/XLCellFormulaShifter.Legacy.cs` is out of scope.** It carries 9 further
  mirror pairs — the third-largest concentration in the tree — and it is **spec 25's file**. Leaving
  it alone keeps this spec file-disjoint from 25, which is worth more than the 9 pairs.
- **`ISheetListener` is out of scope.** Its four members are two mirror pairs
  (`OnInsertAreaAndShiftDown`/`Right`, `OnDeleteAreaAndShiftUp`/`Left`,
  `XLibur/Excel/Cells/ISheetListener.cs`) and it is **spec 33's interface**. This spec collapses the
  *callers* in `XLWorksheetRangeShifter.cs:43-57` and `:86-100`, not the interface.
- **Pivot tables are out of scope.** `XLPivotTable.cs` has 12 same-file pairs, but `RowAxis` /
  `ColumnAxis` there are pivot-model concepts with genuinely different semantics (`DataOnRows`,
  grand totals), not grid-geometry mirrors.
- **Not a performance spec.** Task 8 exists only to show the cost did not move.
- **No public API change.** No entry in `PublicAPI.Shipped.txt` or `PublicAPI.Unshipped.txt` is
  added, removed, or retyped. Defect 2 changes a shipped *value*, not a shipped *signature*.

## Current state

Verified against the tree at `1b41cadd` (2026-08-24).

Every line number below was read at that commit.

**The five algorithm files**

- `XLibur/Excel/Ranges/XLRangeInsertHelper.cs` — 226 lines; column block `:13-118`, row block
  `:120-225`; `internal static` entry points at `:13` and `:120`, called from
  `XLRangeBase.cs:980` and `:1078`
- `XLibur/Excel/Ranges/XLRangeBase.cs` — insert block `:884-1078`; `RowCount()` `:669-672`,
  `ColumnCount()` `:691-694`; `IsEntireRow()` `:542-545`, `IsEntireColumn()` `:547-550`;
  `Delete` `:1104-1161` (the shape to copy); `ShiftColumns` `:1173-1174`, `ShiftRows` `:1176-1177`
- `XLibur/Excel/Ranges/XLRangeShiftHelper.cs` — 144 lines; `ShiftColumns` `:12-76`,
  `ShiftRows` `:78-143`; both are pure except `worksheet.DeleteRange` at `:58` / `:125`
- `XLibur/Excel/XLWorksheetRangeShifter.cs` — 320 lines; `internal sealed class` with a primary
  constructor at `:15`; pairs `:17`↔`:60`, `:103`↔`:115`, `:127`↔`:141`, `:176`↔`:192`,
  `:239`↔`:259`, `:285`↔`:303`; shared `:164`, `:216`, `:273`
- `XLibur/Excel/XLWorksheet.cs` — `NotifyRangeShiftedRows` `:1248-1278`,
  `NotifyRangeShiftedColumns` `:1280-1308`; `CollectRangesShiftedByRows` `:1333-1367`,
  `CollectRangesShiftedByColumns` `:1373-1403`; outline pass-throughs `:1106-1116`

**The defect sites**

- `XLibur/Excel/Rows/XLRow.cs:424-425` — the two `Column` calls on a row
- `XLibur/Excel/XLOutlineTracker.cs:35-39` / `:62-65` — the asymmetric guards
- `XLibur/Excel/Columns/XLColumn.cs:404-407` / `XLibur/Excel/Rows/XLRow.cs:486-489` — `CellCount`
- `XLibur/Excel/Ranges/XLRangeAddress.cs:15-27` — `EntireColumn` / `EntireRow`, which is why
  `CellCount` is wrong
- `XLibur/Excel/IO/SheetViewWriter.cs:270-278` — where `@outlineLevelCol` / `@outlineLevelRow` are
  decided
- `XLibur/Excel/IO/WorksheetPartWriter.cs:170-179` — the two `GetMax*Outline()` calls
- `XLibur/Excel/IO/WorksheetSheetDataReader.cs:1087-1088` — the load path that makes defect 1 a
  round-trip corruption
- `XLibur/Excel/IO/SheetDataWriter.cs:580-581` — per-row `outlineLevel`, correct today

**The adapters**

- `XLibur/Excel/Rows/XLRow.cs` — 702 lines, `internal sealed class XLRow : XLRangeBase, IXLRow`;
  `RangeAddress` `:43-47`; `_flags` bitfield `:20`, `IsHidden` via flags `:404-414`
- `XLibur/Excel/Columns/XLColumn.cs` — 601 lines, `internal sealed class XLColumn : XLRangeBase,
  IXLColumn`; `RangeAddress` `:37-41`; `IsHidden` a plain auto-property `:332`

The `IsHidden` divergence is deliberate — the row carries a flags bitfield because there are up to
1,048,576 of them — and the adapters must preserve it. It is the one place where "the two are the
same" is false.

**Interface widths** (counted from `XLibur/PublicAPI.Shipped.txt`)

| Interface | Own members |
|---|---:|
| `IXLRangeBase` | 66 |
| `IXLRange` | 65 |
| `IXLRow` | 47 |
| `IXLColumn` | 44 |

Gate: `grep -cE '^XLibur\.Excel\.IXLRow\.' XLibur/PublicAPI.Shipped.txt` → 47.

**Types affected are all internal.** `grep -nE '(^|[^A-Za-z])XL(Row|Column|Range|RangeBase|RangeInsertHelper|RangeShiftHelper|WorksheetRangeShifter|OutlineTracker)\b' XLibur/PublicAPI.Shipped.txt`
returns nothing.

**Language level.** `XLibur/XLibur.csproj:4` sets `<LangVersion>default</LangVersion>`, overriding
`Directory.Build.props:33`'s `latest`. `default` resolves per-TFM against
`net10.0;net8.0;net9.0` (`XLibur.csproj:5`), so the floor is **C# 12** (net8.0). Generic struct
constraints, `in` parameters and `static abstract` interface members are all available.

**Prior art in the tree.** `where T : struct` generic constraints are already used —
`XLibur/Excel/Caching/IXLRepository.cs:18`, `XLibur/Excel/Style/XLStyleValue.cs:129`,
`XLibur/Excel/IO/ColumnWriter.cs:269`. `static abstract` interface members are **not** used
anywhere yet; this spec would be the first, which is why the design below does not require them.

## File structure

```
XLibur/Excel/Coordinates/GridAxis.cs                   new   — IGridAxis, RowAxis, ColumnAxis
XLibur/Excel/Ranges/XLRangeInsertHelper.cs             modified — 226 -> ~130, one implementation
XLibur/Excel/Ranges/XLRangeShiftHelper.cs              modified — 144 -> ~90
XLibur/Excel/Ranges/XLRangeBase.cs                     modified — :884-1078 collapses
XLibur/Excel/XLWorksheet.cs                            modified — :1248-1403 collapses
XLibur/Excel/XLWorksheetRangeShifter.cs                modified — 320 -> ~210
XLibur/Excel/Rows/XLRow.cs                             modified — CellCount, OutlineLevel
XLibur/Excel/Columns/XLColumn.cs                       modified — CellCount
XLibur/Excel/XLOutlineTracker.cs                       modified — GetMaxRowOutline guard
CHANGELOG.md                                           modified — the two behaviour changes

XLibur.Tests/Excel/Rows/OutlineRoundTripTests.cs       new   — defects 1 and 1b
XLibur.Tests/Excel/Columns/ColumnTests.cs              modified — defect 2
XLibur.Tests/Excel/Ranges/GridAxisSymmetryTests.cs     new   — the standing gate
```

No file is deleted. No file is added under `XLibur/Excel/Cells/`.

## The design

### The axis type

`Point` is a `readonly struct` with row and column packed into one `ulong`
(`XLibur/Excel/Coordinates/Point.cs:15-56`); `Area` (`Area.cs:12`) and `XLRangeAddress`
(`XLRangeAddress.cs:9`) are `readonly struct` too. **Spec 21 put them there and this spec must not
undo it** — an `IGridAxis` reached through an interface reference would box nothing itself but would
turn every `Major(point)` into a non-inlinable interface call on the hottest path in the library.

So the axis is a **generic type parameter with a struct constraint**, not an interface reference:

```csharp
namespace XLibur.Excel.Coordinates;

/// <summary>
/// One of the grid's two axes, as a value the JIT can specialise over. Every method projects a
/// point, an area or a range address onto the axis being operated on ("index") and the axis it
/// spans ("cross"), so an algorithm can be written once and bound to either direction.
/// </summary>
/// <remarks>
/// Implementations are empty <c>readonly struct</c>s and are always passed as a generic type
/// argument constrained <c>where TAxis : struct, IGridAxis</c>. That is what makes the calls
/// devirtualise: the JIT specialises the method body per axis and the receiver's exact type is
/// known at every call site. Never hold an <see cref="IGridAxis"/> in a field, a local or a
/// parameter typed as the interface — that is the one form that reintroduces the dispatch this
/// shape exists to remove, and spec 21 already paid to learn what a mis-shaped struct costs here.
/// </remarks>
internal interface IGridAxis
{
    /// <summary>1,048,576 for the row axis, 16,384 for the column axis.</summary>
    int MaxIndex { get; }

    /// <summary>The extent of a single line on this axis: 16,384 for a row, 1,048,576 for a column.</summary>
    int MaxCross { get; }

    int IndexOf(in Point point);
    int CrossOf(in Point point);
    int IndexOf(IXLAddress address);
    int CrossOf(IXLAddress address);

    Point PointAt(int index, int cross);
    XLAddress AddressAt(XLWorksheet worksheet, int index, int cross, bool fixedIndex, bool fixedCross);

    bool IsEntireLine(XLRangeBase range);
    int MaxUsedIndex(XLCellsCollection cells);

    void InsertAreaAndShift(XLCellsCollection cells, Area area);
    void ShiftSparklines(XLSparklineGroups groups, Area area, int shift);
    void NotifyRangeShifted(XLWorksheet worksheet, XLRange range, int shift);
}

/// <summary>The row axis: rows are inserted, deleted and shifted; each row spans 16,384 columns.</summary>
internal readonly struct RowAxis : IGridAxis
{
    public int MaxIndex => XLHelper.MaxRowNumber;
    public int MaxCross => XLHelper.MaxColumnNumber;

    public int IndexOf(in Point point) => point.Row;
    public int CrossOf(in Point point) => point.Column;
    public int IndexOf(IXLAddress address) => address.RowNumber;
    public int CrossOf(IXLAddress address) => address.ColumnNumber;

    public Point PointAt(int index, int cross) => new(index, cross);
    // ... remainder mirrors, with the arguments the other way round in ColumnAxis
}
```

`ColumnAxis` is the same members with `Row` and `Column` transposed. The two adapters together are
the *entire* remaining duplication, and each member is one expression.

**Naming.** `XLPivotTable` already exposes `RowAxis` / `ColumnAxis` **properties**
(`XLibur/Excel/IO/PivotTableDefinitionPartReader.cs:94`,
`XLibur/Excel/PivotTables/Areas/XLPivotDataFields.cs:122`). They are properties, not types, so there
is no CS0104. If a file using both namespaces reads ambiguously anyway, rename the structs
`AlongRows` / `AlongColumns` — the change is mechanical and the interface name stays.

**If `IsEntireLine` / `MaxUsedIndex` / `InsertAreaAndShift` turn out to need types that
`XLibur.Excel.Coordinates` cannot see** (`XLCellsCollection`, `XLSparklineGroups` and `XLWorksheet`
live in `XLibur.Excel`), split the type: keep the geometry half (`MaxIndex` … `AddressAt`) in
`Coordinates` and put the model half in a second partial interface under `XLibur/Excel/`. Do not
force a namespace change on `Point` or `Area` to avoid the split.

### How an algorithm binds it

`XLRangeInsertHelper` becomes one method plus two three-line adapters:

```csharp
    internal static IXLRangeColumns? InsertColumnsBefore(XLRangeBase range, bool onlyUsedCells,
        int numberOfColumns, bool formatFromLeft, bool nullReturn)
        => Insert<ColumnAxis>(range, onlyUsedCells, numberOfColumns, formatFromLeft, nullReturn)
            ?.Columns();

    internal static IXLRangeRows? InsertRowsAbove(XLRangeBase range, bool onlyUsedCells,
        int numberOfRows, bool formatFromAbove, bool nullReturn)
        => Insert<RowAxis>(range, onlyUsedCells, numberOfRows, formatFromAbove, nullReturn)
            ?.Rows();

    private static IXLRange? Insert<TAxis>(XLRangeBase range, bool onlyUsedCells, int count,
        bool formatFromPrevious, bool nullReturn)
        where TAxis : struct, IGridAxis
    { /* the body of :13-50, written once */ }
```

The two entry points keep their names and their signatures, so `XLRangeBase.cs:980` and `:1078` do
not change in task 4.

Note what the axis change reveals: `InsertColumnsBefore` returns `.Columns()` and `InsertRowsAbove`
returns `.Rows()` — the *return* differs on the index axis, but `ApplyColumnFormattingFromLeft`
(`:83-98`) styles `rangeToReturn.Row(ro)` while `ApplyRowFormattingFromAbove` (`:190-205`) styles
`rangeToReturn.Column(co)`. **The formatting loop runs on the cross axis, not the index axis.** That
transposition is invisible while the two are written out longhand and is the single most likely
place for a mistake during the collapse. It is why task 3 exists.

### `XLRow` / `XLColumn`

They stay `internal sealed class ... : XLRangeBase`, keep `IXLRow` / `IXLColumn` unchanged, and
shrink to what is genuinely per-axis:

- `RangeAddress` (`XLRow.cs:43-47`, `XLColumn.cs:37-41`) — different factories, stays
- `Children` / `TryApplyToCellStyles` — different slice walk, stays
- `IsHidden` — the row's flags bitfield stays, the column's auto-property stays
- everything computable from the axis — `CellCount`, `OutlineLevel`, `Cells(string)` — moves to a
  shared generic base method bound to the axis

This spec does **not** attempt to unify `XLRow` and `XLColumn` into one generic class. 702 + 601
lines against two 47- and 44-member interfaces is a bigger and riskier job than the five algorithm
files, and the algorithm files are where the defects came from. Recording the option and declining
it is the useful output.

## Global constraints

- Warnings are errors (`TreatWarningsAsErrors=true`); nullable is enabled; new code must be
  null-annotated.
- **Branch per spec; never commit to main.** Branch `task/26`. Commit prefixes `fix:` (tasks 1–3),
  `test:` (task 3's gate), `refactor:` (tasks 4–8), `docs:` (task 9).
- No compound shell commands (`&&`, `||`, `;`) in agent tool calls.
- **Do not use `sed -i` on tracked files.** `.gitattributes` checks out CRLF; Git Bash's `sed -i`
  rewrites the file as LF and turns a one-line change into a whole-file diff. Use the Edit/Write
  tools and verify with `git diff --numstat` — a file whose changed-line count is near its total
  line count was rewritten, not edited.
- Build: `dotnet build XLibur/XLibur.csproj -c Release -v q`
- Tests: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
- Filtering uses `--treenode-filter`, **never** `--filter`. Exit 5 = invalid option; exit 8 = zero
  tests matched. Never filter at solution level — name the `.csproj`.
- Pass `-f net10.0` while iterating; run without it before opening the PR, so net8.0 and net9.0 are
  covered.
- Tests are TUnit and **awaitable**: `await Assert.That(actual).IsEqualTo(expected)`. A missing
  `await` silently passes. Attributes are `[Test]`, `[Arguments(...)]`, `[MethodDataSource(...)]`.
  The suite is serial (`[assembly: NotInParallel]`).

## Work plan

Defects first, as three independent commits each carrying the test that would have caught it. They
are worth landing on their own, and together they are the gate the refactor needs.

| # | Task | Size | Gate |
|---|---|---|---|
| 1 | Row outline levels reach the row counter; `GetMaxRowOutline` stops throwing | S | New round-trip test green; `@outlineLevelRow` emitted |
| 2 | `XLColumn.CellCount()` returns the sheet's row count | XS | New test green; suite unchanged |
| 3 | Symmetry gate: assert the two axes agree, before anything moves | S | New test green on unmodified code, and proven to bite |
| 4 | `IGridAxis` + collapse `XLRangeInsertHelper` — the pattern-setter | M | Suite green; file ≤ 140 lines |
| 5 | Collapse `XLRangeShiftHelper` | S | Suite green; file ≤ 95 lines |
| 6 | Collapse `XLRangeBase.cs:884-1078` | M | Suite green |
| 7 | Collapse `XLWorksheet.cs:1248-1403` | M | Suite green |
| 8 | Collapse `XLWorksheetRangeShifter.cs` — highest risk, pin the ordering | M | Suite green; ordering documented |
| 9 | Confirm per-operation cost is unchanged; changelog | S | Bytes within 1% of baseline |

Tasks 4–8 are in ascending order of risk. 5 is a pure function. 8 touches conditional formats, data
validations, defined names, page breaks, sparklines and the calc engine, and is the file spec 33
also wants.

---

### Task 1 — Row outline levels reach the row counter

**Files:**
- Modify: `XLibur/Excel/Rows/XLRow.cs:424-425`
- Modify: `XLibur/Excel/XLOutlineTracker.cs:62-65`
- Create: `XLibur.Tests/Excel/Rows/OutlineRoundTripTests.cs`

**Interfaces:** none added. `XLWorksheet.IncrementRowOutline` / `DecrementRowOutline`
(`XLWorksheet.cs:1112`, `:1114`) gain their first callers.

- [ ] **Step 1: Write the failing test first**

```csharp
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Rows;

/// <summary>
/// Row outline levels were counted into the column counter (XLRow.cs:424-425, copied verbatim from
/// XLColumn.cs:342-343), so sheetFormatPr/@outlineLevelRow was never emitted and row groups inflated
/// @outlineLevelCol instead. Spec 26 task 1. Nothing asserted either attribute before this file.
/// </summary>
public class OutlineRoundTripTests
{
    private static XElement SheetFormatPr(Stream xlsx)
    {
        xlsx.Position = 0;
        using var doc = SpreadsheetDocument.Open(xlsx, isEditable: false);
        var part = doc.WorkbookPart!.WorksheetParts.Single();
        using var stream = part.GetStream();
        var xml = XDocument.Load(stream);
        var ns = xml.Root!.Name.Namespace;
        return xml.Root.Element(ns + "sheetFormatPr")!;
    }

    [Test]
    public async Task Grouping_rows_emits_outlineLevelRow_and_not_outlineLevelCol()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "x";
            ws.Rows(2, 4).Group();
            ws.Rows(3, 3).Group();     // level 2
            wb.SaveAs(ms);
        }

        var sfp = SheetFormatPr(ms);
        await Assert.That(sfp.Attribute("outlineLevelRow")?.Value).IsEqualTo("2");
        await Assert.That(sfp.Attribute("outlineLevelCol")).IsNull();
    }

    [Test]
    public async Task Grouping_columns_emits_outlineLevelCol_and_not_outlineLevelRow()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "x";
            ws.Columns(2, 4).Group();
            wb.SaveAs(ms);
        }

        var sfp = SheetFormatPr(ms);
        await Assert.That(sfp.Attribute("outlineLevelCol")?.Value).IsEqualTo("1");
        await Assert.That(sfp.Attribute("outlineLevelRow")).IsNull();
    }

    /// <summary>
    /// The load path sets XLRow.OutlineLevel (WorksheetSheetDataReader.cs:1087-1088), so before the
    /// fix, re-saving a file with row groups inflated that file's @outlineLevelCol. This is the
    /// round-trip half of the defect.
    /// </summary>
    [Test]
    public async Task Reloading_and_resaving_row_groups_does_not_inflate_outlineLevelCol()
    {
        using var first = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "x";
            ws.Rows(2, 4).Group();
            wb.SaveAs(first);
        }

        using var second = new MemoryStream();
        first.Position = 0;
        using (var wb = new XLWorkbook(first))
            wb.SaveAs(second);

        var sfp = SheetFormatPr(second);
        await Assert.That(sfp.Attribute("outlineLevelRow")?.Value).IsEqualTo("1");
        await Assert.That(sfp.Attribute("outlineLevelCol")).IsNull();
    }

    /// <summary>
    /// GetMaxRowOutline (XLOutlineTracker.cs:62-65) guards the dictionary's size rather than the
    /// filtered sequence's, so a dictionary holding only zero counts made .Max() throw on an empty
    /// sequence. Unreachable while row outlines were never counted; reachable the moment task 1's
    /// first change lands. RowTests.UngroupFromAll performs exactly this sequence but never saves.
    /// </summary>
    [Test]
    public async Task Grouping_then_ungrouping_every_row_still_saves()
    {
        using var ms = new MemoryStream();
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "x";
        ws.Rows(1, 2).Group();
        ws.Rows(1, 2).Ungroup(true);

        wb.SaveAs(ms);

        var sfp = SheetFormatPr(ms);
        await Assert.That(sfp.Attribute("outlineLevelRow")).IsNull();
    }
}
```

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/OutlineRoundTripTests/*"`

Expected: **FAIL** — three of the four fail before the fix. The column test should pass; if it does
not, the reading of `SheetFormatPr` is wrong, not the library. Fix the test first.

- [ ] **Step 2: Fix the tracker guard**

`XLibur/Excel/XLOutlineTracker.cs:62-65`, matching `GetMaxColumnOutline`'s shape at `:35-39`:

```csharp
    /// <summary>
    /// The filtered sequence is what needs the emptiness guard, not the dictionary: Decrement leaves
    /// a zero count behind, so a sheet whose rows were all ungrouped holds a non-empty dictionary of
    /// zeroes and Max() would throw on an empty sequence.
    /// </summary>
    public int GetMaxRowOutline()
    {
        var list = _rowOutlineCount.Where(kp => kp.Value > 0).ToList();
        return list.Count == 0 ? 0 : list.Max(kp => kp.Key);
    }
```

- [ ] **Step 3: Point the row at the row counter**

`XLibur/Excel/Rows/XLRow.cs:424-425`:

```csharp
            Worksheet.IncrementRowOutline(value);
            Worksheet.DecrementRowOutline(_outlineLevel);
```

- [ ] **Step 4: Confirm the dead members are alive**

Run: `grep -rn "IncrementRowOutline\|DecrementRowOutline" --include=*.cs XLibur`

Expected: `XLRow.cs` now appears alongside `XLOutlineTracker.cs` and `XLWorksheet.cs`. Before the
fix it did not.

- [ ] **Step 5: Run the targeted tests, then the suite**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/OutlineRoundTripTests/*"`
Expected: PASS, all four.

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS.

`@outlineLevelRow` now appears in output that never carried it and `@outlineLevelCol` disappears
from output that carried it wrongly. **Any test that fails here is asserting the defect**, not the
fix — check the assertion against Excel's own output before changing the library back.

- [ ] **Step 6: Commit**

```bash
git add XLibur/Excel/Rows/XLRow.cs XLibur/Excel/XLOutlineTracker.cs XLibur.Tests/Excel/Rows/OutlineRoundTripTests.cs
```

```bash
git commit -m 'fix(rows): count row outline levels into the row counter, not the column one (spec 26 task 1)'
```

---

### Task 2 — `XLColumn.CellCount()`

**Files:**
- Modify: `XLibur/Excel/Columns/XLColumn.cs:404-407`
- Modify: `XLibur/Excel/Rows/XLRow.cs:486-489`
- Modify: `XLibur.Tests/Excel/Columns/ColumnTests.cs`

- [ ] **Step 1: Express both in terms of the axis they mean**

`XLColumn.cs:404-407`:

```csharp
    /// <summary>
    /// A column holds one cell per row in the sheet. Until spec 26 this was a verbatim copy of
    /// <see cref="XLRow.CellCount"/> and measured the column span of an entire-column address —
    /// whose first and last address carry the same column — so it always returned 1.
    /// </summary>
    public int CellCount() => RowCount();
```

`XLRow.cs:486-489`:

```csharp
    /// <summary>A row holds one cell per column in the sheet.</summary>
    public int CellCount() => ColumnCount();
```

`RowCount()` and `ColumnCount()` are inherited from `XLRangeBase` (`:669-672`, `:691-694`) and
compute exactly the same expression the two methods spelled out, so `XLRow.CellCount` is unchanged
in value — it just stops repeating the arithmetic.

- [ ] **Step 2: Assert both, next to each other**

Add to `XLibur.Tests/Excel/Columns/ColumnTests.cs`:

```csharp
    /// <summary>
    /// XLColumn.CellCount() was a verbatim copy of XLRow.CellCount() and measured the wrong axis,
    /// so it returned 1 for every column. Spec 26 task 2. The row is asserted alongside it because
    /// the pair is the point: whatever CellCount means, it must mean the same thing on both axes.
    /// </summary>
    [Test]
    public async Task Cell_count_is_the_extent_of_the_axis_a_line_spans()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        await Assert.That(ws.Column(1).CellCount()).IsEqualTo(XLHelper.MaxRowNumber);
        await Assert.That(ws.Row(1).CellCount()).IsEqualTo(XLHelper.MaxColumnNumber);
    }
```

- [ ] **Step 3: Run**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/ColumnTests/*"`
Expected: PASS.

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS.

**If any existing test fails**, it depends on `CellCount()` returning 1 and the premise in "Why this
spec exists" is wrong. Record which test, revert this task, and say so in a Results section —
a disproved premise is a result. Do not weaken the new assertion to accommodate it.

- [ ] **Step 4: Confirm the public surface did not move**

Run: `git diff --numstat XLibur/PublicAPI.Shipped.txt XLibur/PublicAPI.Unshipped.txt`
Expected: no output.

- [ ] **Step 5: Commit**

```bash
git add XLibur/Excel/Columns/XLColumn.cs XLibur/Excel/Rows/XLRow.cs XLibur.Tests/Excel/Columns/ColumnTests.cs
```

```bash
git commit -m 'fix(columns): CellCount measures the rows a column spans, not its own column (spec 26 task 2)'
```

---

### Task 3 — The symmetry gate

Tasks 4–8 are behaviour-preserving, so they need a test that notices when one axis stops matching
the other. No existing test compares the two.

**Files:**
- Create: `XLibur.Tests/Excel/Ranges/GridAxisSymmetryTests.cs`

**Interfaces:**
- Produces: `GridAxisSymmetryTests`, the gate for tasks 4–8.

- [ ] **Step 1: Write it against the transpose**

The gate is a transpose: do the operation row-wise on one sheet, column-wise on a transposed sheet,
and assert the results transpose into each other. Anything the two implementations disagree about
shows up as an asymmetry.

```csharp
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Ranges;

/// <summary>
/// Row-wise and column-wise algorithms were written twice, line for line, and had already drifted
/// three times (spec 26). Spec 26 collapses them onto one axis-parameterised implementation; this is
/// the gate that says the collapse changed nothing. Every case builds the same content twice, once
/// transposed, runs the mirrored operation, and asserts the two land on transposed addresses.
/// </summary>
public class GridAxisSymmetryTests
{
    private static IXLWorksheet Populate(IXLWorksheet ws, bool transposed)
    {
        for (var i = 1; i <= 6; i++)
            for (var j = 1; j <= 4; j++)
            {
                var cell = transposed ? ws.Cell(j, i) : ws.Cell(i, j);
                cell.Value = $"{i}.{j}";
            }

        return ws;
    }

    [Test]
    [Arguments(2, 3)]
    [Arguments(1, 1)]
    [Arguments(4, 2)]
    public async Task Insert_before_moves_the_same_content_on_both_axes(int at, int count)
    {
        using var wb = new XLWorkbook();
        var rows = Populate(wb.AddWorksheet("Rows"), transposed: false);
        var cols = Populate(wb.AddWorksheet("Cols"), transposed: true);

        rows.Row(at).InsertRowsAbove(count);
        cols.Column(at).InsertColumnsBefore(count);

        for (var i = 1; i <= 6 + count; i++)
            for (var j = 1; j <= 4; j++)
                await Assert.That(cols.Cell(j, i).GetString())
                    .IsEqualTo(rows.Cell(i, j).GetString());
    }

    [Test]
    [Arguments(2, 3)]
    [Arguments(1, 2)]
    public async Task Delete_moves_the_same_content_on_both_axes(int at, int count)
    {
        using var wb = new XLWorkbook();
        var rows = Populate(wb.AddWorksheet("Rows"), transposed: false);
        var cols = Populate(wb.AddWorksheet("Cols"), transposed: true);

        rows.Rows(at, at + count - 1).Delete();
        cols.Columns(at, at + count - 1).Delete();

        for (var i = 1; i <= 6; i++)
            for (var j = 1; j <= 4; j++)
                await Assert.That(cols.Cell(j, i).GetString())
                    .IsEqualTo(rows.Cell(i, j).GetString());
    }

    /// <summary>
    /// XLRangeShiftHelper.ShiftColumns/ShiftRows repositions live ranges. A held range must survive
    /// or die identically on both axes, including the destroyed-by-shift case at :49 / :116.
    /// </summary>
    [Test]
    [Arguments(1, 3)]
    [Arguments(3, -2)]
    [Arguments(2, -6)]
    public async Task A_live_range_is_repositioned_identically_on_both_axes(int at, int shift)
    {
        using var wb = new XLWorkbook();
        var rows = Populate(wb.AddWorksheet("Rows"), transposed: false);
        var cols = Populate(wb.AddWorksheet("Cols"), transposed: true);

        var heldRows = rows.Range(2, 1, 5, 4);
        var heldCols = cols.Range(1, 2, 4, 5);

        if (shift > 0)
        {
            rows.Row(at).InsertRowsAbove(shift);
            cols.Column(at).InsertColumnsBefore(shift);
        }
        else
        {
            rows.Rows(at, at - shift - 1).Delete();
            cols.Columns(at, at - shift - 1).Delete();
        }

        await Assert.That(heldCols.RangeAddress.IsValid).IsEqualTo(heldRows.RangeAddress.IsValid);
        if (!heldRows.RangeAddress.IsValid)
            return;

        await Assert.That(heldCols.RangeAddress.FirstAddress.ColumnNumber)
            .IsEqualTo(heldRows.RangeAddress.FirstAddress.RowNumber);
        await Assert.That(heldCols.RangeAddress.LastAddress.ColumnNumber)
            .IsEqualTo(heldRows.RangeAddress.LastAddress.RowNumber);
    }

    /// <summary>
    /// XLWorksheetRangeShifter moves conditional formats, data validations, defined names and page
    /// breaks. Each was written twice; each must move the same distance on both axes.
    /// </summary>
    [Test]
    public async Task Conditional_formats_and_validations_move_identically_on_both_axes()
    {
        using var wb = new XLWorkbook();
        var rows = Populate(wb.AddWorksheet("Rows"), transposed: false);
        var cols = Populate(wb.AddWorksheet("Cols"), transposed: true);

        rows.Range("A3:D5").AddConditionalFormat().WhenNotEmpty().Fill.SetBackgroundColor(XLColor.Red);
        cols.Range("C1:E4").AddConditionalFormat().WhenNotEmpty().Fill.SetBackgroundColor(XLColor.Red);

        rows.Range("A3:D5").CreateDataValidation().WholeNumber.Between(1, 10);
        cols.Range("C1:E4").CreateDataValidation().WholeNumber.Between(1, 10);

        rows.Row(2).InsertRowsAbove(2);
        cols.Column(2).InsertColumnsBefore(2);

        var rowCf = rows.ConditionalFormats.Single().Ranges.Single().RangeAddress;
        var colCf = cols.ConditionalFormats.Single().Ranges.Single().RangeAddress;
        await Assert.That(colCf.FirstAddress.ColumnNumber).IsEqualTo(rowCf.FirstAddress.RowNumber);
        await Assert.That(colCf.LastAddress.ColumnNumber).IsEqualTo(rowCf.LastAddress.RowNumber);

        var rowDv = rows.DataValidations.Single().Ranges.Single().RangeAddress;
        var colDv = cols.DataValidations.Single().Ranges.Single().RangeAddress;
        await Assert.That(colDv.FirstAddress.ColumnNumber).IsEqualTo(rowDv.FirstAddress.RowNumber);
        await Assert.That(colDv.LastAddress.ColumnNumber).IsEqualTo(rowDv.LastAddress.RowNumber);
    }
}
```

Some of these builder calls are written from the interfaces as documented. If one does not compile,
find the equivalent form already used in `XLibur.Tests/Excel/Columns/InsertColumnBeforeDataValidationTests.cs`
or `XLibur.Tests/Excel/Ranges/` and use that. **Do not weaken an assertion to make it pass.**

- [ ] **Step 2: Run it on unmodified code**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/GridAxisSymmetryTests/*"`

Expected: PASS.

**If a case fails here, you have found a fourth drift.** Record which case, which axis is wrong, and
against what Excel does. Split it into its own `fix:` commit *before* task 4 — a pre-existing
asymmetry must not land inside the refactor's diff, or nobody can tell which change caused it. If it
cannot be settled quickly, replace the assertion with the current behaviour plus a comment naming
the gap, and report it.

- [ ] **Step 3: Prove the gate bites**

Temporarily change `XLRangeInsertHelper.ShiftRowHeights` (`:164-180`) to return immediately.
Re-run.
Expected: at least one case FAILS. Restore it.

Then temporarily swap `XLWorksheetRangeShifter.ShiftDataValidationRows`' `affected` area (`:200-202`)
to use `last.RowNumber` instead of `first.RowNumber + rowsShifted - 1`. Re-run.
Expected: `Conditional_formats_and_validations_move_identically_on_both_axes` FAILS. Restore it.

A gate that does not bite is not a gate. If either mutation passes, the test is not reaching the
code tasks 4–8 change and must be strengthened before proceeding.

- [ ] **Step 4: Commit**

```bash
git add XLibur.Tests/Excel/Ranges/GridAxisSymmetryTests.cs
```

```bash
git commit -m 'test(ranges): assert the two grid axes agree, before collapsing them (spec 26 task 3)'
```

---

### Task 4 — `IGridAxis`, and `XLRangeInsertHelper` as the pattern-setter

226 lines that are two 106-line copies. It has one caller pair, both one-line
(`XLRangeBase.cs:980`, `:1078`), so the blast radius is the smallest of the five.

**Files:**
- Create: `XLibur/Excel/Coordinates/GridAxis.cs`
- Modify: `XLibur/Excel/Ranges/XLRangeInsertHelper.cs`

**Interfaces:**
- Produces: `IGridAxis`, `RowAxis`, `ColumnAxis` (all `internal`).
- `XLRangeInsertHelper.InsertColumnsBefore` and `InsertRowsAbove` keep their exact signatures.

- [ ] **Step 1: Write `GridAxis.cs`**

Start from the interface sketched in "The design". Add a member only when the collapse in step 2
demands it, and stop when the collapse compiles — an interface grown from the call sites is the only
way to keep it honest. **The expected width is 15–20 members.** That is wide, and it is wide because
the duplication is wide: the test is whether ~18 one-expression members replace ~444 duplicated
lines, not whether the interface is small.

Every method is `readonly` on both structs and none of them holds state, so both are zero-byte
values.

- [ ] **Step 2: Collapse the six method pairs into one generic set**

```csharp
    private static IXLRange? Insert<TAxis>(XLRangeBase range, bool onlyUsedCells, int count,
        bool formatFromPrevious, bool nullReturn)
        where TAxis : struct, IGridAxis
    {
        var axis = default(TAxis);
        if (count <= 0 || count > axis.MaxIndex)
            throw new ArgumentOutOfRangeException(nameof(count),
                $"Number of lines to insert must be a positive number no more than {axis.MaxIndex}");
        // ... :19-49, once
    }
```

**Watch the transposition.** `ApplyColumnFormattingFromLeft` (`:83-98`) styles
`rangeToReturn.Row(ro)` while `ApplyRowFormattingFromAbove` (`:190-205`) styles
`rangeToReturn.Column(co)`. The formatting loop runs on the **cross** axis. So does
`ApplyColumnFormattingFromExistingRows` (`:100-118`), which reads
`Worksheet.Internals.RowsCollection`, against `ApplyRowFormattingFromExistingColumns` (`:207-225`),
which reads `ColumnsCollection`. Bind these to `axis.CrossOf` / a cross-axis line accessor, not to
the index axis. Getting this backwards compiles cleanly and fails task 3's
`Insert_before_moves_the_same_content_on_both_axes`.

Two other asymmetries to carry across deliberately, not silently:

- The row path has a comment at `:152` (`// Skip calling .Rows() for performance reasons if
  required.`) that the column path lacks. Keep it, once.
- `ShiftColumnWidths` (`:57-73`) tests `range.IsEntireColumn()` **inside** the loop; `ShiftRowHeights`
  (`:164-180`) tests `range.IsEntireRow()` inside its loop too. Both are loop-invariant. Hoisting
  the check is a behaviour-preserving improvement the collapse makes obvious; do it, and say so in
  the commit message.

- [ ] **Step 3: Verify the file shrank and the callers did not change**

Run: `git diff --numstat XLibur/Excel/Ranges/XLRangeInsertHelper.cs XLibur/Excel/Ranges/XLRangeBase.cs`

Expected: `XLRangeInsertHelper.cs` heavily changed; `XLRangeBase.cs` **0 0** — its two call sites
(`:980`, `:1078`) are untouched in this task.

Run: `wc -l XLibur/Excel/Ranges/XLRangeInsertHelper.cs`
Expected: ≤ 140.

- [ ] **Step 4: Build and run**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Expected: clean. Only Polyfill warnings are normal.

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS, including `GridAxisSymmetryTests`.

**If `static abstract` or the generic constraint trips an analyzer** — Sonar or the nullable
analyzer, both of which are errors here — do not suppress it globally. Either narrow the suppression
at the member with a reason, matching how the repo handles `S4136` in `XLCellFormulaShifter`, or
fall back to two `static` adapter classes over one generic implementation. The design's requirement
is "no interface-typed receiver on the hot path", not "this exact syntax".

- [ ] **Step 5: Commit**

```bash
git add XLibur/Excel/Coordinates/GridAxis.cs XLibur/Excel/Ranges/XLRangeInsertHelper.cs
```

```bash
git commit -m 'refactor(ranges): one insert implementation, bound to a grid axis (spec 26 task 4)'
```

---

### Task 5 — `XLRangeShiftHelper`

The lowest-risk of the four remaining: two pure functions, `:12-76` and `:78-143`, whose only side
effect is `worksheet.DeleteRange` at `:58` / `:125`.

**Files:**
- Modify: `XLibur/Excel/Ranges/XLRangeShiftHelper.cs`

**Interfaces:**
- `ShiftColumns` and `ShiftRows` keep their signatures; `XLRangeBase.cs:1173-1177` is untouched.

- [ ] **Step 1: Collapse**

The two bodies differ only in which coordinate is read and which is written. In axis terms:

- `allRowsAreCovered` (`:21-23`) / `allColumnsAreCovered` (`:87-89`) → *the shifted range spans this
  range on the cross axis*
- `shiftLeftBoundary` / `shiftTopBoundary` (`:28-31`, `:94-98`) → *the leading edge moves*
- `shiftRightBoundary` / `shiftBottomBoundary` (`:33-34`, `:100-101`) → *the trailing edge moves*
- `destroyedByShift` (`:49`, `:116`) → *the trailing edge crossed the leading one*

Name them for the axis, not for the direction — `leadingEdgeMoves`, `trailingEdgeMoves`,
`spannedOnCrossAxis` — because "left" and "top" are the same concept and the current names are the
reason nobody noticed they were the same code.

- [ ] **Step 2: Check the file shrank**

Run: `wc -l XLibur/Excel/Ranges/XLRangeShiftHelper.cs`
Expected: ≤ 95.

- [ ] **Step 3: Build and run**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS. `A_live_range_is_repositioned_identically_on_both_axes` is the case that covers this
file directly, including its `destroyedByShift` branch via `[Arguments(2, -6)]`.

- [ ] **Step 4: Commit**

```bash
git add XLibur/Excel/Ranges/XLRangeShiftHelper.cs
```

```bash
git commit -m 'refactor(ranges): one address-shift implementation for both axes (spec 26 task 5)'
```

---

### Task 6 — `XLRangeBase.cs:884-1078`

Ten methods, twice, with identical line counts: **4 / 21 / 4 / 4 / 19 / 4 / 21 / 4 / 4 / 3**.

**Files:**
- Modify: `XLibur/Excel/Ranges/XLRangeBase.cs:884-1078`

- [ ] **Step 1: Keep every public signature; collapse only the bodies**

Eight of the twenty are `public` on `IXLRangeBase` / `IXLRange` (`InsertColumnsAfter` ×3,
`InsertColumnsBefore` ×3 and their row twins, plus the `*Void` overloads). They stay exactly as they
are — one line each, delegating to a generic private. Only the two 19–21 line bodies collapse:

- the `expandRange` address rebuild (`:889-909` ↔ `:987-1007`, and `:946-966` ↔ `:1044-1064`)
- the after/below → before/above relocation (`:921-939` ↔ `:1019-1037`)

The relocation body is the one with real content: it computes the next block's first and last index,
clamps both to `axis.MaxIndex`, computes the cross-axis span, clamps that to `axis.MaxCross`, and
delegates. In axis terms it is nine lines.

- [ ] **Step 2: Confirm the public surface is untouched**

Run: `git diff --numstat XLibur/PublicAPI.Shipped.txt XLibur/PublicAPI.Unshipped.txt`
Expected: no output.

Run: `grep -c 'public IXLRange\(Columns\|Rows\) Insert' XLibur/Excel/Ranges/XLRangeBase.cs`
Expected: 8, unchanged.

- [ ] **Step 3: Build and run on both frameworks**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj`
Expected: PASS on net8.0, net9.0 and net10.0. This is the first task where the generic instantiation
count changes materially, so run all three here rather than only before the PR.

- [ ] **Step 4: Commit**

```bash
git add XLibur/Excel/Ranges/XLRangeBase.cs
```

```bash
git commit -m 'refactor(ranges): collapse the mirrored insert block onto one axis (spec 26 task 6)'
```

---

### Task 7 — `XLWorksheet.cs:1248-1403`

Two pairs: `NotifyRangeShifted*` (`:1248-1278` ↔ `:1280-1308`) and `CollectRangesShiftedBy*`
(`:1333-1367` ↔ `:1373-1403`).

**Files:**
- Modify: `XLibur/Excel/XLWorksheet.cs:1248-1403`

- [ ] **Step 1: Collapse `CollectRangesShiftedBy*` first**

It is the pair carrying the pointer comment. The predicate is three tests (`:1344-1352` ↔
`:1384-1392`): address validity, *entire-line-on-the-other-axis* exclusion, and past-the-shift-line
plus cross-axis containment. The sort (`:1361-1364` ↔ `:1397-1400`) is the same comparison on the
index axis.

The 35-line remarks block at `:1310-1332` moves onto the single implementation, and the pointer
comment at `:1369-1372` is deleted rather than kept — there is nothing left for it to point at.
Update its wording: it currently names `XLRangeShiftHelper.ShiftRows` (`:1316`), which task 5 has
renamed.

- [ ] **Step 2: Collapse `NotifyRangeShifted*`**

The two bodies differ in three tokens: `CollectRangesShiftedByRows`/`ByColumns`,
`WorksheetRangeShiftedRows`/`Columns` (worksheet-level, `:1254`/`:1286`) and
`storedRange.WorksheetRangeShiftedRows`/`Columns` (`:1267`/`:1297`). The last is a `virtual` on
`XLRangeBase` overridden in `XLRange` (`XLRange.cs:355-358` and its column twin), so it must reach
the axis through a dispatch member on `IGridAxis`, not through a delegate — a delegate here would
allocate once per shifted range and this loop is spec 05's hot path.

`RangeShiftPasses++` (`:1250`, `:1282`) happens once either way. Keep it once.

The `collapsed` logic (`:1261-1277` ↔ `:1291-1307`) is character-identical apart from those tokens.

- [ ] **Step 3: Build and run**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS.

- [ ] **Step 4: Commit**

```bash
git add XLibur/Excel/XLWorksheet.cs
```

```bash
git commit -m 'refactor(worksheet): one range-shift notification pass for both axes (spec 26 task 7)'
```

---

### Task 8 — `XLWorksheetRangeShifter`, and the ordering question

The highest-risk file: 6 mirror pairs against 3 shared methods, touching merged ranges, defined
names, conditional formats, data validations, page breaks, sparklines, the calc engine and
hyperlinks. It is also the file **spec 33** wants, which is why 26 runs first.

**Files:**
- Modify: `XLibur/Excel/XLWorksheetRangeShifter.cs`

- [ ] **Step 1: Settle the ordering drift before collapsing anything**

`ShiftColumns` (`:37-41`) and `ShiftRows` (`:80-84`) run the same five steps in different orders:

| | `ShiftColumns` | `ShiftRows` |
|---|---|---|
| 1 | `ShiftConditionalFormattingColumns` | `ShiftConditionalFormattingRows` |
| 2 | `ShiftDataValidationColumns` | `ShiftDataValidationRows` |
| 3 | `ShiftDataValidationFormulaColumns` | `ShiftDataValidationFormulaRows` |
| 4 | `ShiftPageBreaksColumns` | `RemoveInvalidSparklines` |
| 5 | `RemoveInvalidSparklines` | `ShiftPageBreaksRows` |

`RemoveInvalidSparklines` (`:273-283`) reads only `worksheet.SparklineGroups` and address validity.
`ShiftPageBreaks*` (`:103-125`) reads and writes only `worksheet.PageSetup.*Breaks`. They share no
state, so they commute — but that is an inspection, not a test.

Write the test that makes it a fact:

```csharp
    /// <summary>
    /// XLWorksheetRangeShifter ran page-break shifting and sparkline cleanup in opposite orders on
    /// the two axes (:40-41 vs :83-84), with nothing stating whether that mattered. It does not: the
    /// two touch disjoint state. Spec 26 task 8 collapses them onto one order; this pins the outcome
    /// so the collapse is provably a no-op.
    /// </summary>
    [Test]
    public async Task Page_breaks_and_sparklines_survive_a_shift_on_both_axes()
    {
        // ... a sheet carrying both a page break past the insert point and a sparkline
        //     inside the deleted region, shifted on each axis, asserted transposed.
    }
```

Add it to `GridAxisSymmetryTests`. Run it **before** the collapse.
Expected: PASS. If it fails, the order does matter, this task's premise is wrong, and the fix is a
separate `fix:` commit that must land before the collapse — record which order Excel agrees with.

- [ ] **Step 2: Collapse the six pairs**

Take them in the order they are least entangled:

1. `ShiftPageBreaks*` (`:103` ↔ `:115`) — 11 lines each, reads one index
2. `MoveDefinedNames*` (`:285` ↔ `:303`) — 17 lines each, differ only in
   `XLCellFormulaShifter.ShiftFormulaRows`/`Columns`
3. `ShiftDataValidationFormula*` (`:239` ↔ `:259`) — same single difference. Their two doc comments
   are asymmetric: the column one (`:228-238`) carries the full 11-line rationale and the row one
   (`:253-258`) is a 6-line pointer to it, the same pattern as `XLWorksheet.cs:1369-1372`. Keep the
   full text, once.
4. `ShiftConditionalFormatting*` (`:127` ↔ `:141`) and `ShiftDataValidation*` (`:176` ↔ `:192`) —
   the `affected` area construction is the axis-dependent part; the transform lambdas name
   `InsertAndShiftRight`/`Down` and `DeleteAndShiftLeft`/`Up`, which become two axis members.
   `ShiftConditionalFormats` (`:164-174`) and `ShiftDataValidations` (`:216-226`) are already shared
   and do not change.
5. `ShiftColumns` / `ShiftRows` (`:17` ↔ `:60`) last — the merged-range split (`:19-33` ↔ `:62-76`)
   plus the calc-engine and hyperlink notification (`:43-57` ↔ `:86-100`).

- [ ] **Step 3: Leave `ISheetListener` alone**

`:43` and `:86` both do `ISheetListener hyperlinks = worksheet.Hyperlinks;` and then call one of the
four mirrored members. Route the choice through `IGridAxis`; **do not change
`XLibur/Excel/Cells/ISheetListener.cs`.** That interface is spec 33's, and touching it here would
create the conflict this ordering exists to avoid.

- [ ] **Step 4: Check the file shrank**

Run: `wc -l XLibur/Excel/XLWorksheetRangeShifter.cs`
Expected: ≤ 215.

- [ ] **Step 5: Build and run on both frameworks**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj`
Expected: PASS on net8.0, net9.0 and net10.0.

- [ ] **Step 6: Commit**

```bash
git add XLibur/Excel/XLWorksheetRangeShifter.cs XLibur.Tests/Excel/Ranges/GridAxisSymmetryTests.cs
```

```bash
git commit -m 'refactor(worksheet): one range-shifter for both axes, with one step order (spec 26 task 8)'
```

---

### Task 9 — Cost, and the changelog

The insert path is spec 05's territory and this spec restructures it, so the cost must be shown not
to have moved. Spec 21's result is the reason this task has teeth: converting an enumerator to a
struct was free, but **embedding** one by value cost +60%, measured. `RowAxis` and `ColumnAxis` are
zero-byte values passed as type arguments rather than embedded, so the same mechanism should not
apply — but "should not" is what 21 disproved.

- [ ] **Step 1: Baseline the merge-base**

```
git stash
```

```
dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile structural
```

Record all seven probe rows. `StructuralEditProfile` (`XLibur.Benchmarks/StructuralEditProfile.cs:20`)
runs 1,000 single-row inserts across probes that isolate the range pass from the formula pass.

```
git stash pop
```

- [ ] **Step 2: Measure the branch**

Same command.

- [ ] **Step 3: Compare bytes, not milliseconds**

`StructuralEditProfile.cs:43` says it: *"Bytes are exact. Times are single-shot — use BenchmarkDotNet
for time claims."* This machine carries ~40% run-to-run timing variance, so a time delta from this
probe means nothing.

**Decision rule.** Allocated bytes must be within **1%** of baseline on every probe. Anything above
that is a regression that must be explained before this spec lands — the likely cause is an axis
value being captured into a closure or held in an `IGridAxis`-typed local, both of which allocate
where the generic form does not. Grep for it:

```
grep -n 'IGridAxis ' XLibur/Excel/Coordinates/GridAxis.cs XLibur/Excel/Ranges/*.cs XLibur/Excel/XLWorksheet*.cs
```

Expected: matches only in the interface declaration and the `where TAxis : struct, IGridAxis`
constraints. **An `IGridAxis`-typed field, local or parameter anywhere else is the bug.**

If bytes hold and a time claim is still wanted, run BenchmarkDotNet — but no time claim is required
by this spec's acceptance criteria.

**The profile is row-only** (`StructuralEditProfile.cs:100`, `:111` both call `InsertRowsAbove`).
That is acceptable: the collapse makes the two axes the same code, so a row measurement now covers
both. Do not add a column probe to prove that — it would be proving the refactor with the refactor.

- [ ] **Step 4: Write the changelog**

Two behaviour changes ship. Both go under `## Unreleased` in `CHANGELOG.md:15`, following the
existing entry style (a `### 🐛 Bug Fixes` heading, one bolded lead sentence per fix, then what
changes for a caller).

- `sheetFormatPr/@outlineLevelRow` now appears on sheets with grouped rows and no longer appears as
  an inflated `@outlineLevelCol`. Files XLibur previously saved carry the wrong attribute; opening
  and re-saving them corrects it.
- `IXLColumn.CellCount()` returns 1,048,576 instead of 1.

- [ ] **Step 5: Full suite, all frameworks, then commit**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj`
Expected: PASS on net8.0, net9.0 and net10.0.

```bash
git add CHANGELOG.md docs/specs/26-grid-axis.md
```

```bash
git commit -m 'docs(specs): record the structural-edit numbers and the two behaviour changes for spec 26'
```

---

## Acceptance criteria

Each is mechanically checkable at `task/26`'s tip.

1. **`IncrementRowOutline` has a caller.**
   `grep -rn 'IncrementRowOutline' --include=*.cs XLibur` names `XLibur/Excel/Rows/XLRow.cs`.
2. **No row writes the column outline counter.**
   `grep -n 'ColumnOutline' XLibur/Excel/Rows/XLRow.cs` returns nothing.
3. **`GetMaxRowOutline` guards the filtered sequence.** A grouped-then-ungrouped sheet saves without
   throwing (`OutlineRoundTripTests.Grouping_then_ungrouping_every_row_still_saves`).
4. **`@outlineLevelRow` is emitted and `@outlineLevelCol` is not inflated by row groups**, on a fresh
   save and on a reload-and-resave (`OutlineRoundTripTests`, 4 cases).
5. **`ws.Column(1).CellCount() == XLHelper.MaxRowNumber`** and
   **`ws.Row(1).CellCount() == XLHelper.MaxColumnNumber`**.
6. **`GridAxisSymmetryTests` passes**, with no assertion weakened from the form in task 3, and it was
   shown to bite under both mutations in task 3 step 3.
7. **File sizes.** `wc -l` gives: `XLRangeInsertHelper.cs` ≤ 140 (from 226);
   `XLRangeShiftHelper.cs` ≤ 95 (from 144); `XLWorksheetRangeShifter.cs` ≤ 215 (from 320).
8. **Net line reduction ≥ 250** across the five algorithm files plus `GridAxis.cs`:
   `git diff --numstat main -- XLibur/Excel/Coordinates/GridAxis.cs XLibur/Excel/Ranges/XLRangeInsertHelper.cs XLibur/Excel/Ranges/XLRangeShiftHelper.cs XLibur/Excel/Ranges/XLRangeBase.cs XLibur/Excel/XLWorksheet.cs XLibur/Excel/XLWorksheetRangeShifter.cs`
   — deletions minus additions ≥ 250.
9. **No `IGridAxis`-typed field, local or parameter.**
   `grep -rn ': IGridAxis\|IGridAxis [a-z]' --include=*.cs XLibur` matches only the two struct
   declarations and the generic constraints.
10. **No mirrored pair survives in the five files.**
    `grep -cE 'Columns?\(' ...` is too crude; use the pair script from "Why this spec exists"
    restricted to the five files. Expected: **0** same-file pairs in `XLRangeInsertHelper.cs`,
    `XLRangeShiftHelper.cs` and `XLWorksheetRangeShifter.cs`; ≤ 2 in `XLRangeBase.cs` and
    `XLWorksheet.cs` (the public `IXLRangeBase` overloads and `RowCount`/`ColumnCount`, which are
    interface obligations, not duplication).
11. **`XLibur/Excel/Cells/XLCellFormulaShifter.Legacy.cs` and
    `XLibur/Excel/Cells/ISheetListener.cs` are untouched.**
    `git diff --numstat main -- XLibur/Excel/Cells/` returns nothing.
12. **No public API change.**
    `git diff --numstat main -- XLibur/PublicAPI.Shipped.txt XLibur/PublicAPI.Unshipped.txt` returns
    nothing, and the four interface counts still read 66 / 65 / 47 / 44.
13. **Allocated bytes on every `profile structural` probe within 1% of the merge-base**, recorded in
    a Results section.
14. **Full suite green on net8.0, net9.0 and net10.0.**
15. **`CHANGELOG.md` documents both behaviour changes.**
16. **Any premise disproved along the way is recorded in a Results section**, whether or not the
    task that found it was completed. Specifically: if `CellCount()` was meant to return 1
    (task 2), if the page-break/sparkline order matters (task 8), or if the generic constraint
    costs allocations (task 9).

## Conflicts

- **Spec 33 (sheet-listener seam)** — the live conflict, written in parallel today. It shares
  `XLibur/Excel/XLWorksheetRangeShifter.cs` and `XLibur/Excel/XLWorksheet.cs`.

  **26 runs first.** It halves both files before 33 reorganises them: `XLWorksheetRangeShifter.cs`
  goes 320 → ~210 and its 6 mirror pairs become 6 single methods, so 33 rebases onto half as many
  call sites. The reverse order means 26 collapsing pairs that 33 has already moved, which is a
  worse merge and a worse review.

  26 stays out of `ISheetListener.cs` entirely (task 8 step 3, acceptance criterion 11), so the two
  specs' *interface* surfaces do not overlap — only the caller file does.

  If 33 must start first, 26's tasks 1–7 are file-disjoint from it and can run concurrently; only
  task 8 waits.

- **Spec 25 (formula shifter seam)** — **disjoint by construction.**
  `XLCellFormulaShifter.Legacy.cs` carries 9 mirror pairs, the third-largest concentration in the
  tree, and this spec deliberately leaves them. Criterion 11 enforces it. The two specs can run
  concurrently.

- **Spec 14 (`Clear`/`CopyTo` scalability)** — same file, disjoint regions. 14 modifies
  `XLRangeBase.Clear` and `CreateDataValidation` (spec 14, "File structure"); 26 modifies
  `XLRangeBase.cs:884-1078`. Textually non-overlapping, so either order merges — but 26's task 6
  moves line numbers, so whichever lands second re-reads the file. No coordination needed beyond
  that.

- **Spec 05 (structural-edit scalability)** — **done**, and adjacent. 05 owns the cost of the insert
  path this spec restructures, and its Results record where that cost is: formula shifting 68%, the
  range-shift pass 8%. This spec must not move either, which is what task 9 checks against
  `StructuralEditProfile` — the probe harness 05 built. 05 also leaves an unscoped follow-on (the
  ~665 ms fixed per-insert cost traced to `RelocateRange` probing the range repository for `XLRow`
  instances never stored there); **26 does not address it and must not appear to.**

- **Spec 21 (hot-path struct candidates)** — **done**, and load-bearing here. 21 established that
  `Point`, `Area`, `XLAddress` and `XLRangeAddress` are already structs, and its headline is a
  disproved premise: embedding a struct enumerator by value cost **+60%** on the walk even though
  converting it to a struct was free, with interface dispatch ruled out as the cost because dynamic
  PGO had already devirtualised it. **`IGridAxis` must not undo that.** It is passed as a generic
  type argument, never embedded and never held behind an interface reference (criterion 9), and
  task 9 measures rather than assumes — 21's own task 4 was empowered to revert its task 2, and
  task 9 here carries the same authority.

- **Spec 20 (style key struct size)** and **spec 23 (single style facade)** — no shared file. 23 is
  done.

- No other spec in `docs/specs/` touches `XLRangeInsertHelper.cs`, `XLRangeShiftHelper.cs`,
  `XLRow.cs`, `XLColumn.cs` or `XLOutlineTracker.cs`.

---

## Results

Landed as tasks 1–9 on `task/26` (base `c569b95a`). Suite green: 28,264 tests, both TFMs.

### What the spec predicted that turned out wrong

**The axis needs 32 members, not 15–20.** The interface was grown from the call sites as
instructed and stopped when the collapse compiled. Five files each contribute their own
axis-dependent operations; the estimate was about half the true width. This is why acceptance
criterion 8 (net reduction ≥ 250) is unreachable: the five algorithm files shrank 324 lines and
`GridAxis.cs` costs 244, for a net 80. Two full struct implementations of 32 one-expression
members cannot fit in the ~74 lines criterion 8 implies.

**`IXLAddress` on the axis was an allocation bug, and the first measurement caught it.**
`XLAddress` is a `struct` and `XLRangeAddress` exposes `FirstAddress`/`LastAddress` as that
struct, so axis methods typed `IXLAddress` boxed on every projection — once per repository entry
per insert. Task 9's first run showed +20–33% on four probes. `in XLAddress` overloads removed it.
The spec's criterion-9 grep does not catch this: it looks for an `IGridAxis`-typed receiver, and
the boxing was on the *argument*. Spec 21's mechanism, one level down from where it was expected.

**Collapsing made the insert path cheaper, not merely neutral.** Every probe now sits at or below
the merge base, four of them 38–50% below. The mechanism is mechanical: the row and column copies
each read `thisRangeAddress.FirstAddress`/`.LastAddress` through `IXLRangeAddress` **15 times per
call**, boxing each time; the single implementation hoists them into two locals and reads them
**twice**. Writing the algorithm once removed 13 boxes per shifted range that neither copy could
see. The spec framed task 9 as "show the cost did not move"; it moved down.

**The `profile structural` probe is not byte-exact.** `StructuralEditProfile` says "Bytes are
exact", and five of its seven probes are bimodal, swinging up to ±30% run-to-run on *identical*
code. Only the formula-below and batch probes are stable to <0.5%. Criterion 13's flat 1% rule
cannot be applied to the other five; modal values across four interleaved runs per side are what
this Results section reports. Anyone making a future allocation claim from this harness needs to
know that before quoting a single run.

**The `expandRange` overload count in task 6 step 2 is a miscount.** The spec expects
`grep -c 'public IXLRange\(Columns\|Rows\) Insert'` to return 8. It returns 12, at the merge base
and at the tip. The invariant (unchanged) holds; the number does not.

**Three of four outline tests fail before the fix — actually two.** Defect 1b is genuinely latent,
as the spec's own text says, so `Grouping_then_ungrouping_every_row_still_saves` passes before the
fix and after it. The column test passes throughout, as predicted.

### Premises criterion 16 names, settled

- **Was `CellCount()` ever meant to return 1?** No. The full suite passes with
  `XLColumn.CellCount()` returning 1,048,576, and no existing test asserted it — the only
  `CellCount` assertion in the suite is on `IXLRangeRow`. The premise held.
- **Does the page-break / sparkline order matter?** No. `Page_breaks_and_sparklines_survive_a_shift_on_both_axes`
  was written and **run green against the two-order code** before task 8 collapsed it, so adopting
  a single order is provably a no-op rather than an inspection. The two passes touch disjoint
  state: `RemoveInvalidSparklines` reads only `SparklineGroups` and address validity,
  `ShiftPageBreaks` only `PageSetup.*Breaks`. The pinned order is documented on `Shift<TAxis>`,
  along with the separate fact that the *first three* steps are not interchangeable.
- **Does the generic constraint cost allocations?** No. `RowAxis`/`ColumnAxis` are zero-field
  `readonly struct`s passed as type arguments; criterion 9's grep matches only the two struct
  declarations. The sort comparer in `CollectRangesShiftedBy` deliberately builds `default(TAxis)`
  inside the lambda rather than capturing the axis local, so the closure holds only `direction`
  and allocates exactly what the two longhand copies allocated. What *did* cost allocations was
  the boxed address argument, above — not the constraint.

### What was deliberately not done

- **`XLRow`/`XLColumn` were not unified.** The spec declines this and the decision stands: 702 +
  601 lines against 47- and 44-member interfaces is a bigger job than the five algorithm files,
  and the defects came from the algorithm files.
- **`XLCellFormulaShifter.Legacy.cs` (9 pairs) and `ISheetListener.cs` are untouched**, byte for
  byte. `git diff --numstat c569b95a -- XLibur/Excel/Cells/` is empty.
- **`XLWorksheetRangeShifter.cs` is 222 lines, over criterion 7's ≤215.** Its *code* is 150 lines,
  down from 248. The overage is the ordering rationale above. Deleting documentation the spec asks
  to preserve, to satisfy a line count that exists to prove the file shrank, is the wrong trade.
- **Criterion 10's ≤2 thresholds for `XLRangeBase.cs` and `XLWorksheet.cs` are not met** (7 and
  17). Every survivor is an interface obligation — `Row`/`Column`, `RowCount`/`ColumnCount`,
  `FirstRowUsed`/`FirstColumnUsed`, the outline pass-throughs — not a duplicated algorithm. The
  detector counts every Row/Column-named member, so it cannot distinguish API surface from
  duplication; the thresholds were calibrated against the wrong quantity.
- **Task 3's gate had to be strengthened to bite.** As written it passed under the spec's own
  `ShiftRowHeights` mutation. `XLRow.InsertRowsAbove` moves line sizes itself via
  `RowsCollection.ShiftRowsDown` and then calls the helper with `onlyUsedCells: true`, which is the
  first thing `ShiftRowHeights`/`ShiftColumnWidths` bail on — so those methods are unreachable from
  the `IXLRow`/`IXLColumn` entry point every other case used. A `Line_sizes_are_carried_along_identically_on_both_axes`
  case that inserts through an entire-line *range* was added; both mutations then failed and were
  reverted.
- **A reference workbook was regenerated.** `XLibur.Tests/Resource/Examples/Misc/Outline.xlsx`
  asserted defect 1 — it omits `@outlineLevelRow` although `Examples/Misc/Outline.cs` groups rows
  to level 2. Only the `sheetFormatPr` attribute differs.

### What spec 33 inherits

`XLWorksheetRangeShifter.cs` is 320 → 222 lines and its six mirror pairs are six single generic
methods, so 33 rebases onto half as many call sites. `ISheetListener.cs` is unchanged; only the
choice between its four members is routed through `IGridAxis.OnInsertAreaAndShift` /
`OnDeleteAreaAndShift`, called from `Shift<TAxis>`. `XLWorksheet.cs` is 47 lines shorter, with
`NotifyRangeShifted*` and `CollectRangesShiftedBy*` collapsed to one generic each. Anything 33 adds
to the shifter should take the axis as a type argument and must not accept `IXLAddress` where the
caller holds the struct — see the boxing finding above.
