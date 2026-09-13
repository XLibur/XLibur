# Spec 55 — A sheet is deleted or renamed through one door, and every holder hears about it

**Area:** Architecture · **Defects (4 executed, ~9 read)** · **Behaviour change (`fix!:`)**
**Effort:** L (~6–8 days, not counting the Excel fixtures)
**Dependencies:** **Hard: spec 54** (rename and delete rewrite through `FormulaText`) and **parser 4.0.0**,
which is this spec's task 0. **Excel-authored fixtures from the owner** gate tasks 3, 5 and 6. Sequenced
**after spec 44** and **before spec 48** — see *Conflicts*.
**Status:** Proposed. From the 2026-09-13 architecture review (round 4). Every design decision below
was taken by the owner in a design interview (see *Decisions*). Also recorded:
`docs/adr/0002-refused-formula-never-rewritten.md`, and the terms *defined name*, *scope*,
*3D reference* and *unsupported sheet* in `CONTEXT.md`.

Line numbers are from `d68dffdc` (2026-09-13) unless marked as the parser fork. Verify them before
editing.

---

## Goal

A sheet is deleted or renamed through one door: the worksheet collection. Every holder of text that
names a sheet — at every scope — hears about it through one listener seam, and changes the way Excel
changes it.

## Why this spec exists

### Two doors, and only one of them does the work

`ws.Delete()` (`XLWorksheet.cs:613-619`) does three things by hand, then calls the collection:

```csharp
IsDeleted = true;
Workbook.DefinedNamesInternal.OnWorksheetDeleted(Name);   // workbook-scoped names only
Workbook.NotifyWorksheetDeleting(this);                    // the calc engine only
Workbook.WorksheetsInternal.Delete(Name);
```

`wb.Worksheets.Delete(name)` (`XLWorksheets.cs:198-225`) is public, and it is the door the
documentation recommends (`docs-website/docs/worksheets.md:113`) and the examples and three tests use.
It removes the sheet, renumbers positions and calls `Cleanup()`. `Cleanup()`
(`XLWorksheet.cs:1157-1163`) only disposes internals. **It notifies nothing.**

Executed at `d68dffdc`, with `Sheet1!A1 = 5` and `Sheet2!A1 = Sheet1!A1*2` read once:

| | `wb.Worksheets.Delete("Sheet1")` | `ws.Delete()` |
|---|---|---|
| deleted sheet's `IsDeleted` | False | True |
| `Sheet2!A1.Value` | **10 — stale** | `#REF!` |
| after save and reload | **10, written to the file as the cached value** | `#REF!` |
| `Sheet2!A1` formula text | `Sheet1!A1*2` | `Sheet1!A1*2` |

The formula text is unchanged either way. So if a sheet named `Sheet1` is added later, the formula
silently binds to it.

### The seam covers one event of two

`IWorkbookListener` (`Cells/IWorkbookListener.cs`) has one member, `OnSheetRenamed`. Its registry
(`XLWorksheets.cs:256-277`) reaches the calc engine, every sheet's cells collection, and every defined
name. Delete has no member, so deletion is the hand-placed calls above, on the door callers do not
have to use.

Spec 33 gave the structural-edit seam, `ISheetListener`, 11 adapters. The workbook seam has 3.
Several holders are reached by the first and not the second. Data-validation criteria are the clearest
case: a row insert rewrites them on every sheet (`XLDataValidations.cs:184-196`), and a rename ignores
them.

### Every holder, against every event

From a read-only inventory at `d68dffdc`. **R** = read in code, **I** = inferred. Task 1 executes every
row that is not already executed.

| Holder | Stored as | Rename | `ws.Delete()` | Collection `Delete` |
|---|---|---|---|---|
| Cell formulas | A1 text | handled (`XLCellsCollection.cs:692-724`) | **text unchanged**; evaluates `#REF!` only through `IsDeleted` (R) | **text unchanged, value stale** (executed) |
| Calc dependency tree | keyed by name | handled (`XLCalcEngine.cs:783`) | purged (`XLWorkbook.cs:1370`) | **not purged**; nothing marked dirty (R) |
| Workbook-scoped names | text | handled | handled → `#REF!` (`XLDefinedName.cs:276-305`) | **not handled** (R) |
| Sheet-scoped names on other sheets | text | handled | **not handled** — `:616` walks workbook scope only (executed) | not handled (R) |
| Validation criteria (`MinValue`/`MaxValue`, list source) | text | **not handled**; saved as a dangling cross-sheet reference via `UsesExternalSheet` (`DataValidationWriter.cs:64-82`) (R) | not handled (R) | not handled (R) |
| Conditional-format values (incl. `cfvo`) | text | **not handled** (R) | not handled (R) | not handled (R) |
| Chart series references | text; loaded references are written back only if reassigned (`ChartSeriesFormatXml.cs:131-146`) | **not handled** (R) | dangle (I) | dangle (I) |
| Pivot cache source | sheet name text (`XLPivotSourceReference.cs:38-44`) | **not handled** — `TryGetSource` fails (`:100`), writer emits the stale `sheet=` (`PivotTableCacheDefinitionPartWriter.cs:143`) (R) | cache untouched; save deletes only cache parts naming the sheet (`XLWorkbook_Save.cs:123-136`) (R/I) | same (R/I) |
| Internal hyperlinks | text; sheet-qualified text is frozen (`XLHyperlink_public.cs:78-104`) | **not handled** (R) | not handled (R) | not handled (R) |
| Print area written as a formula | raw text (`XLPrintAreas.cs:17`) | **not handled** (R) | n/a (own sheet) | n/a |
| Sparklines | live ranges, turned into text at save | fine (R) | written as `#REF` via `IsDeleted` (R) | **written with the deleted sheet's name** (R) |
| Unsupported sheets | name, position | **name clash not checked** (`XLWorksheets.cs:243`) (R) | positions shifted (`:222`) | same |

Two further sheet-list defects (R):

- **The `Position` setter** (`XLWorksheet.cs:278-303`) does not shift unsupported sheets, although
  `Add` (`XLWorksheets.cs:157`) and `Delete` (`:222`) do. Positions can end up duplicated.
- **`Add` (`:169`) and `Rename` (`:243`) check only modelled sheets.** A sheet can take a chartsheet's
  name: D34's shape, which the loader guards against at `XLWorkbook_Load.cs:329-330`.

### The four defects, executed

- **D53** — `wb.Worksheets.Delete` skips `IsDeleted`, the name fix-up and the calc-engine purge.
  Dependents keep, and save, the deleted sheet's values. (The table above.)
- **D54** — a sheet-scoped name keeps pointing at a deleted sheet.
  `ws2.DefinedNames.Add("N", "Sheet1!$A$1")`; delete Sheet1: `N` still reads `Sheet1!$A$1`, through a
  save and reload, written as `<definedName name="N" localSheetId="0">Sheet1!$A$1</definedName>`. The
  workbook-scoped control `W` correctly becomes `#REF!`. `OpenXmlValidator` reports no errors — the
  dangling name is schema-valid. The code's own comment (`XLDefinedName.cs:285-288`) says Excel treats
  it as a broken file; task 3's fixture confirms or corrects that.
- **D55** — rename skips sheet-qualified names and 3D references in defined names. After renaming
  Sheet1 to Data: `Q = Sheet1!Local` is unchanged, `D3 = SUM(Sheet1:Sheet3!$A$1)` is unchanged, and the
  control `P = Sheet1!$A$1` correctly becomes `Data!$A$1`. The cause is the gate at
  `XLDefinedName.cs:309`. Called directly with the gate bypassed, the rewriter renames all of these
  forms correctly.
- **D56** — on delete, the rewriter produces invalid or non-Excel text. With the gate bypassed
  (executed against `RenameRefModVisitor` directly):

  | Input | Delete Sheet1 gives | Should give |
  |---|---|---|
  | `Sheet1!#REF!` | **`#REF!#REF!`** | `#REF!` |
  | `SUM(Sheet1:Sheet3!$A$1)` | `SUM(#REF!)` | `SUM(Sheet2:Sheet3!$A$1)` — Excel narrows a 3D reference when an endpoint sheet is deleted |

  This is latent today: the gate stops the rewriter from ever seeing these forms, and
  `DropSheetPrefixOfRefError` (`XLDefinedName.cs:295-305`) patches the first one in text. Removing the
  gate, as this spec does, would expose it. The cause is in the parser fork: `FormulaRewriter` prefixes
  a deleted sheet's error with `SheetPrefix.Deleted` (fork `FormulaRewriter.cs:139-143`), and
  `Reference3D` (`:207-213`) asks `ModifySheet` about each endpoint separately, so it cannot narrow.

## Decisions (owner, 2026-09-13)

| # | Decision |
|---|---|
| Q8 | The collection's `Delete` is the door. `ws.Delete()` becomes a one-line delegate. |
| Q9 | A reference to a deleted sheet becomes `#REF!` in names at every scope, as in Excel. Names scoped to the deleted sheet go with it. |
| Q10 | Remove the `ContainsSheet` gate and let the rewriter decide. `DropSheetPrefixOfRefError` goes once parser 4.0.0 covers `Sheet!#REF!`. |
| Q11 | The door reaches every holder of sheet-qualified text. For holders owned by an open spec (44, 48/49), this spec adds only the listener registration. |
| Q12 | A listener must not throw. That is part of the interface, and a test pins it for each adapter. No two-phase prepare/commit. |
| Q27 | A cell formula that points at a deleted sheet is rewritten to `#REF!`, as in Excel. This stops the silent rebinding. `fix!:`. |
| Q28 | For every other holder, spec 33's rule: find out what Excel does from an Excel-authored fixture, record it, implement it. Assume nothing. |
| Q29 | Sheet copy is out of scope — a follow-on candidate. |
| Q30 | `IWorkbookListener` gains one method, raised **before** the sheet is removed. No event struct. |
| Q31 | The `Position` setter and the unsupported-sheet name check are fixed here. |
| Q33 | The `Sheet!#REF!` fix goes in the parser fork. |
| Q34 | Deleting an endpoint sheet narrows a 3D reference, as in Excel. The narrowing lives in XLibur's rename visitor; the fork supplies the hook. |
| Q35 | The fork change is this spec's task 0: one fork release, 4.0.0 (planned as 3.2.0; design §1 says why it changed). |
| Q37 | The owner makes the Excel fixtures from this spec's recipes before the task that needs them is dispatched. A task whose fixture is missing stays blocked. On 2026-09-14 the owner replaced `chartsheet-name.xlsx` with an existing Excel-authored file (design §5), so three fixture pairs remain to make. |

## Non-goals

- **Sheet copy.** It creates holders rather than notifying existing ones. Follow-on candidate; the
  inventory's copy findings are recorded in `TASKLIST-architecture-deepening-4.md` §8.
- **Structural edits.** Spec 33.
- **Print areas and print titles as defined names**, and the load-side `localSheetId` mismatch
  (`DefinedNameReader.cs:81,113`). Round-4 candidate 04, not chosen; see the tasklist backlog. This
  spec renames the print-area formula text only as one more holder.
- **Text inside strings.** `INDIRECT("Sheet1!A1")` is never rewritten, and a test pins that.
- **External workbook references.** The rewriter already leaves them alone.
- **Refused formulas.** Unchanged by rename or delete (ADR 0002), which spec 54 guarantees.

## File structure

| File | Change |
|---|---|
| **Parser fork** `src/ClosedXML.Parser/FormulaRewriter.cs`, `FormulaModifier.cs`, tests, `CHANGELOG.md` | Task 0 |
| `XLibur/XLibur.csproj`, `CLAUDE.md` (*Key Dependencies*) | `XLibur.ClosedXML.Parser` 3.1.0 → 4.0.0 |
| `XLibur/Excel/Cells/IWorkbookListener.cs` | Gains `OnSheetDeleting` |
| `XLibur/Excel/XLWorksheets.cs` | `Delete` is the door; `Rename` keeps the name and the key together; the registry grows; unsupported-sheet checks |
| `XLibur/Excel/XLWorksheet.cs` | `Delete()` delegates; the `Name` setter delegates entirely; the `Position` setter shifts unsupported sheets |
| `XLibur/Excel/XLWorkbook.cs` | `NotifyWorksheetDeleting` goes; the calc engine is an adapter |
| `XLibur/Excel/CalcEngine/XLCalcEngine.cs` | `OnDeletingSheet` becomes the adapter method |
| `XLibur/Excel/Cells/XLCellsCollection.cs`, `XLCellFormula.cs` | Delete adapter — rewrite to `#REF!` |
| `XLibur/Excel/DefinedNames/XLDefinedName.cs`, `XLDefinedNames.cs` | Gate removed; delete at every scope; `DropSheetPrefixOfRefError` deleted |
| `XLibur/Excel/CalcEngine/Visitors/RenameRefModVisitor.cs` | 3D narrowing by tab order |
| `XLibur/Excel/DataValidation/*`, `ConditionalFormats/*`, `Charts/*`, `PivotTables/XLPivotSourceReference.cs`, `Hyperlinks/*`, `PageSetup/XLPrintAreas.cs` | One adapter each, behaviour from the fixtures |
| `XLibur.Tests/Resource/Other/SheetLifecycle/*.xlsx` | **New.** Excel-authored fixtures |
| `XLibur.Tests/Excel/Worksheets/SheetLifecycleTests.cs` | **New.** |

## The design

### 1. Parser 4.0.0 (task 0, in the fork)

The fork is at `D:\Data\_CodeOS\ClosedXML.Parser` (remote `XLibur/ClosedXML.Parser`). Two changes, one
release:

1. **A deleted sheet's `#REF!` loses its prefix.** When `ModifySheet` returns `null` for the sheet in
   front of an error, the rewriter emits a bare `#REF!` (`FormulaRewriter.cs:139-143`), not
   `SheetPrefix.Deleted` followed by the error.
2. **A hook that sees both endpoints of a 3D reference.** A new `protected virtual` on
   `FormulaModifier`, `ModifySheetRange(ModContext ctx, string firstSheet, string lastSheet)`, receives
   the first and last sheet together and returns the new `SheetRange?`, or `null` for `#REF!`. Its
   default calls `ModifySheet` on each endpoint, so existing modifiers behave exactly as today.
   `Reference3D` (`:207-213`) calls it.

A new `protected virtual` is an additive API change, which alone would have made the release 3.2.0.
Both changes shipped as **4.0.0** on 2026-09-13 (fork PR #54; fork issue #52 is closed), because the
same release dropped netstandard2.0 and netstandard2.1 (fork #56). The release also carries fork #50,
which refuses the formulas that ran the stack out, and fork #55, which holds a `RowCol` to a row and a
column a sheet has. The XLibur bump shows whether either changes an XLibur result. The fork's
`CHANGELOG.md` and, where a term changes, its `CONTEXT.md` are updated. XLibur then bumps its
package reference and the *Key Dependencies* line in `CLAUDE.md`.

### 2. The door

`XLWorksheets.Delete(int position)` is the only implementation of deleting a sheet. In order:

1. Raise `OnSheetDeleting(sheetName)` on every listener. The sheet can still be resolved.
2. Set `IsDeleted`.
3. Remove the sheet; renumber modelled and unsupported sheets.
4. `Cleanup()`.

`XLWorksheet.Delete()` becomes `Workbook.WorksheetsInternal.Delete(Name)`.
`XLWorkbook.NotifyWorksheetDeleting` and the name fix-up at `XLWorksheet.cs:616` are deleted, because
the calc engine and the names are now adapters.

**Rename keeps the name and the key together.** Today the `Name` setter calls `Rename`, and assigns
`_name` only if `Rename` returns. D49 showed what happens when it does not. With spec 54 in place and
Q12's invariant, a listener cannot throw. Even so, the dictionary key and the sheet's `Name` change
together in `XLWorksheets.Rename`, and the setter delegates to it entirely.

### 3. The port

```csharp
internal interface IWorkbookListener
{
    void OnSheetRenamed(string oldSheetName, string newSheetName);

    /// Raised before the sheet is removed, while it can still be resolved.
    /// Must not throw: a failure here would leave the workbook part-way through the delete.
    void OnSheetDeleting(string sheetName);
}
```

The "must not throw" invariant is part of the interface. Task 2 adds one test per adapter that asserts
it for the adapter's worst input — a refused formula, an empty collection, an already-`#REF!` value.

### 4. The registry

`XLWorksheets.GetWorkbookListeners()` yields, in an order that task 2 pins with a test:

| Adapter | Rename | Delete |
|---|---|---|
| Calc engine | exists | the existing `OnDeletingSheet`, moved behind the port |
| Each sheet's cells collection | exists | **rewrite references to `#REF!`** (Q27) |
| Defined names, every scope | exists; **gate removed** | **every scope**, not workbook scope only |
| Each sheet's data validations | **new** — per fixture | **new** — per fixture |
| Each sheet's conditional formats | **new** — per fixture | **new** — per fixture |
| Each sheet's charts | **new** — per fixture; mark references as assigned so the patcher writes them | **new** — per fixture |
| Pivot caches | **new** — per fixture | **new** — per fixture |
| Each sheet's hyperlinks | **new** — per fixture | **new** — per fixture |
| Each sheet's page setup (print-area formula text) | **new** — per fixture | **new** — per fixture |

Sparklines hold live ranges and are fixed by the door setting `IsDeleted`. Task 1 confirms this; if it
does not hold, they become an adapter too.

**Order.** Start from rename's existing order: calc engine first. If delete needs a different order,
for example so the engine marks formulas dirty after the cells collection has rewritten their text,
record why next to the test that pins it. Spec 33 did the same for its registry.

### 5. The Excel rule, and the fixtures

For every holder marked "per fixture", the behaviour comes from a file Excel wrote. Nothing is
guessed. The pivot cache source and the chart references are where a guess is most likely to be
wrong.

**Recipes.** Make each in Excel desktop, save as `.xlsx`, and put it in
`XLibur.Tests/Resource/Other/SheetLifecycle/`. Each "before" file is saved *before* the edit, and each
"after" file is the same workbook saved *after* the edit in Excel. The test loads "before", makes the
same edit in XLibur, saves, and compares each holder's text with "after".

| File | Build | Edit in Excel |
|---|---|---|
| `rename-before.xlsx` / `rename-after.xlsx` | Sheets `Data`, `Other`. `Data!A1:A3` = 1, 2, 3; `Data!B1:B3` = x, y, z; `Data!C1` = "S". On `Other`: `A1` = `=Data!A1*2`. `B1`: data validation, List, source `=Data!$A$1:$A$3`. `C1`: conditional format, "Use a formula", `=Data!$A$1>0`. `C2:C4`: a colour scale whose minimum is type *Formula*, `=Data!$A$1`. A column chart with one series: values `=Data!$A$1:$A$3`, categories `=Data!$B$1:$B$3`, name `=Data!$C$1`. A pivot table from source `Data!$A$1:$B$3`, after adding headers in row 1 and shifting the data down (so the source is `Data!$A$1:$B$4`). `D1`: a hyperlink, *Place in this document*, `Data!A1`. Names: `W` (workbook) `=Data!$A$1`; `L` (scope `Other`) `=Data!$A$1`; `Local` (scope `Data`) `=Data!$B$1`; `Q` (workbook) `=Data!Local`. On `Data`, a print area `=OFFSET(Data!$A$1,0,0,3,2)` (Name Manager → `Print_Area`, scope `Data`) | Rename `Data` to `Renamed` |
| `delete-before.xlsx` / `delete-after.xlsx` | As `rename-before`, plus sheets `First` and `Last` placed so the tab order is `First`, `Data`, `Last`, `Other`, each with a number in `A1`. Names: `T1` (workbook) `=SUM(First:Last!$A$1)`; `T2` `=SUM(Data:Last!$A$1)` | Delete `First`, then delete `Data` |
| `refdelete-before.xlsx` / `refdelete-after.xlsx` | Sheets `Data`, `Other`. Name `R` `=Data!$A$5`. Delete row 5 on `Data`, so `R` reads `=Data!#REF!`. Save this as "before" | Delete `Data` |
| ~~`chartsheet-name.xlsx`~~ **not needed** | Replaced by owner decision on 2026-09-14 with the existing Excel-authored `XLibur.Tests/Resource/Other/PivotTableReferenceFiles/ChartsheetAndPivotTable.xlsx`: worksheets `Data` and `Pivot`, chartsheet `Chart` | none — used to assert that `Add("Chart")` and renaming `Data` to `Chart` are refused |

If Excel refuses an edit, or asks a question (for example, deleting a sheet that holds the pivot
table's source), record Excel's prompt and the choice made in *Results*.

### 6. Cell formulas on delete

The cells-collection adapter rewrites every reference to the deleted sheet to `#REF!`, through
`FormulaText` and parser 4.0.0. A refused formula keeps its text (ADR 0002). This changes formula text
a caller can read, so its changelog entry is `fix!:`.

### 7. The sheet list

- The `Position` setter shifts unsupported sheets, as `Add` and `Delete` do.
- `Add` and `Rename` refuse a name held by an unsupported sheet, compared the way sheet names are
  compared everywhere else (`XLHelper.SheetComparer`).

## Global constraints

- **No public API change.** `IWorkbookListener` is internal. The behaviour of `IXLWorksheets.Delete`
  and `IXLWorksheet.Delete` changes; their signatures do not.
- **ADR 0002 holds.** A refused formula is never rewritten.
- **Listeners do not throw.** No `try`/`catch` inside the door to compensate. If an adapter can throw,
  the adapter is wrong.
- **Run all four test projects.** Spec 33's double shift reached CI because only `XLibur.Tests` was
  run, and `XLibur.Report` rewrites sheet references too.
- The ground rules in `TASKLIST-architecture-deepening-4.md` §6 apply.

## Work plan

One owner, one branch: `fix/55-sheet-lifecycle`, plus the fork branch for task 0.

### Task 0 — Parser 4.0.0

In the fork: both changes from design §1, each with its own tests. `FormulaModifierTests` must cover
`Sheet!#REF!` under delete, and the new hook with the default and an override. Release 4.0.0. Then, in
XLibur, bump the package on its own commit.

**The fork half is done:** 4.0.0 was released on 2026-09-13 (fork PR #54). What remains is the XLibur
bump.

**Gate:** fork suite green. XLibur suite green on the bump alone, `FormulaShifterCorpus.tsv` unchanged.

### Task 1 — Pin today's behaviour; execute the read findings

A holder × event characterization test — rename, `ws.Delete()`, collection `Delete` — asserting
**today's** behaviour, with the wrong answers named in the commit message. D53, D54 and D55 land as
separate tests asserting the **correct** behaviour; they fail. Execute every **R**/**I** row of the
inventory table and record in *Results* which are confirmed. Confirmed ones are added to `DEFECTS.md`.

### Task 2 — The door and the port

`OnSheetDeleting`; the collection door; `ws.Delete()` and the `Name` setter delegate; calc engine,
cells collection and names become adapters of both events; the order is pinned; one "must not throw"
test per adapter. **D53 turns green.**

### Task 3 — Names at every scope *(needs `refdelete` and `delete` fixtures)*

Remove the gate. Delete reaches every scope. Delete `DropSheetPrefixOfRefError`. Add 3D narrowing in
`RenameRefModVisitor` through the new hook, using tab order. **D54, D55 and D56 turn green.**

### Task 4 — Cell formulas

Delete rewrites to `#REF!`. Add a test for the delete-then-add rebinding: delete `Sheet1`, add a new
`Sheet1`, and assert that the old formula does not bind to it.

### Task 5 — Validation, conditional formats, hyperlinks, print-area text *(needs `rename` and `delete` fixtures)*

One adapter each, following the fixture. For validation and conditional formats, register the listener
and rewrite the text; do not restructure the owning module (specs 44 and 48/49).

### Task 6 — Charts and the pivot cache source *(needs `rename` and `delete` fixtures)*

Kept separate because the fixture is most likely to surprise here, and because a loaded chart is
patched in place (spec 10): renaming a reference must mark it as assigned, or the patcher leaves the
old text in the part.

### Task 7 — The sheet list

The `Position` setter and the unsupported-sheet name check, with
`PivotTableReferenceFiles/ChartsheetAndPivotTable.xlsx` (design §5).

### Task 8 — Cost, recorded

Time `Name =` and `Delete()` on the template round-trip workbook (`TemplateRoundTripBenchmarks`'
input), three runs, medians, before and after. This is a correctness change, so there is no revert
authority. Record the numbers in *Results*.

### Task 9 — Changelog

`fix!:` entries for the two behaviour changes a caller can see: the collection `Delete` now does what
`ws.Delete()` did, and a reference to a deleted sheet reads `#REF!` in formulas and in names at every
scope. `fix:` for the rest.

## Acceptance criteria

| # | Criterion |
|---|---|
| 1 | `IWorkbookListener` has two members, and `XLWorksheet.Delete()` is one line |
| 2 | `grep -n "OnWorksheetDeleted\|NotifyWorksheetDeleting\|DropSheetPrefixOfRefError" -r XLibur/` returns nothing |
| 3 | D53–D56 are green, in XLibur and, for D56, in the fork |
| 4 | Every holder in the inventory table either has an adapter whose behaviour matches an Excel fixture, or has a recorded reason for needing none |
| 5 | Deleting a sheet and adding one with the same name does not rebind any formula or name |
| 6 | One "must not throw" test per adapter |
| 7 | `XLibur.ClosedXML.Parser` is at 4.0.0 |
| 8 | All four test projects are green on net8.0 and net10.0 |

## Conflicts

| Spec | Shared ground | Resolution |
|---|---|---|
| **54** | `FormulaText`, `XLCellFormula.cs`, `XLDefinedName.cs` | **Hard. 54 first.** |
| **44** | `XLDataValidations.cs`, `DataValidationWriter.cs` | **44 first.** 55 then adds one listener to the reorganised module |
| **48, 49** | `XLConditionalFormat*.cs` | **55 before 48.** 55 adds only a listener; 48 and 49 then work inside the module |
| **41** | `PivotTableCacheDefinitionPartWriter.cs` | Soft. 41 owns cache items; 55 changes only the emitted source sheet. Either order |
| **31** | worksheet part writers | None, unless task 5 finds the validation writer needs changing — then sequence after 31 |
| **33** | `XLWorksheet.GetSheetListeners` | Done. 55 touches the workbook registry, not the sheet registry |
| **14** | `CopyTo` | None. Copy is out of scope |
