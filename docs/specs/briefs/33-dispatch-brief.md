# Dispatch brief — spec 33, sheet listener seam

**Dispatched:** 2026-08-26 · **Worktree:** `D:\Data\_CodeOS\xl-wt-33` · **Branch:** `refactor/33-sheet-listener-seam`,
created off `upstream/main` at `806d69f7`. Never commit to `main`.

**Spec:** `C:\data\ai\Xlibur\specs\33-sheet-listener-seam.md`. That folder is the source of truth for
specs; the repo's `docs/specs` is a copy on its way out. Read the spec in full and work its seven
tasks in order.

## What this spec does

`ISheetListener` is a declared seam that only two types use, reached by name from four hardcoded
lines in `XLWorksheetRangeShifter`. Seventeen sheet-scoped features must survive a structural edit
and they do it four different ways — and four of them (chart anchors, note callout anchors,
freeze/split panes, pivot table `Area`) do not react at all. This spec makes registering a listener
the one thing that makes a feature edit-aware, and fixes those four as a consequence.

## Ordering constraints

- **Task 1 gates everything.** Four of its tests assert the current **wrong** behaviour on purpose;
  tasks 4 and 5 must re-point *and rename* them. A test still named `does_not_move_yet` while
  asserting that it moves is worse than no test.
- **Task 3 step 1 settles the `SheetEdit` premise before anything moves.** If the arithmetic is
  disproved, narrow the port back to `(sheet, area)`, record the finding, and revise task 2.
  A disproved premise is a better result than a wider port — do not work around it.
- **Task 3 step 2: do not substitute `edit.Area` for an existing `affected` computation** because it
  looks equivalent. Check the arithmetic; if they differ, keep the existing computation in the
  adapter and say so in a comment. A silent substitution here is exactly the class of defect this
  spec family keeps finding.
- **Tasks 4 and 5 change output for existing documents.** Each needs its own changelog entry under
  `### Fixed` and a commit body naming every assertion deliberately reversed.
- **Task 7 is the perf gate.** ~40% run-to-run variance on this machine: three runs, compare
  medians, a single run proves nothing. Do **not** pre-emptively cache the listener list — measure
  first, and only cache if an allocation regression actually shows up.

## The Excel-behaviour questions — you cannot open Excel

Task 4 step 1, task 5 step 1 and task 6 step 2 each ask what Excel does. **Do not guess.** Answer
from ECMA-376 / OOXML anchor semantics and from what XLibur already models (`XLDrawingAnchor`,
`MoveAndSizeWithCells` at `XLChart.cs:87`), **state the source for each answer**, and flag any you
could not settle — ask the orchestrator rather than inventing one. An adapter that guesses is worse
than no adapter, because it produces confidently wrong output. Every answer goes in the spec's
`## Results` section (acceptance criterion 10).

## Inherited decisions — do not re-open

- **Spec 26 (PR #409) has landed** and collapsed the shifter's six mirror pairs behind
  `IGridAxis`/`Axis`. Anything you add to the shifter takes the axis as a generic type argument, and
  **must not accept `IXLAddress` where the caller holds the concrete `XLAddress` struct** — that
  boxing cost 20–33% allocation on four probes before 26 caught it.
- **Spec 26 task 8 already reconciled the page-break/sparkline ordering discrepancy.** The order test
  records that outcome; it does not decide it.
- The range-repository **spatial index was declined on measurement** in spec 05. Out of scope.
- `XLCellFormulaShifter*.cs` belongs to **spec 25**. This spec calls it and does not modify it.
- Sparkline `ShiftRows`/`ShiftColumns` (called from `XLRangeInsertHelper` / `XLRangeBase`) is a
  **different dispatch point upstream of the shifter** and is out of scope — note it in the adapter's
  remarks and leave it for a follow-on.
- **No public API change.** `PublicAPI.Unshipped.txt` untouched.

**Line numbers in the spec are from `1b41cadd`**, and specs 26 and 28 have moved code since. Verify
every one against the current tree before editing.

**Exclusions:** nothing else is running against this repo right now, so no file is owned by another
agent. If that changes the orchestrator will tell you.

## Ground rules this repo will punish you for ignoring

Read `D:\Data\_CodeOS\xl-wt-33\CLAUDE.md` before your first edit. In particular:

- **No compound shell commands** (`&&`, `||`, `;`) in tool calls. Never `cd <folder> && git ...` —
  use `git -C <path>`.
- **Never read-modify-write a whole tracked file.** `sed -i`, `python`, `perl -i`, `Set-Content` and
  shell redirection over the original are all out — `.gitattributes` checks out CRLF and they write
  LF, turning a one-line change into a whole-file diff. Use the Edit/Write tools and verify with
  `git diff --numstat`.
- The runner is **TUnit on Microsoft.Testing.Platform**, not VSTest: filtering is
  `--treenode-filter "/*/*/ClassName/*"`, **never** `--filter`. Exit 5 = invalid option, exit 8 =
  zero tests matched. Never filter at solution level — name the `.csproj`.
- **Assertions are awaitable.** A missing `await` on `Assert.That(...)` means the assertion never
  runs and the test passes regardless.
- `TreatWarningsAsErrors=true`, nullable enabled. New code must be null-annotated.
- The test project targets **net8.0 and net10.0 only**. "Green on net9.0" is unsatisfiable, not a
  missed run.
- **Diff against the merge base, never the moving upstream tip:**
  `git -C /d/Data/_CodeOS/xl-wt-33 diff --numstat $(git -C /d/Data/_CodeOS/xl-wt-33 merge-base upstream/main HEAD)`

Gate after every task: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Run unfiltered (both TFMs) before opening the PR.

## Documentation — the work is not done until this is done

- Record any defect you find that is **not the work in front of you** in
  `C:\data\ai\Xlibur\DEFECTS.md` with its call sequence and consequence, then carry on. Do not let it
  widen the task. A defect that only lives in your final report dies with the worktree.
- Tick the task boxes with the **commit sha** in
  `C:\data\ai\Xlibur\specs\TASKLIST-architecture-deepening-2.md` as they land, and flip the spec's
  summary row, its `### Spec 33` heading, and its status cell in `C:\data\ai\Xlibur\specs\README.md`.
- Spec 34 and any spec naming 33 a blocker gets its `**Status:**` and dependency line corrected.
- When the spec is finished, write its `## Results` section in
  `C:\data\ai\Xlibur\specs\33-sheet-listener-seam.md`: every recorded Excel-behaviour answer **and
  its source**, the structural-edit numbers from task 7, what the spec predicted that turned out
  wrong, what you deliberately did not do and why, and what the next consumer inherits.
- Two `CHANGELOG.md` entries under `## Unreleased` → `### Fixed`, one for tasks 4 and 6, one for
  task 5.
- Then copy `C:\data\ai\Xlibur\specs\` over the repo's `docs/specs` and include that in the PR, while
  that copy still exists.

## Autonomy

Work autonomously and report progress after each task. **Stop and ask the orchestrator only when a
decision would change what gets built:** an Excel-behaviour question you cannot settle from the
spec, an acceptance criterion that turns out arithmetically unreachable (spec 26's criterion 8 was —
**report it, do not delete documentation or copy a method body to satisfy a count**), or a hard
conflict.

Do not open the PR without being asked. When all seven tasks are green, report and wait.
