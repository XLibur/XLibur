# Dispatch brief — spec 29, write-path resolvers

**Dispatched:** 2026-08-26 · **Worktree:** `D:\Data\_CodeOS\xl-wt-29` · **Branch:**
`refactor/29-write-path-resolvers`, created off `upstream/main` at `806d69f7`. Never commit to `main`.

**Spec:** `C:\data\ai\Xlibur\specs\29-write-path-resolvers.md`. That folder is the source of truth for
specs; the repo's `docs/specs` is a copy on its way out. Read the spec in full and work its seven
tasks in order.

## What this spec does

XLibur has two write paths — the ordinary DOM save and `XLStreamingWorkbook` — and they share exactly
one seam (`CellXmlWriter`) and re-implement everything above it. Four elements have two
implementations that must agree by hand, and **one of them already disagrees in shipped code**: the
DOM path writes `state="frozenSplit"` for every pane while the streaming path writes `state="frozen"`.
The reader normalises both back to the same model, so no load-and-compare test can see it. This spec
extracts the *decision* into value-in/value-out resolvers both paths consume, and delivers a
cross-path agreement test that reads the bytes.

**The agreement test is the artefact this spec is really delivering.** The resolvers are how the
agreement becomes structural rather than coincidental.

## The spec's line numbers are still good — verified today

Unusually for this repo, the spec's "Current state" section survives specs 26 and 28 intact.
Confirmed against the worktree at `806d69f7`:

- `SheetViewWriter.cs:124` — `pane.State = PaneStateValues.FrozenSplit;` **still there**, 280 lines
- `XLStreamingWorksheet.cs:502` — `xml.WriteAttributeString("state", "frozen");` **still there**, 575 lines
- `XLStreamingWorkbook.cs:225` — the fabricated `new SaveContext()` **still there**, 366 lines
- `ColumnWriter.cs` — 282 lines, as the spec says
- `WorkbookStylesPartWriter.cs` — **1212 lines, not 1211**; spec 28 (PR #411/#412) moved code in this
  file. `GenerateContent` is still at `:18` and `GenerateStreamingContent` still at `:100`, but
  **re-verify `ResolveFonts` and the helper line numbers in task 5 before editing** — that is the one
  file in this spec that spec 28 touched.

Verify anything else you rely on rather than trusting the number.

## Ordering constraints

- **Task 1 is the important one and it lands FAILING.** That is deliberate: it is a defect report
  that runs. Do not fix anything in task 1. Do not weaken an assertion to make it green.
- **Task 1 step 2 decides two premises the spec has not confirmed.** Record both verbatim:
  1. The `xSplit="0"` claim — does the DOM path really write `xSplit="0"` for a rows-only freeze where
     streaming omits the attribute? If both omit it, **say so and drop the split-omission rule from
     task 3's resolver.** A disproved premise is a result, not a problem (acceptance criterion 13).
  2. If `Both_write_paths_agree_on_a_column` **passes** on the first run, that is also a result:
     `<col>` is already in agreement for that case and task 4 becomes a refactor with no defect behind
     it. Say so and **keep task 4 anyway** — the point is that agreement stops being coincidental.
- **Task 2 is one line.** If an existing test fails on it, it is asserting the state indirectly through
  a file comparison or an `output.xlsx` fixture. **Do not revert.** A reference `output.xlsx`
  containing `frozenSplit` is an XLibur-authored artefact of this defect, not evidence about Excel —
  regenerate it and name which fixture changed.
- **Task 4 step 2 has a double-rounding trap.** `ctx.WorksheetColumnWidth` is already through
  `GetColumnWidth().SaveRound()` at `SheetViewWriter.cs:264`, so passing it into `Resolve` as a raw
  width rounds twice. Either pass the raw worksheet width or add an overload taking an already-resolved
  width — and prove whichever you choose with task 1's column test.
- **Task 6's deliverable is a recorded decision, either way.** A recorded "no resolver, and here is the
  evidence with line references" is a complete result. Do not manufacture an `XLSheetSettings` that
  carries three verbatim fields with no defaulting — that is indirection without a decision.
- **Task 7 gates streaming's bounded memory,** not speed. Spec 01 records ~40% run-to-run timing
  variance on this machine: treat elapsed as indicative and **peak heap as the real gate** (107.9 MB
  shared strings / 14.0 MB inline strings at 1M × 10).

## Inherited decisions — do not re-open

- **Not merging the two write paths.** Two adapters over one resolved value is the target. The DOM path
  exists so unmodelled markup survives a round trip; the streaming path exists for bounded memory and
  owns its own `ZipArchive` in Create mode. Neither can become the other. Both serializers stay
  (acceptance criterion 8).
- **Spec 01 declined the wide `IXLSheetDataSource` seam** in favour of sharing the *leaf* serializers.
  This spec applies that rule one level up: **share the leaf decision, not the traversal.** Read spec
  01's Results before starting.
- **The resolvers do not own *which* columns get written.** The DOM path expands, back-fills and
  collapses runs; streaming writes one `<col>` per registered range. Different products, both stay.
  Only the per-`<col>` attribute decision is shared.
- **Do not touch `CellXmlWriter`** (already the right shape, and spec 03's territory), **`SheetDataWriter`
  internals** (spec 01/03), or **`WorksheetElementReader.LoadSheetViewPane`** — it must keep accepting
  both spellings, because files in the wild carry both (acceptance criterion 6).
- **No public API change.** `PublicAPI.Unshipped.txt` untouched.

## Concurrency — spec 33 is running right now in `xl-wt-33`

It owns `XLWorksheet.cs`, `XLWorksheetRangeShifter.cs`, `XLSheetView.cs` (**the model type — not
`SheetViewWriter.cs`**), `Cells/ISheetListener.cs`, and the drawings, pivot-table, conditional-format,
data-validation, defined-name, page-setup, sparkline, hyperlink and calc-engine collection types.
**Do not edit any of those.** Reading `xlWorksheet.SheetView.SplitColumn` from `SheetViewWriter` is
fine; changing `XLSheetView.cs` is not. Spec 29 and spec 33 are otherwise file-disjoint.

**Spec 31 is blocked on this spec** and must not start until it lands — 31 rewrites the interface
`SetupPane` and `BuildColumnElement` sit behind, and if it landed first it would carry
`PaneStateValues.FrozenSplit` forward into new structure built around the wrong value.

## Ground rules this repo will punish you for ignoring

Read `D:\Data\_CodeOS\xl-wt-29\CLAUDE.md` before your first edit. In particular:

- **No compound shell commands** (`&&`, `||`, `;`) in tool calls. Never `cd <folder> && git ...` —
  use `git -C <path>`.
- **Never read-modify-write a whole tracked file.** `sed -i`, `python`, `perl -i`, `Set-Content` and
  shell redirection over the original are all out — `.gitattributes` checks out CRLF and they write
  LF, turning a one-line change into a whole-file diff. Use the Edit/Write tools and verify with
  `git diff --numstat`. Task 2 step 4 is explicit about this: a changed-line count near 280 on
  `SheetViewWriter.cs` means the file was rewritten, not edited — discard and redo.
- The runner is **TUnit on Microsoft.Testing.Platform**, not VSTest: filtering is
  `--treenode-filter "/*/*/ClassName/*"`, **never** `--filter`. Exit 5 = invalid option, exit 8 =
  zero tests matched. Never filter at solution level — name the `.csproj`.
- **Assertions are awaitable.** A missing `await` on `Assert.That(...)` means the assertion never runs
  and the test passes regardless. Treat CS4014 as an error.
- `TreatWarningsAsErrors=true`, nullable enabled. New code must be null-annotated.
- The test project targets **net8.0 and net10.0 only**. "Green on net9.0" is unsatisfiable, not a
  missed run.
- **Diff against the merge base, never the moving upstream tip:**
  `git -C /d/Data/_CodeOS/xl-wt-29 diff --numstat $(git -C /d/Data/_CodeOS/xl-wt-29 merge-base upstream/main HEAD)`

Gate after every task: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Run unfiltered (both TFMs) before opening the PR.

## Documentation — the work is not done until this is done

**A second agent (spec 33) is editing the same conductor folder concurrently.** So:

- **Append only** to `C:\data\ai\Xlibur\DEFECTS.md`. Record any defect you find that is not the work in
  front of you, with its call sequence and consequence, then carry on — do not let it widen the task.
- For `C:\data\ai\Xlibur\specs\README.md` and
  `C:\data\ai\Xlibur\specs\TASKLIST-architecture-deepening-2.md`, **edit only your own rows**, with
  targeted single-line Edit operations, and **re-read the file immediately before each edit**. Never
  rewrite either file wholesale.
- `C:\data\ai\Xlibur\specs\29-write-path-resolvers.md` is yours alone. Tick its task boxes with the
  commit sha as they land, and write its `## Results (2026-08-26)` section covering: the task 1
  first-run readings attribute by attribute, whether the `xSplit="0"` premise held, whether the column
  test passed cold, task 6's decision on `<sheets>` and styles.xml **with line references**, task 7's
  memory numbers, what you deliberately did not do and why, and what spec 31 inherits.
- **`CHANGELOG.md` in the repo:** task 2 changes output for existing documents, so it needs an entry
  under `## Unreleased` → `### Fixed`. Spec 33's branch is also adding entries there — expect a
  trivial merge conflict and resolve it by keeping both.
- Then copy `C:\data\ai\Xlibur\specs\` over the repo's `docs/specs` and include that in the PR, while
  that copy still exists.

## Autonomy

Work autonomously and report progress after each task. **Stop and ask the orchestrator only when a
decision would change what gets built:** the task 4 double-rounding choice if task 1 cannot settle it,
an `output.xlsx` fixture that would need regenerating, an acceptance criterion that turns out
arithmetically unreachable (report it — do not delete documentation or weaken a test to satisfy a
count), or a hard conflict with spec 33.

Do not open the PR without being asked. When all seven tasks are green, report and wait.
