# Spec 38 — Sheet view state gets one module

**Area:** Architecture · **Defect (4 shipped)**
**Effort:** M (~4–5 days)
**Dependencies:** None. Checked against specs 29 and 31 — neither claims these properties.
**Status:** 🟩 Implemented on `task/38` (2026-08-30), unmerged — see Results. From the 2026-08-30 architecture review (round 3).

## Problem Statement

How a worksheet looks when it is opened — gridlines, headers, zoom, freeze panes, the view mode, the
tab colour — is described in six different places in the library, each written out by hand. Four of
the six have drifted from the other two.

What a user sees:

- **Copying a worksheet loses its appearance.** The copy comes back with gridlines on when the
  original had them off, with the default zoom rather than the original's, and in Normal view rather
  than the original's view mode. Copying between workbooks is worse: the copy picks up the *target*
  workbook's defaults rather than the source sheet's settings.
- **The view mode is lost on every round trip.** Open a workbook that Excel saved in Page Layout or
  Page Break Preview, save it, reopen it — it is in Normal view. The attribute is written on save and
  never read on load.
- **Zoom survives a save but not a copy.** The two paths disagree about the same four settings.

None of this is exotic. Copying a sheet and re-saving a file are two of the most common things anyone
does with the library.

## Solution

The view state of a worksheet becomes one module that owns the whole property list — the boolean
display flags, the four zoom scales, the view mode, the tab colour, and the pane and split settings it
already holds.

Loading, saving, copying and seeding a new sheet's defaults each become one operation against that
module rather than a hand-written list of properties. The properties stay exactly where they are on
the public worksheet interface and delegate inward, so nothing a consumer can see changes shape.

## User Stories

1. As a library consumer, I want a copied worksheet to keep its gridline setting, so that the copy
   looks like the sheet I copied.
2. As a library consumer, I want a copied worksheet to keep its row and column header setting, so that
   a header-less sheet stays header-less.
3. As a library consumer, I want a copied worksheet to keep its zoom level, so that the copy opens at
   the magnification I chose.
4. As a library consumer, I want a copied worksheet to keep its view mode, so that a sheet designed in
   Page Layout is still in Page Layout after copying.
5. As a library consumer, I want a copied worksheet to keep its tab colour, so that colour-coded
   sheets stay colour-coded.
6. As a library consumer, I want a worksheet copied into a *different* workbook to keep its own
   appearance rather than adopting the destination's defaults, so that cross-workbook copying is
   predictable.
7. As a library consumer, I want a workbook saved in Page Layout view to reopen in Page Layout view,
   so that re-saving a file does not change how it presents.
8. As a library consumer, I want a workbook saved in Page Break Preview to reopen the same way, so
   that a print-layout session survives a round trip through my application.
9. As a library consumer, I want all four zoom scales to survive both a save and a copy, so that the
   two operations agree.
10. As a library consumer, I want the formula-display, zero-display, outline-symbol, whitespace and
    tab-selected settings to survive a round trip, so that no display setting is silently dropped.
11. As a library consumer, I want freeze panes and splits to keep behaving exactly as they do today,
    so that this change is invisible where the library is already correct.
12. As a library consumer, I want every view property to remain on `IXLWorksheet` with the same name
    and type, so that my code does not have to change.
13. As an XLibur maintainer, I want adding a view property to be one edit rather than six, so that the
    next property cannot ship half-wired.
14. As an XLibur maintainer, I want each property's OOXML default polarity stated next to the property
    it belongs to, so that the omit-when-default rule cannot be got wrong in one place and right in
    another.
15. As an XLibur maintainer, I want a test that iterates the property list, so that a property added
    later is covered without anyone remembering to write a test for it.
16. As an XLibur maintainer, I want copy fidelity and round-trip fidelity to be verified by the same
    mechanism, so that the two cannot diverge again.
17. As a contributing agent, I want one obvious place to look for "what is a sheet's view state", so
    that I do not have to find six lists and diff them.

## Implementation Decisions

**The seam is the existing sheet-view type, promoted.** It already owns panes, splits, the top-left
address and the zoom scales. It absorbs the eight boolean display flags and the tab colour, which
currently live as fields directly on the worksheet. This was preferred to a new type: the seam already
exists, it is already the thing the writer reads from for half these properties, and promoting it
keeps the number of seams at one.

**The property list becomes data.** The module carries the list — name, accessor, OOXML default — and
the four operations are driven from it rather than each restating it. This is the same shape spec 39
proposes for the pivot definition and, like that one, it is what makes an enumerating test possible.

**Copying becomes a single operation on the module.** The worksheet's copy currently lists seven
things and stops; the sheet-view type's own copy constructor lists five and resets the rest by
delegating to the default constructor. Both are replaced by one copy of the view state.

**Default seeding is a distinct operation from copying, and stays one.** A new sheet takes workbook
defaults; a copied sheet takes the source's values. Today those two paths are conflated because the
copy silently falls through to the constructor's seeding. They become explicit and separate.

**The unread view-mode attribute gets a reader.** The enum mapping for it already exists in both
directions; only the load-side call is missing. Wiring it is a one-line consequence of the module
owning the list, and it removes the library's only currently-unused enum mapping in that area.

**No public API change.** Every property keeps its name, type and location on the public worksheet
interface and delegates. `PublicAPI.Shipped.txt` is untouched. This was a deliberate constraint on the
design, not an outcome.

## Testing Decisions

**What makes a good test here.** A good test sets view properties through the public worksheet
interface, performs an observable operation — save and reload, or copy — and asserts the properties
through the same interface. It does not inspect the module's internals and does not assert on the
emitted XML except where the omit-when-default rule is the thing under test.

**The centrepiece is a property-enumerating test, run twice.** Set every view property to a
non-default value; then (a) save and reload, and (b) copy the sheet — and in both cases assert that
every property came back. Because the module owns the list, the test can iterate it, so a property
added later is covered without a new test. This is the test the current shape cannot express, and its
absence is why four of six enumerations drifted.

**Named regression tests for the four defects**, so a failure identifies itself: gridlines lost on
copy, zoom lost on copy, view mode lost on round trip, cross-workbook copy adopting the wrong
defaults.

**One byte-level assertion for default polarity.** Several of these attributes are omitted when they
hold their OOXML default, and two are written with inverted polarity. A reload-based test cannot see
a polarity error because the reader inverts it back. One test that reads the emitted attributes
covers this, in the style of the existing write-path agreement tests.

**Prior art.** The worksheet element round-trip tests are organised one case per worksheet element,
which is why the sheet-view element has a single assertion covering one property and twelve others
are invisible to it. The sheet-view copy test asserts exactly one property — the one the copy
constructor happens to handle. Both tests are correct as far as they go; this spec makes them go
further.

**Test seam.** `IXLWorksheet`, worksheet copy, and a save/reload round trip. No new seam.

## Out of Scope

- Panes and splits, beyond keeping their current behaviour. Spec 29 resolved the frozen-pane
  divergence between the two write paths and that work stands.
- The other twelve responsibility clusters on the worksheet type. The review's conclusion was that
  most of the worksheet's width is the public interface and cannot move; view state is the one cluster
  that both can move and has defects.
- Print setup and page setup, which are a separate concern with their own module.
- The streaming write path's view handling, except where it shares the resolver work spec 29 already
  established.

## Further Notes

The worksheet type is the repository's single hottest file, and the last several commits touching it
are all structural-edit propagation work. It is not an undifferentiated blob: the clusters that can be
extracted have been extracted, one at a time — the outline tracker, the range shifter, the data
inserter, the merged-range and drawing-anchor listeners. View state is the conspicuous cluster that
has never had that treatment, and it is where the defects are. That is the argument for doing this
one next rather than any other part of the file.

The four defects were confirmed by reading both sides of each divergence; the unread view attribute
was verified directly against the reader and writer while writing this spec.

## Results

**Implemented 2026-08-30 on `task/38` (worktree `xl-wt-38`), head `ba9af10d`, 11 commits, cut from
`upstream/main` `37c986bb`. Not yet pushed or merged.** Full suite green on both TFMs: 28,516 total,
0 failed, 10 pre-existing skips. `PublicAPI.Shipped.txt` and `PublicAPI.Unshipped.txt` untouched.

**Named regression tests** (`XLSheetViewTests.cs`, all seen red at `286920d0`): `Copy_loses_gridlines`,
`Copy_loses_zoom`, `ViewMode_lost_on_round_trip`,
`CrossWorkbookCopy_keeps_source_appearance_not_targets_defaults`. Enumerating tests:
`AllViewProperties_survive_save_and_reload`, `AllViewProperties_survive_copy`, plus a themed
`TabColor` cross-workbook copy case. Byte-level: `SheetViewDefaultPolarityTests` (fresh-save and
load-mutate-resave, both directions), proven non-vacuous by inverting one writer condition and
watching all four fail.

**What the spec predicted that turned out wrong.** "Two attributes are written with inverted
polarity today" — **disproved.** All nine booleans were probed against the ECMA-376 `CT_SheetView`
defaults by reading the emitted bytes, left at default and flipped, fresh and after a
load-mutate-resave cycle. All nine were already correct. The polarity test now pins that; no writer
change was made for it. The four copy/round-trip defects were real and are fixed.

**What was done.** `sheetView/@view` is now read on load (`6e408f0f`, the one-line fix the spec
predicted). `XLSheetView` absorbed the nine display flags and the tab colour (`e45286e6`); the load
and DOM-save paths read and write them straight from the module (`be8c3d2a`, `7f5657d4`). Copy and
default-seeding are now **two explicit constructors** on `XLSheetView` — bare-default,
workbook-seed, copy — each listing every non-pane property once. The property list exists as typed
data: `XLViewProperty.All` (`XLibur/Excel/XLViewProperty.cs`, 15 entries — property, attribute,
default, polarity) drives both enumerating tests and doubles as the readable table.

**Deliberately not done, and why.** The four production consumers do **not** iterate
`XLViewProperty.All` at runtime — only the tests do. A reflection- or delegate-driven reader/writer
would have touched `WorksheetPartWriter`/`WorksheetElementReader` structurally, in the repo's hottest
file, for marginal benefit over two symmetric constructors that an enumerating test already polices.
Adding a property is one field, one line in each of three constructors, one `XLViewProperty.All`
entry — and the test catches a missed one. The streaming write path emits none of these attributes
(panes only, via `XLPaneSettings`), so `WritePathAgreementTests` is unchanged and has nothing new to
cover; whether streaming *should* emit them is spec 31's §5.1 question.

**Defects found outside the spec:** none.

**What the next consumer inherits.** Spec 31 rebases onto a `SheetViewWriter` whose flag emission
reads from `XLSheetView`, and inherits `SheetViewDefaultPolarityTests` as a byte gate for the
sheet-view element. Spec 29's pane resolver is untouched.
