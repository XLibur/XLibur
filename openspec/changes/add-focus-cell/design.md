## Context

A worksheet's on-open scroll position is governed by two OOXML attributes, and XLibur today exposes only one of them:

| Attribute | Meaning | XLibur surface today |
|---|---|---|
| `CT_SheetView/@topLeftCell` | Scroll of the (frozen-header) view as a whole | `SheetView.TopLeftCellAddress` |
| `CT_Pane/@topLeftCell` | Scroll of the *scrollable region* below/right of a freeze | **none** |

On a frozen sheet, `<pane topLeftCell>` is what parks the active cell off-screen. The current IO behaviour, verified in source:

- **Reader** (`WorksheetElementReader.LoadSheetViewPane`, ~line 73): reads only `HorizontalSplit`/`VerticalSplit` from the pane. `<pane topLeftCell>` is **never read** — it is discarded on load.
- **Writer** (`SheetViewWriter.SetupPane`, line 133): **unconditionally recomputes** `pane.TopLeftCell = GetColumnLetterFromNumber(SplitColumn + 1) + (SplitRow + 1)`.

Consequence — and a correction to the original framing of this change: XLibur does **not** preserve a loaded `<pane topLeftCell>`. A genuine load→save round-trip *self-heals* a parked pane to `split+1`. The persistent "parked at A966" symptom that motivated this work comes from code paths that **bypass** the XLibur worksheet writer (raw template byte passthrough), which no XLibur property can reach. This change therefore targets *write-path* consumers and gives them an explicit, supported knob — it does not, by itself, fix byte-passthrough export paths.

Honest accounting of what's genuinely new in the headline scenario (`FocusCell("A3")` on a row-frozen sheet that had `pane=A966`, `sheetView=G1`):
- `<pane topLeftCell>` becoming `A3` already happens today (A966 is dropped, recomputed to `split+1`).
- What is **new**: clearing the residual `<sheetView topLeftCell="G1">` horizontal scroll (orthogonal-axis reset), setting active cell + `sqref`, and — via the primitive — the *ability* to set a non-default pane anchor on purpose.

## Goals / Non-Goals

**Goals:**
- Expose `<pane topLeftCell>` as a first-class, nullable sheet-view property.
- Provide a one-call convenience that sets the active cell and makes it visible at the top-left of the scrollable region, frozen-aware.
- Deterministic, viewport-free output with literal, testable attribute values.
- Strictly additive; existing behaviour and the self-healing default are preserved.

**Non-Goals:**
- "Scroll so the cell is centred" or any margin/context rows — out of scope; this is top-left anchoring only. (Deferred: a future optional margin parameter can slot in additively if a real need arises.)
- Viewport math / "scroll only if off-screen" — XLibur has no rendered viewport, so it is not computable.
- Faithful preservation of an arbitrary loaded `<pane topLeftCell>` on bare round-trips (see Decision 1).
- True draggable split windows. XLibur's writer always emits `state="frozenSplit"` with row/column-*count* splits and exposes no path to a twip-based split, so split-window anchoring is not reachable and is not handled (see Decision 5).
- Multiple `<sheetView>` entries / split-window second views.

## Decisions

### Decision 1 — `PaneTopLeftCellAddress` is intent-only; the reader stays unchanged (normalize-to-top default)

`PaneTopLeftCellAddress` is nullable. `null` ⇒ writer keeps recomputing `split+1` (today's behaviour). A set value is honored. **The reader will NOT populate it** from `<pane topLeftCell>`.

- **Why:** Making the reader faithful (read + write the loaded value) would flip the default from *normalize-to-top* to *preserve-as-loaded*. That re-introduces the exact bug class on the plain load→save path: a template parked at `A966` would round-trip its bad scroll instead of self-healing to `A3`. The self-healing default is desirable and is what the motivating bug wants.
- **Alternative considered (rejected):** Faithful read+write for full round-trip fidelity of intentional scroll. Rejected for v1 because nobody has requested preserving a deliberate pane scroll, and the regression risk outweighs it. Can be revisited as an opt-in later.
- **Net:** "`null` = let XLibur anchor to `split+1`; set = honor it." This is a deliberate policy, not a no-op gap-closer.

### Decision 2 — Intent-named methods, not a `scrollIntoView` boolean

Scroll vs. no-scroll is expressed by *which method* the consumer calls, not by a control flag. This also sidesteps the `SetActive` overload trap (`IXLCell.SetActive(bool value = true)` already binds `value` positionally, so a `cell.SetActive(scrollIntoView: true)` form would be ambiguous).

- **Worksheet surface:**
  - `IXLWorksheet SetActiveCell(string address)` / `(IXLCell cell)` — sets active cell + single-cell selection (`sqref`); **never moves scroll**.
  - `IXLWorksheet FocusCell(string address)` / `(IXLCell cell)` — `SetActiveCell` **plus** scroll-into-view.
- **Fluent cell surface:**
  - `IXLCell SetActive(bool value = true)` — existing, **untouched**.
  - `IXLCell Focus()` — new; focus + scroll into view.
- **Why over a boolean:** `FocusCell("A3")` vs `SetActiveCell("A3", true)` — the verb states intent at the call site, avoids boolean-blindness, and matches the consumer's original `ws.FocusCell(cell)` instinct. One extra method name is cheaper than a behaviour-changing flag.
- **Alternative considered (rejected):** single `SetActiveCell(address, bool scrollIntoView = false)` + `SetActive(bool value, bool scrollIntoView)` overload. Rejected: control flag + overload ambiguity for no benefit.

### Decision 3 — "Make visible" = top-left anchor of the scrollable region

`FocusCell` sets the relevant `topLeftCell` to the target cell itself (frozen ⇒ `<pane topLeftCell>`; unfrozen ⇒ `<sheetView topLeftCell>`), and resets the orthogonal axis to its origin.

- **Why:** deterministic and literally assertable in tests; no viewport assumptions.
- **Trade-off:** anchoring e.g. `M50` to the top-left hides rows 1–49 / cols A–L. Accepted: "make-visible at top-left," not centring. A future optional small margin is an Open Question, not v1.

### Decision 4 — Selection/pane consistency for frozen-region targets

When the focus target sits inside the frozen rows/cols, the scrollable pane resets to `split+1` and the emitted `<selection pane=...>` SHALL name the pane that actually owns the active cell (top-left / top-right / bottom-left), rather than always using the computed scroll-region pane as `SetupSelections` does today.

- **Why:** avoids emitting an internally inconsistent `<selection pane="bottomLeft" activeCell="A1">` where `A1` lives in the frozen `topLeft` pane.
- **Scope correction (discovered during implementation):** this refinement is **general**, not focus-only. The writer has only `ActiveCell` to work from and cannot distinguish a `FocusCell` cursor from a plain `SetActive` one, so the active-pane/selection-pane is now derived from the active cell's position for *all* active cells. An earlier draft claimed "non-focus selection writing is unchanged" — that is not achievable, and satisfying the frozen-region scenario forces the general behaviour.
- **Impact on existing output:** the `FreezePanes` example's "Split View" sheet (active `B2`, split 3×3) previously emitted `activePane="bottomRight"`; it now emits `topLeft`, matching where the cursor actually sits (and what Excel itself produces). The committed reference fixtures `FreezePanes.xlsx` and `SheetViews.xlsx` were regenerated accordingly.

### Decision 5 — Frozen panes only; no split-window handling

Anchoring logic targets the frozen-pane case exclusively. `SheetViewWriter` sets `pane.State = PaneStateValues.FrozenSplit` unconditionally and treats `SplitRow`/`SplitColumn` as row/column counts; there is no public API path to a twip-based draggable split window, and `FreezePanes` is internal-only.

- **Why:** a true split window cannot be constructed through XLibur's public surface, so distinct split anchoring would be dead, untestable code.
- **Consequence:** the freeze-shape test matrix covers row-only, column-only, and both-axis freezes — not split windows. Spec wording says "frozen sheet," not "frozen or split."

## Risks / Trade-offs

- **[Default-flip regression]** If Decision 1 is implemented wrong (reader starts populating the property), bare round-trips would preserve parked scroll → reintroduces the bug. → Mitigation: explicit test asserting an unset property still yields `split+1`; reader left untouched and covered by a regression scenario.
- **[Doesn't fix byte-passthrough paths]** Export/template paths that don't serialize through XLibur are unaffected by any property here. → Mitigation: documented in proposal Impact; those paths remain the app-level normalizer's responsibility.
- **[Aggressive top-left scroll]** Top-left anchoring can hide leading rows/cols. → Mitigation: documented semantics; centring/margin explicitly out of scope and tracked as an Open Question.
- **[Pane-ownership complexity]** Decision 4 adds branching to selection emission. → Mitigation: confine to the focus path; cover each freeze shape (row-only, column-only, both-axis, split) with tests.

## Migration Plan

Purely additive public API; no migration for existing consumers. New API surface is added to `PublicAPI.Unshipped.txt`. No rollback steps beyond reverting the additive change; `null` default guarantees existing files serialize identically.

## Resolved Questions

- **Non-focused axis when not scrolling** → `SetActiveCell` never moves scroll. Since the reader never populates `PaneTopLeftCellAddress`, any non-null value is a deliberate consumer set, so leaving it untouched preserves intent. (Decision 2; backward compatible.)
- **Context margin** → Deferred. Strict top-left for v1; an optional margin can be added additively later if requested. (Now a Non-Goal.)
- **Naming** → Intent-named methods `FocusCell` (scrolls) and `SetActiveCell` (does not); no `scrollIntoView` boolean, no `SetActive` overload. (Decision 2.)
- **Split vs frozen panes** → Frozen only; XLibur cannot produce a split window. (Decision 5.)

## Open Questions

- None outstanding. Remaining specifics (exact owning-pane computation per freeze quadrant) are covered by the spec scenarios and the freeze-shape test matrix.
