# Spec 09 — Threaded Comments: Full Model + Write Path (and Round-Trip Fidelity)

**Area:** Feature + Compatibility
**Effort:** M (1–2 weeks)
**Dependencies:** None.
**Status:** Implemented — see [Implementation notes](#implementation-notes)

## Summary

Threaded comments (Office 365's reply-style comments) are currently read **lossily**: the whole thread is flattened into the legacy note's text joined by newlines, discarding authors, timestamps, reply structure, and mentions — and there is no write path, so saving a file that contained threaded comments silently downgrades them to a legacy note. This spec adds a first-class threaded-comment model with full round-trip, plus a small fidelity audit for other silently-dropped content.

## Current state

- Read: `LoadThreadedComments` / `ApplyThreadedCommentsToCell` in `XLibur/Excel/XLWorkbook_Load.cs` (~lines 736–786) — flattens thread → legacy note text.
- Write: nothing. No `WorksheetThreadedCommentsPart`, no `PersonPart`/`personList` anywhere in the codebase.
- Legacy notes are complete: `XLibur/Excel/Comments/` + `IO/CommentPartWriter.cs` + `IO/VmlDrawingPartWriter.cs`.
- Related fidelity gaps (audit scope, task 6): chartsheets/dialog sheets/macro sheets land in `XLWorkbook.UnsupportedSheet` (`XLWorkbook.cs` ~line 784, `XLWorkbook_Load.cs` ~line 261) and are **dropped on save**; form controls/ActiveX are dropped; slicers/timelines are dropped.

## OOXML background (what Excel writes)

- `xl/threadedComments/threadedComment{N}.xml` (one per sheet with threads): `<threadedComment ref="A1" dT="..." personId="{guid}" id="{guid}" [parentId="{guid}"]>` with `<text>`; replies reference the root via `parentId`.
- `xl/persons/person.xml` (workbook-level `PersonPart`): `<person displayName="..." id="{guid}" userId="..." providerId="..."/>`.
- **Compatibility pairing:** Excel also writes a legacy comment (+ VML) fallback for each thread root, whose text is prefixed with `[Threaded comment]` boilerplate. Older Excel shows the fallback; 365 shows the thread. XLibur must write both, like Excel does, and must **hide the fallback note from the user model** when the threaded part is present (current reader already prefers the threaded content — keep that, but stop flattening).

## Design

### Public API (new files under `XLibur/Excel/Comments/Threaded/`)

```csharp
public interface IXLThreadedComment
{
    string Text { get; set; }
    IXLPerson Author { get; }
    DateTime CreatedUtc { get; }
    IXLThreadedComment? Parent { get; }                 // null for thread root
    IReadOnlyList<IXLThreadedComment> Replies { get; }
    IXLThreadedComment AddReply(IXLPerson author, string text);
    bool Resolved { get; set; }                          // root only ('done' flag)
}

public interface IXLPerson { string DisplayName { get; } string? UserId { get; } string? ProviderId { get; } Guid Id { get; } }

// IXLCell additions:
IXLThreadedComment? GetThreadedComment();
IXLThreadedComment CreateThreadedComment(IXLPerson author, string text);
bool HasThreadedComment { get; }

// IXLWorkbook addition:
IXLPersons Persons { get; }   // Add(displayName), lookup by id
```

Design rules:
- A cell has *either* a legacy note *or* a threaded comment as its user-visible annotation; creating a threaded comment on a cell with a legacy note throws (or converts — decide and document; **recommend throw** for v1, explicit `ConvertToThread()` later).
- Storage: extend the comment slot in the misc slice area. `XLMiscSliceContent` (`XLibur/Excel/Cells/XLMiscSliceContent.cs`) holds `XLComment?` today; either widen the type to a shared base/`object` holding note-or-thread, or add a per-worksheet side dictionary keyed by `Point` (threads are rare — a side dictionary avoids fattening the misc struct for every cell; **recommend side dictionary**, consistent with how rare data should be stored per the architecture's slice philosophy).
- Mentions (`<mention>` runs): **out of scope v1** — preserve raw XML for round-trip (store the original `<text>` payload when unmodified) but no mention API.

### Read path rework

`LoadThreadedComments`: build the model (persons, roots, replies ordered by `dT`), attach to cells, and **do not** create the flattened legacy text. The paired legacy fallback note is detected (same ref as a thread root) and suppressed from `cell.GetComment()`.

### Write path

New `ThreadedCommentPartWriter` + `PersonPartWriter` in `XLibur/Excel/IO/`:
1. Emit `person.xml` for all referenced persons.
2. Emit per-sheet `threadedComment{N}.xml` with stable GUIDs (persist loaded ids; generate for new).
3. Emit the legacy fallback note + VML for each thread root (reuse `CommentPartWriter`/`VmlDrawingPartWriter` with the `[Threaded comment]` boilerplate text Excel uses).
4. Wire into `XLWorkbook_Save.cs` part orchestration + content types/relationships.

## Work plan

| # | Task | Size |
|---|------|------|
| 1 | Model + API (`IXLThreadedComment`, `IXLPerson`, cell/workbook surface) + side-dictionary storage | M |
| 2 | Reader rework (structure-preserving; fallback-note suppression) | M |
| 3 | Writers (persons, threads, legacy fallback pairing) + save orchestration | M |
| 4 | Tests: Excel-authored file with threads/replies/resolved flags in `XLibur.Tests/Resource/` → load asserts structure; round-trip byte-level part comparison where stable; new-thread-from-scratch opens in Excel (manual check recorded in PR) | M |
| 5 | Editing semantics tests: delete thread root deletes replies; copy/move cell behavior (match legacy-note behavior); `CopyTo` across workbooks maps persons | S |
| 6 | **Fidelity audit (separate PR):** enumerate content dropped on round-trip (chartsheets, form controls, slicers, timelines, custom XML). For each: document, and where cheap, **preserve the raw part** through load→save instead of dropping (part pass-through for unmodified sheets is often near-free with the OpenXML SDK since unknown parts stay in the package when not removed — verify what XLibur's save actually removes vs preserves, and stop deleting what it doesn't understand). Deliverable: `docs/round-trip-fidelity.md` + pass-through fixes where trivial. | M |

## Acceptance criteria

1. Excel-authored threaded comments load with authors, UTC timestamps, reply order, and resolved state intact; `GetComment()` no longer returns flattened thread text.
2. Round-trip: load → save → open in Excel 365 shows identical threads; open in Excel 2016 shows the fallback notes.
3. Threads created via API from scratch open correctly in Excel 365 (manual verification recorded in PR).
4. Legacy-note behavior unchanged for files without threads (full existing comment test suite green).
5. Fidelity audit doc merged; at least chartsheet part pass-through evaluated with a clear do/can't-do conclusion.

## Risks

- Excel is picky about person/thread GUID and relationship wiring — the manual open-in-Excel check is the gate, automate reopen-via-XLibur for CI.
- The legacy-fallback pairing is underdocumented; copy exactly what current Excel emits (author a sample file, inspect parts, mirror it).

## Implementation notes

All six tasks landed. Full suite green (6650 tests, 0 failures) across net8.0/net9.0/net10.0.

### Manual verification (acceptance criteria 2 and 3)

Four workbooks were generated and opened in Excel 365 — all four opened cleanly with no repair
prompt: threads created from scratch via the API (including replies from multiple authors, a
resolved thread and a non-ASCII display name, across two sheets); an Excel-authored file
round-tripped unchanged; the same file with its root text edited and a reply appended; and a sheet
mixing a legacy note with a thread.

### Deviations from the design above

1. **Storage: misc slice, not a side dictionary.** The design recommended a per-worksheet side
   dictionary keyed by `Point`. It went into `XLMiscSliceContent` instead, because row/column
   shifting, `Swap` and area-clear all iterate `XLCellsCollection._slices` — a side dictionary would
   have needed every one of those reimplemented by hand. The slice is sparse, so the cost is 8 bytes
   per cell that already holds misc content. All 20 editing-semantics tests passed on the first run
   with no additional code, which is the payoff.
2. **The fallback note omits `shapeId`.** Excel writes `shapeId="0"` on a thread's fallback comment,
   but the schema the OpenXML SDK validates against does not declare the attribute, so every
   `validate: true` save failed with it present. It is optional and Excel ignores it; the pairing is
   carried by the `tc={rootId}` author and the `xr:uid`.
3. **`done` is written only on the thread root**, matching Excel, so `Resolved` reads through from a
   reply but throws when set on one.

### Correction to the current-state analysis

Two claims in this spec turned out not to match the code:

- **Threaded comment parts were already surviving a round trip.** Saving reopens the original package
  and rewrites only modelled parts, so `threadedComments/` and `persons/` were carried through — but
  `comments1.xml` was regenerated without the `tc={rootId}` author or `xr:uid`, so the result was an
  *inconsistent pairing* rather than a clean downgrade to a note.
- **Chartsheets are not dropped on save.** `WorkbookPartWriter` reorders modelled sheets around the
  unsupported ones rather than rewriting the `<sheets>` list, so both the `<sheet>` entry and the
  part survive. Task 6 therefore found no pass-through work to do; its deliverable is
  `docs/round-trip-fidelity.md` plus regression tests locking the behaviour in. Slicers remain
  unverified — the repo has no fixture containing `xl/slicers/`.
