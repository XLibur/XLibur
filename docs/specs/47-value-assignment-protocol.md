# Spec 47 — One implementation of assigning a value to a cell

**Area:** Architecture · **Defect (quote prefix survives; a public method skips invalidation)**
**Effort:** S–M (~3 days)
**Dependencies:** None hard. Touches the calc-engine invalidation call that spec 42 relocates; if both
are scheduled, 42 lands first and this spec calls the relocated seam.
**Status:** Proposed. From the 2026-08-30 architecture review (round 3).

## Problem Statement

Assigning a value to a cell is not one operation. It is a short protocol: work out whether the value
requires a style change, apply that style, strip the leading apostrophe that marks text as literal
because the style now records it, store the value, and tell the calc engine that dependents are stale.

The rule that decides the style is a single well-factored module. The protocol around it is written out
four times — in the cell type, in the worksheet's direct value setter, in the bulk data inserter and in
the streaming writer — and one comment in the library even notes that its copy is "the same as" another
one.

Two of the four are wrong:

- **The bulk inserter keeps the apostrophe.** It performs the strip only inside a branch that runs when
  the style actually changed. Style values are interned, so when the inherited style *already* records
  the quote prefix, nothing changes, the branch is skipped, and the apostrophe stays in the stored
  text. Setting the same value through a cell strips it. The two disagree:

  ```
  ws.Column(1).Style.IncludeQuotePrefix = true;
  ws.Cell("A1").Value = "'abc";                 // stored: abc
  ws.Cell("A2").InsertData(new[] { "'abc" });   // stored: 'abc
  ```

- **The worksheet's direct value setter never invalidates.** It is public and shipped. Its documentation
  lists three things it deliberately skips and does not mention this one. A value written through it
  leaves dependent formulas stale, once anything in the workbook has been evaluated. Its only caller in
  the repository is a benchmark, and it has no tests.

## Solution

The protocol joins the rule. A caller supplies a value and a location and receives back what to store
and what style to set; it no longer decides how to strip, when to strip, or whether to invalidate.

Callers keep whatever performance shortcuts they have around that call — the bulk inserter still writes
in bulk, the streaming writer still streams — but they stop re-deciding the rule inside it.

## User Stories

1. As a library consumer, I want text beginning with an apostrophe to be stored the same way whether I
   assign it to a cell or insert it in bulk, so that the two entry points agree.
2. As a library consumer, I want that to hold when the target already carries the quote-prefix style, so
   that pre-existing formatting does not change how my data is stored.
3. As a library consumer, I want the same behaviour through the streaming writer, so that all three
   write routes agree.
4. As a library consumer, I want a value that looks like a number, a date or a boolean to receive the
   same inferred style through every entry point, so that bulk and single writes format alike.
5. As a library consumer, I want a value written through the worksheet's direct setter to invalidate
   dependent formulas, or to be documented as not doing so, so that I can use it correctly.
6. As a library consumer, I want a formula that reads a bulk-inserted range to recalculate, so that
   `InsertData` composes with the calc engine.
7. As a library consumer, I want inserting data over existing formulas to behave as documented, so that
   overwriting is predictable.
8. As a library consumer, I want bulk insertion to stay as fast as it is today, so that correctness here
   does not cost throughput.
9. As a library consumer, I want a value assigned to a cell to keep behaving exactly as it does today,
   so that the entry point that is already correct is not disturbed.
10. As an XLibur maintainer, I want the strip rule stated once, so that a gate around it cannot be added
    in one copy and not another.
11. As an XLibur maintainer, I want the invalidation rule stated once, so that an entry point cannot omit
    it.
12. As an XLibur maintainer, I want a new write path to inherit both rules, so that the next one cannot
    drift.
13. As an XLibur maintainer, I want the worksheet's direct setter's documented exclusions to match what it
    actually skips, so that its contract is true.
14. As an XLibur maintainer, I want that public method to have tests, so that a shipped entry point is not
    entirely uncovered.
15. As an XLibur maintainer, I want one table-driven test covering all four entry points, so that a
    divergence between them fails a test.
16. As a contributing agent, I want one documented way to assign a value, so that I do not copy whichever
    of four protocols I find first.

## Implementation Decisions

**The seam is the existing value-style rule module, widened to own the protocol.** It currently answers
"what style should this value have" and returns nothing when the style is unchanged. It grows to answer
"what should be stored and what style should be set", which is the question its callers actually have.

**The interned-style return is the mechanism of the defect and is designed out.** Returning "no change"
and returning "no strip needed" are currently the same value, which is why one caller conflated them.
After this spec, the value to store and the style to set are independent parts of the answer.

**Invalidation becomes part of the protocol, with an explicit opt-out.** The load path and the streaming
writer legitimately do not invalidate — the first has nothing stale, the second has no engine. That is
a named mode, not an omission.

**The worksheet's direct setter's contract is fixed either way, deliberately.** Either it invalidates
like its siblings, or its documentation states that it does not and why. What it must not remain is a
public method whose documented exclusions are incomplete. The spec's preference is that it invalidate,
since it is a general-purpose public entry point; if measurement shows that defeats its purpose, the
finding is recorded and the documentation corrected instead.

**Callers keep their shortcuts.** The bulk inserter marks its whole inserted region dirty once rather
than per cell, and that stays — this spec unifies the rule, not the batching.

**Performance is a gate.** Bulk insertion and the create path are benchmarked. Measure before and after.

## Testing Decisions

**What makes a good test here.** A good test writes a value through one of the four public entry points
and asserts what is stored and what style resulted, through the public interface. It does not test the
rule module in isolation — the rule was always right; the defect is in how it was called, which is
exactly the case a pure-function unit test cannot see.

**The centrepiece is an entry-point by input matrix.** For each of the four entry points, and for each
input — text with a leading apostrophe, the same with the quote-prefix style already applied to the
target, ordinary text, a number, a date, a boolean, an empty string — assert the stored value and the
resulting style. The second input is the defect; the rest are the regression net.

**Dependent-recalculation cases.** For each entry point, write a value that a formula depends on and
assert the formula updates without a forced full recalculation.

**First tests for the direct setter.** It is public, shipped, and has none. Whatever its contract ends
up being, it gets tests asserting it.

**A performance guard.** The bulk insertion benchmark before and after, recorded in the results.

**Prior art.** The cell value tests are the right home and demonstrate the correct behaviour, which makes
them the control arm for every new case. The rule module's own tests are correct and stay; they are just
not sufficient, and that is the point worth recording.

**Test seam.** `IXLCell.Value`, `IXLCell.InsertData`, `IXLWorksheet.SetCellValue`, and the streaming
worksheet. No new seam.

## Out of Scope

- The style inference rules themselves, which are correct.
- Number format detection and culture handling.
- The bulk inserter's batching strategy and its dirty-marking granularity.
- Formula writes, which are spec 42.
- Adding new value types.

## Further Notes

This candidate is a useful counter-example to the assumption that extracting a pure function makes
something testable. The rule *is* extracted, it *is* a pure function, and it *does* have unit tests
that pass. The bug is one branch away from it, in a caller, in a condition that looks reasonable and is
wrong only because of a property of the value being returned. Locality did not follow the extraction:
the knowledge stayed in four places while only the calculation moved to one.

The second leg — a shipped public method that skips invalidation, has one benchmark as its only caller,
and no tests — is worth resolving in this spec rather than leaving as a note, because whichever way it
is resolved it is a one-line change plus documentation, and leaving it ambiguous is what allowed it to
sit unexamined.

The divergence was reproduced through public API before this spec was written; the missing invalidation
was established by tracing all four entry points to the engine.
