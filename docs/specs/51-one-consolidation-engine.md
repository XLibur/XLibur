# Spec 51 — One consolidation engine, two adapters

**Area:** Architecture · Refactor
**Effort:** S–M (~3 days)
**Dependencies:** **Spec 36 should land first.** The only live divergence between the two engines is
caused by the un-normalised geometry that spec 36 removes; running this spec first would collapse two
implementations onto one while that input still misbehaves.
**Status:** Proposed. From the 2026-08-30 architecture review (round 3). **Structural, not
defect-driven** — see the caveat below.

## Problem Statement

Merging a set of overlapping and adjacent rectangles into the smallest set of maximal blocks is
implemented twice: once over the value-typed rectangle, and once over live worksheet ranges.

Three of the seven method pairs are byte-identical. The rest differ only in how they reach a
rectangle's corners and in what they yield. The duplication is documented in a comment rather than
tested.

The two have disjoint consumers and disjoint test suites — conditional formats and data validations
use one, the public range consolidation uses the other — so nothing runs both on the same input, and
neither suite protects the other. A change to one is not a change to the other, and there is no
mechanism by which anyone would notice.

## Solution

One implementation, over the value-typed rectangle, with the live-range consumer becoming an adapter
that projects its ranges into rectangles, consolidates, and projects the results back.

This is the shape spec 26 established for the grid axis: one algorithm, two adapters, one place to
change and one place to test.

## User Stories

1. As a library consumer, I want consolidating a set of ranges to give the same blocks as consolidating
   the equivalent conditional format areas, so that the two features agree.
2. As a library consumer, I want overlapping ranges to merge into maximal blocks, unchanged from today.
3. As a library consumer, I want horizontally adjacent ranges to merge, unchanged from today.
4. As a library consumer, I want vertically adjacent ranges to merge, unchanged from today.
5. As a library consumer, I want ranges that touch only at a corner not to merge, so that adjacency is
   interpreted consistently.
6. As a library consumer, I want a set of ranges that cannot be merged to come back unchanged and in a
   predictable order, so that consolidation is stable.
7. As a library consumer, I want consolidation to be deterministic for a given input, so that saved files
   are reproducible.
8. As a library consumer, I want conditional formats and data validations to keep consolidating exactly as
   they do today, so that saved output is unchanged.
9. As a library consumer, I want consolidation performance to be no worse than today, so that saving large
   styled sheets does not slow down.
10. As an XLibur maintainer, I want the bit-matrix algorithm to exist once, so that a fix applies to both
    consumers.
11. As an XLibur maintainer, I want the two existing test suites to become two views of one
    implementation, so that each suite's cases protect the other's consumer.
12. As an XLibur maintainer, I want a differential test across both entry points, so that any future
    divergence fails a test rather than being discovered by a review.
13. As an XLibur maintainer, I want the live-range adapter to be thin enough to read in one sitting, so
    that the projection is obviously correct.
14. As a contributing agent, I want one consolidation implementation to find, so that I do not fix the
    one my consumer does not use.

## Implementation Decisions

**The seam is the existing value-typed consolidator.** It is the more constrained of the two — it
operates on immutable value types with no worksheet dependency, which makes it the easier one to test
and the natural implementation. The live-range engine becomes an adapter.

**The adapter projects both ways.** Live ranges become rectangles on the way in and are recreated from
the worksheet on the way out. That projection is the whole of the adapter; nothing else survives from
the second implementation.

**Enumeration order must be preserved.** The two implementations yield in the same order today, and
saved output depends on it. The adapter preserves the existing order for its consumer, and a
byte-comparison of saved output before and after is part of the acceptance evidence.

**No public API change.** The public consolidation method keeps its signature and behaviour.

**Performance is a gate.** Consolidation runs on every save for every sheet with conditional formats or
validations. The projection adds an allocation per range; if that proves measurable on the bulk-styling
benchmark, the projection is made allocation-free rather than the merge abandoned.

**This spec is scheduled after spec 36 deliberately.** The only divergence found between the engines is
that one handles a reversed rectangle and the other silently contributes nothing. That is spec 36's
defect, not this one's, and fixing it first means this spec is a pure refactor with a byte-identical
output guarantee — which is a much safer thing to review.

## Testing Decisions

**What makes a good test here.** A good test supplies a set of rectangles through a public entry point
and asserts the set of cells covered by the result — not the number of blocks or their order, except
where order is separately pinned. Asserting covered cells rather than block identity is what allows the
same test to run against both entry points.

**The centrepiece is a differential test.** For a generated corpus of multi-rectangle inputs, run both
public entry points — range consolidation, and consolidation via identical data validations — and
assert the covered-cell sets are equal. The review ran four hundred such cases as a one-off probe and
found no disagreement on normalised input; this spec makes that probe a permanent test.

**A byte-identity baseline.** Save a corpus of workbooks with conditional formats and validations
before the change, and assert byte-identical output after. This is the acceptance evidence that the
refactor is behaviour-preserving, and it is the same technique spec 22 used.

**Both existing suites must pass unchanged.** The range consolidation tests and the conditional format
consolidation tests are the regression net; neither should need editing. If one does, the refactor
changed behaviour and the change needs justifying in the results.

**Order stability tests.** Assert the yield order for a fixed input, so that a future change to the
projection cannot silently reorder saved output.

**Prior art.** Spec 26's grid axis is the model — one algorithm, two adapters, with the adapters
carrying only the projection. Its results section is worth reading for how the byte-identity baseline
was used.

**Test seam.** `IXLRanges.Consolidate()` and the conditional format and data validation save paths. No
new seam.

## Out of Scope

- Rectangle normalisation — spec 36, which this spec depends on.
- The consolidation algorithm itself. This spec merges two implementations of it; it does not improve
  it.
- The range index and its quad-tree, which are a different spatial structure.
- Adding consolidation to features that do not currently consolidate.

## Further Notes

**Honest caveat, and the reason this is not marked Strong.** A four-hundred-case differential fuzz
across both engines found *zero* disagreement on normalised input. The only divergence that exists
today is the one spec 36 removes. So this spec is not fixing a bug — it is removing the possibility of
one, in a place where the same algorithm is maintained twice by hand with nothing testing the
agreement.

That is a weaker case than the rest of this round and should be scheduled accordingly. It is included
because the cost is low, the risk is low once spec 36 has landed, and one hundred and fifty lines of
bit-matrix arithmetic maintained in duplicate is exactly the shape that produced four of this round's
confirmed defects elsewhere. The argument is prevention, and it should be presented as prevention
rather than dressed up as a fix.
