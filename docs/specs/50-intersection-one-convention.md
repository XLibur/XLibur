# Spec 50 — One `Intersection`, one absence convention

**Area:** Architecture · **Correctness (undeclared error mode)** · API (behaviour clarification)
**Effort:** S (~2 days)
**Dependencies:** None hard. Soft ordering after spec 36 — that spec makes the geometry it relies on
consistent, though this spec does not depend on it.
**Status:** Proposed. From the 2026-08-30 architecture review (round 3). Smallest architecture spec of
the round.

## Problem Statement

Asking for the intersection of two ranges gives an answer whose *kind* depends on whether an optional
predicate was supplied.

| input | no predicate | with predicate |
|---|---|---|
| disjoint ranges | an invalid `#REF!` address, not null | null |
| overlapping ranges | the overlap | the overlap |

The declared return type is nullable, so the obvious caller check — test for null — is correct with a
predicate and wrong without one, where the disjoint answer is a non-null address that happens to be
invalid. Which convention applies is not documented; the interface says only that it returns the
address of the intersection.

There is a second, quieter difference. With a predicate, the result is not an intersection at all: it
is the bounding box of the cells that match. Two matching cells at opposite corners produce a
rectangle containing cells that do not match and were never in the intersection.

Both behaviours are unpinned — nothing in the suite asserts either convention.

## Solution

One implementation of intersecting two ranges, with the predicate applied as a filter over the result
rather than as a switch onto a different algorithm. One stated convention for "there is no
intersection", used in both cases and documented on the interface.

## User Stories

1. As a library consumer, I want the same absence convention whether or not I pass a predicate, so that
   one null check is correct everywhere.
2. As a library consumer, I want that convention documented on the interface, so that I do not have to
   read the implementation to know what an empty result looks like.
3. As a library consumer, I want disjoint ranges to report no intersection in a way I can test in one
   line, so that the common case is simple.
4. As a library consumer, I want overlapping ranges to give the overlap, unchanged from today, so that
   the working case is undisturbed.
5. As a library consumer, I want a predicate to narrow the intersection rather than change what
   intersection means, so that the parameter is a filter and not a mode switch.
6. As a library consumer, I want a predicate that matches nothing to report absence using the same
   convention as a disjoint intersection, so that the two empty outcomes look alike.
7. As a library consumer, I want a predicate that matches every cell to give the same result as no
   predicate at all, so that the parameter is neutral when it excludes nothing.
8. As a library consumer, I want the result of a predicated intersection to contain only cells inside the
   geometric intersection, so that the result is a subset of the unfiltered one.
9. As a library consumer, I want ranges on different worksheets to report absence consistently, so that
   cross-sheet calls follow the same rule.
10. As a library consumer, I want a range intersected with itself to return itself, so that the identity
    case is obvious.
11. As a library consumer, I want touching-but-not-overlapping ranges to report absence, so that adjacency
    is not mistaken for intersection.
12. As an XLibur maintainer, I want one implementation, so that a change to intersection semantics is one
    edit.
13. As an XLibur maintainer, I want the absence convention expressed in the signature, so that the type
    tells the caller what to check.
14. As an XLibur maintainer, I want the four-cell behaviour matrix asserted, so that the contract is
    pinned rather than emergent.
15. As a contributing agent, I want the predicate's role stated in one sentence, so that I do not have to
    infer it from two code paths.

## Implementation Decisions

**The seam is the existing public intersection method.** No new seam; this spec removes one
implementation rather than adding anything.

**The predicate becomes a filter over one algorithm.** Compute the geometric intersection, then apply
the predicate. The current predicated path computes something else entirely — the bounding box of
matching cells — which is why it can include non-matching cells.

**The bounding-box behaviour is a decision, not an accident, and must be made explicitly.** Two
readings are defensible: a caller filtering an intersection probably wants the matching cells, not a
rectangle around them; but a method returning a range address can only return a rectangle. The spec's
preference is that the predicated result be the smallest rectangle containing the matching cells
*within* the geometric intersection — which preserves the rectangle-returning contract while fixing
the case where the result escapes the intersection. Whichever is chosen is documented on the interface.

**One absence convention, and it is the invalid address.** Returning an invalid `#REF!` address is
consistent with how the library represents an unresolvable reference elsewhere, and it is what the
more commonly used unpredicated path already does. The predicated path changes to match. The return
type stops being nullable, which makes the convention visible in the signature rather than documented
in prose.

**This changes observable behaviour for predicated callers**, who currently receive null. It is a small
surface and the current behaviour is undocumented and untested, but it is a change and belongs in the
release notes.

**The interface documentation is part of the deliverable.** The current summary describes neither
convention. That omission is the reason both behaviours could coexist unnoticed.

## Testing Decisions

**What makes a good test here.** A good test calls the public method with two ranges and asserts on the
returned address — its validity, and its extent. There is nothing internal worth testing; the whole
defect is in the observable contract.

**The centrepiece is the behaviour matrix.** Disjoint and overlapping, by predicate and no predicate:
four cases, asserting both the absence convention and the extent. This is the matrix the review used to
characterise the divergence and it becomes the contract test.

**Extended cases.** Identity, adjacency without overlap, a predicate matching everything, a predicate
matching nothing, a predicate matching only opposite corners, and ranges on different worksheets.

**The corner case is the one that matters.** A predicate matching two opposite corners of the
intersection asserts what the predicated result means. It is the case that distinguishes the three
possible designs, so it is the case the test must state.

**Prior art.** The range address tests are the right home. Note that neither convention is currently
asserted anywhere, so this spec creates the coverage rather than extending it.

**Test seam.** `IXLRangeBase.Intersection`. No new seam.

## Out of Scope

- Range geometry and normalisation — spec 36. This spec assumes whatever geometry it is given.
- The other set operations on ranges — union, difference, grow, shrink.
- Implicit intersection in the calc engine, which is a different concept with the same name and is
  covered by spec 37.
- Performance.

## Further Notes

Small, but a good example of a specific failure: an optional parameter that changes not just what is
computed but what the return value *means*. A caller reading the signature sees a nullable address and
writes the obvious check; that check is right half the time, and which half depends on an argument they
may have passed for unrelated reasons.

The fact that neither convention is asserted anywhere is what allowed two to coexist. A single test of
the disjoint case, written at any point, would have forced the question.

The four-cell matrix was measured against a scratch build before this spec was written.
