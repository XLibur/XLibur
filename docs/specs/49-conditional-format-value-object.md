# Spec 49 — One conditional format value object

**Area:** Architecture · **API (breaking)** · **Defect (silent misalignment)**
**Effort:** M–L (~1–1.5 weeks)
**Dependencies:** **Spec 48 should land first.** It fixes the crash and the data loss without an API
change; this spec then fixes the shape that made them easy to write. Running this one first would
block a crash fix behind an API decision.
**Status:** Proposed. From the 2026-08-30 architecture review (round 3). Split from spec 48 at the
owner's direction.

## Problem Statement

A conditional format's thresholds are exposed on the public interface as four separate collections —
the values, the colours, the content types and the icon-set operators — which the caller must keep
index-aligned. Nothing enforces the alignment, and the reader does not maintain it: it appends to
three of the collections unconditionally and to the fourth only when the corresponding attribute is
present.

The result is that a single threshold with no explicit type shifts every later entry onto the wrong
type. An icon set or colour scale re-saves with its thresholds attached to the wrong comparison — no
error, no exception, just a rule that now means something else.

The same shape shows up in how a data bar is assembled: its two halves live in different namespaces
and are joined by matching a string identifier, so the information about one visual element is split
across two converter hierarchies with nothing tying them together.

For a caller, the interface is as wide as the implementation: four collections, a one-based indexing
convention, an alignment requirement, and a per-format-type arity rule, all of which they must know and
none of which the types express.

## Solution

A threshold becomes one thing — its type, its value, its colour and its operator together — held in
order. Alignment is a property of construction rather than a rule the caller follows. The arity rules
for each conditional format type live on that concept.

The data bar becomes one module with two spellings, the way spec 44 proposes for data validation,
rather than two converter hierarchies joined by a string match.

## User Stories

1. As a library consumer, I want a conditional format's thresholds to be a single ordered collection, so
   that I cannot misalign them.
2. As a library consumer, I want each threshold to carry its own type, value, colour and operator, so
   that reading one threshold does not mean indexing four collections.
3. As a library consumer, I want a threshold with no explicit type to take the format's default rather
   than shift its neighbours, so that a valid file loads with the right meaning.
4. As a library consumer, I want an icon set loaded from a file to re-save with the same thresholds
   attached to the same types, so that the rule keeps its meaning.
5. As a library consumer, I want the same for a colour scale, so that the two types behave alike.
6. As a library consumer, I want an invalid threshold count for a given format type to be reported when I
   build it, so that I find out at the point of the mistake.
7. As a library consumer, I want a three-stop colour scale and a two-stop colour scale to both be
   expressible and both validated, so that arity is enforced rather than assumed.
8. As a library consumer, I want a data bar's main and extension settings to be one object, so that I do
   not have to know the format has two halves.
9. As a library consumer, I want to set a data bar's border and bar lengths through the same object as
   its colours, so that one visual element has one home.
10. As a library consumer, I want a migration path from the four collections, so that upgrading is
    mechanical.
11. As a library consumer, I want the change documented with before-and-after examples for each
    conditional format type, so that I can update my code without guessing.
12. As an XLibur maintainer, I want the alignment invariant enforced at construction, so that no reader
    or converter can break it.
13. As an XLibur maintainer, I want the arity rules stated once per format type, so that each converter
    stops assuming them.
14. As an XLibur maintainer, I want a threshold testable on its own, so that the type-missing and
    value-missing cases are unit tests rather than fixture-dependent integration tests.
15. As an XLibur maintainer, I want the data bar's two halves joined by a type rather than a string
    match, so that the join cannot silently fail.
16. As an XLibur maintainer, I want new data bar attributes to be a field on one type, so that adding one
    is not an edit to a reader, two converters and a public interface.
17. As a contributing agent, I want the arity and alignment rules discoverable from the types, so that I
    do not have to infer them from converter code.

## Implementation Decisions

**The seam is a threshold value object, ordered.** It replaces the four collections. Its construction
enforces that every threshold has a type — supplying the format's default where the file omits one —
which is where the misalignment defect is designed out.

**Arity belongs to the format type.** A colour scale takes two or three thresholds; an icon set takes
as many as its icon count; a data bar takes two. These rules currently live implicitly in each
converter's indexing. They move onto the concept and are validated once.

**This is a breaking change to the public conditional format interface, and it is the point of the
spec.** It needs a deprecation path: the four collections remain, marked obsolete, projecting onto the
new collection for at least one release, so that consumers have a mechanical migration rather than a
compile break. `PublicAPI` files change; the change is called out in the release notes.

**The data bar's two halves become one module with two spellings.** This mirrors spec 44's two-adapter
seam. The string match that currently joins them is removed — a join that can fail silently is worse
than one the type system checks.

**Spec 48's behavioural fixes are preserved, not redone.** If 48 has landed, this spec must not
regress its fixtures; they are the acceptance evidence that the reshape is behaviour-preserving.

**Ordering.** The value object and its construction rules land first, with the four collections
projecting onto it and every existing test still passing. Then the reader and converters move to it.
Then the data bar unification. The obsolete collections are removed in a later release, not here.

## Testing Decisions

**What makes a good test here.** A good test exercises the public interface and asserts observable
results. The threshold value object is small enough to have direct unit tests for its construction
rules, and those are legitimate — it is a value type with invariants, not an extracted helper.

**The centrepiece is behaviour preservation.** Every existing conditional format test, and every
fixture from spec 48, must pass unchanged through the reshape. That is the acceptance criterion for the
projection stage; if a test needs changing, the reshape has altered behaviour and the change must be
justified in the results.

**Construction rule tests.** A threshold with no type takes the default. A threshold count wrong for the
format type is rejected. Alignment is not expressible as a mistake. These are unit tests on the value
object.

**The misalignment regression test.** An Excel-authored fixture whose value object omits a type, loaded,
saved and reloaded, asserting each threshold keeps its meaning. This test can only pass after the
reshape, which makes it the spec's proof.

**Migration tests.** The obsolete collections continue to return what they did, projected from the new
one, for every conditional format type.

**A data bar unification test.** Setting main-namespace and extension settings through one object and
asserting both reach the file.

**Prior art.** The conditional format converter dispatch table is the area's clean seam — one
dictionary, one interface, seventeen adapters — and is worth reading as the model for what good looks
like here. Spec 44's two-adapter design is the model for the data bar.

**Test seam.** `IXLConditionalFormat` through a save/reload round trip, plus direct unit tests on the
value object. The value object is a new seam, justified because it carries invariants.

## Out of Scope

- The three behavioural defects in spec 48 — the load crash, the save throw and the stripped data bar
  attributes. They ship first, without an API change.
- Removing the obsolete collections, which happens a release later.
- Differential format decoding — spec 28.
- Adding conditional format types or icon sets.
- The pivot conditional format priority handshake.

## Further Notes

Four parallel collections with a positional relationship is a shape worth naming, because it recurs: it
is what a record type looks like when it has been flattened into its fields and the fields have been
stored separately. Every operation then has to re-establish the relationship, and every operation is an
opportunity to fail to.

The give-away in this instance is that the reader appends to three collections unconditionally and to
the fourth conditionally. Nobody would write that if the four values were one object; it is only
writable because the collections are independent, and it is only wrong because they are not.

Splitting this from spec 48 was the right call: the crash is worth fixing this week, and a breaking
change to a public interface deserves its own decision on its own timetable.
