# Spec 45 — The text codec applies at the seam, not per caller

**Area:** Architecture · **Defect (crash on save, silent corruption in Excel)**
**Effort:** S (~2 days)
**Dependencies:** None. Touches no file any other open spec touches.
**Status:** Proposed. From the 2026-08-30 architecture review (round 3). The cheapest real win in the
round.

## Problem Statement

XML cannot carry certain characters. The spreadsheet format works around this with an escape
convention: a character XML forbids is written as an underscore-delimited hexadecimal code, and a
literal occurrence of that pattern in a user's text is itself escaped so it survives the round trip.

XLibur has a correct, unit-tested codec for this. It is applied at five of the seven places that need
it. The two that miss it are both ends of the *inline string* path — the way cell text is written when
string sharing is turned off, and the way it is read back.

What a user sees:

- **A save that throws.** Put a control character in a cell — the kind that arrives routinely in data
  scraped from other systems — and save with string sharing off, or through the streaming writer with
  inline storage. The XML writer rejects the character and the save fails. The same text saves without
  complaint when sharing is on, so whether the library works depends on a setting unrelated to the
  data.
- **Silent corruption, visible only in Excel.** A cell containing the literal text `_x0041_` is
  written unescaped. Excel decodes it on open and shows `A`. XLibur skips the decode too, so its own
  round trip looks clean and no save-and-reload test can see it. The corruption appears only in the
  application the file is written for.

Both switches that reach the unencoded path are public.

## Solution

The codec applies where the text element is written and where it is read, rather than at each caller.
A caller writing cell text cannot skip it, because the caller no longer chooses.

The five sites that currently encode correctly stop doing it themselves, so the encoding happens
exactly once and cannot be applied twice.

## User Stories

1. As a library consumer, I want a cell containing a control character to save successfully with string
   sharing off, so that my data does not dictate a configuration setting.
2. As a library consumer, I want the same through the streaming writer with inline string storage, so
   that the streaming path is as robust as the DOM path.
3. As a library consumer, I want a cell containing a control character to reload as the same text, so
   that the round trip is lossless.
4. As a library consumer, I want a cell containing the literal text `_x0041_` to appear in Excel as
   `_x0041_`, so that my text is not silently rewritten.
5. As a library consumer, I want the same for any literal text that resembles an escape sequence, so
   that the rule has no exceptions.
6. As a library consumer, I want text to round-trip identically whether it is stored as a shared string
   or inline, so that the storage mode is an implementation detail.
7. As a library consumer, I want text to round-trip identically through the DOM writer and the
   streaming writer, so that the write path is also an implementation detail.
8. As a library consumer, I want rich text runs to keep their current correct behaviour, so that this
   change does not disturb what already works.
9. As a library consumer, I want leading and trailing spaces to keep being preserved, so that the
   existing whitespace handling is unaffected.
10. As a library consumer, I want a worksheet name or a defined name containing awkward characters to
    behave as it does today, so that the change is scoped to cell text.
11. As an XLibur maintainer, I want the encode and decode to be unskippable, so that a future write path
    inherits them.
12. As an XLibur maintainer, I want the encoding applied exactly once, so that moving it to the seam
    cannot double-escape text at the sites that already encode.
13. As an XLibur maintainer, I want the character-checking setting on the XML writers to be deliberate
    rather than defaulted, so that the crash cannot come back through a different route.
14. As an XLibur maintainer, I want a round-trip test across both write paths and both storage modes, so
    that all four combinations are covered rather than the two that happen to work.
15. As a contributing agent, I want the codec's application to be a property of the element, so that I
    do not have to know which callers need it.

## Implementation Decisions

**The seam is the text element itself.** The helper that writes the text element applies the encode;
the code that reads it applies the decode. This is the highest available seam — every path that
produces cell text goes through that helper — and it already exists, so nothing new is introduced.

**Callers stop encoding.** The three write sites that currently encode correctly must be changed in the
same commit as the seam, or text will be escaped twice. This is the only real risk in the spec and is
why it is one commit rather than several.

**The character-checking setting becomes deliberate.** One reader in the library already disables it
explicitly, which shows the hazard was understood in one place and not propagated. After this spec the
encode makes the setting irrelevant for cell text, but it is set consciously on both sheet writers
rather than left at its default, so that a future change cannot silently reintroduce the crash.

**Rich text and shared strings are unaffected in behaviour.** They already encode and decode; they
simply stop doing it themselves.

**No public API change.**

## Testing Decisions

**What makes a good test here.** A good test writes a string through the public object model, saves,
reloads, and asserts the string came back unchanged. For the corruption case, an assertion on the
emitted bytes is also required, because XLibur's own round trip is clean — the encode and the decode
are both missing, so the error cancels out and only Excel sees it.

**The centrepiece is a four-combination matrix.** Two write paths — DOM and streaming — by two storage
modes — shared and inline. For each combination, round-trip a set of awkward strings: a control
character, a literal escape-sequence lookalike, text with leading and trailing spaces, a newline, and
ordinary text as a control. Today only the two shared-string combinations are exercised.

**A byte-level assertion for the lookalike case.** Assert that the literal escape sequence appears
escaped in the emitted XML. This is the only way to catch a defect where both halves of a codec are
missing.

**A double-encoding guard.** Round-trip a string containing an already-escaped-looking sequence
through the shared-string path, which previously encoded at the caller, and assert it is not escaped
twice. This is the regression the seam move could introduce.

**Prior art.** The codec has its own unit tests, which are correct and stay — what is missing is any
test asserting that the paths which need it actually call it. The existing streaming write tests
validate their output against the SDK schema validator, which passes here because the problem is not a
schema violation.

**Test seam.** `IXLCell.Value` with `ShareString`, the streaming worksheet's string storage option, and
a save/reload round trip plus emitted-byte assertions. No new seam.

## Out of Scope

- Text handling outside cell values — sheet names, defined names, and part names have their own rules.
- Rich text structure. Only its text content is in scope, and its behaviour does not change.
- The codec's own escaping rules, which are correct and tested.
- Choosing when to use shared versus inline storage.

## Further Notes

This is the smallest spec in the round and one of the two most severe, which is an unusual pairing. The
severity comes from the failure modes rather than the size: one is an unavoidable crash on data a user
did not choose to have, and the other is invisible to every test the library could write about itself.

The self-cancelling nature of the second defect is worth dwelling on. A missing encode paired with a
missing decode produces a system that is perfectly consistent with itself and wrong about the outside
world. No amount of round-trip testing finds it. The only tests that can are ones that assert on the
bytes, or ones that use a fixture the library did not write — which is the same lesson specs 41 and 46
arrive at from different directions.

The five encoding and two decoding sites were enumerated directly, and the unencoded write site was
verified while writing this spec.
