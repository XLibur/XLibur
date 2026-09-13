# Spec 52 — The fuzz harness gets an oracle worth the name

**Area:** Test infrastructure · **Tooling**
**Effort:** M (~3–4 days)
**Dependencies:** None. File-disjoint from every spec in the 41–51 queue except the library-side
exception change, which touches `XLWorkbook_Load.cs` and `Excel/IO/PartStructureException.cs`.
**Status:** ✅ **Merged**, 2026-08-31 — [#425](https://github.com/XLibur/XLibur/pull/425) (harness,
`42fe03c3`) then [#426](https://github.com/XLibur/XLibur/pull/426) (the defects it found, **breaking
`!`**, `08493c3f`), in that order, which was forced: #426 was stacked on #425. Worktree `xl-wt-52`
was based on `6570ed57`; it and its branches are removed. Branch `task/52` held the original
interleaved history; `task/52-harness` and `task/52-fixes` were the split that was pushed, verified
content-identical to it.

**Second loop session, 2026-09-01.** D37–D41, all on `task/52-fixes`. D37, D39, D40 and D41 are
fixed; **D38 is open and deliberately so**, and it now gates the `formula` target — see below.

| Defect | Target | Phase | State |
|---|---|---|---|
| D37 — every public `Evaluate` throws an internal exception type | `formula` | — | Fixed. Needed a second pass: the first fix wrapped only the evaluation, and the fuzzer found the reduction step seven minutes later |
| D38 — a range operand is not intersected: wrong answer, 600 ms/column | `formula` | — | **Open**, with a full [investigation](D38-investigation-implicit-intersection-on-operands.md). Naive fix tried, measured, rejected on six failing tests. The perf half is separable and behaviour-preserving |
| D39 — a date format makes an unwritable date out of any number | `workbook-structured` | **save** | Fixed. Only visible because `ArgumentException` is not tolerated during save |
| D40 — a sheet named `Ann''s` loads empty | `workbook-structured` | load, then **save** | Fixed at three sites. The third was unreachable until the first two were fixed |
| D41 — `IsValidRow` accepts `" 1"`, `"+1"`, `"1\0"` | `address` | — | Fixed, in two rounds. `NumberStyles.None` was not enough |

**Where the loop stands.** `workbook-structured` ran 147,236 executions clean and `address`
15,061,531 clean, after their fixes. `formula` cannot get a clean run until D38 is fixed: the shape
is trivially reachable by mutation and costs minutes per input. **The next useful move on the formula
target is fixing D38, not more fuzzing.**

**What this session added to the harness**, all on `task/52-fixes` rather than `task/52-harness`,
because each change only makes sense beside the library change it accompanies:
- `NotImplementedException` is inventoried by message rather than reported as a crash — a deliberate
  gap is not a Finding, but swallowing it would let every run spend its budget rediscovering the same
  ones. Three are recorded so far: the range-intersection operator, Excel 2016 vs E2019+ `@`
  intersection, and Bang references.
- A **load**-phase finding now dumps the package, not only a reload failure. The structured target
  builds its bytes from fuzzer input, so an artifact alone cannot be opened; D40 was diagnosed in
  minutes from the dumped `workbook.xml` and was guesswork before it.
- Corpus seeds for every finding except D38's, which is withheld on purpose.

> **Merge #425 first**, then rebase #426 with `git rebase --onto upstream/main <old-base> task/52-fixes`
> — squash-merge, so a plain `git rebase upstream/main` replays work already on `main`.

> **Read this before starting.** Two of the premises this work was built on were disproven during
> triage, *before* a line was written. See [Premises already disproven](#premises-already-disproven).
> The harness has existed since 2026-08-25 and has found exactly one defect in that time; the reasons
> are both recorded below and neither of them is "XLibur has no bugs".

## Problem Statement

`XLibur.Fuzz` drives three libFuzzer targets through SharpFuzz: `workbook` (load an `.xlsx`, save it
again), `formula` (`XLWorkbook.Evaluate(string)`), and `address` (three `XLHelper.IsValid*`
predicates). `fuzz.ps1` publishes the harness, instruments `XLibur.dll`, seeds a corpus and runs
libFuzzer over it.

Neither the harness nor the script is tracked in git. Both are untracked working-tree files; the
corpus and crash artifacts live under `temp/`, which is gitignored. Nobody but their author can run
this, reproduce a finding, or review the rules by which it decides something is wrong.

Three things are wrong with it beyond that.

**The oracle is a flat exception allowlist that ignores which phase threw.** `IsExpectedWorkbookFailure`
permits `ArgumentException`, `InvalidDataException`, `IOException`, `XmlException`,
`OpenXmlPackageException`, `PartStructureException`, `FormatException` and `OverflowException` — and
applies that list to the whole `load` + `save` sequence. An `OverflowException` thrown out of `SaveAs`
on a workbook that *loaded cleanly* is therefore discarded as though it were a malformed-input
rejection. It is not. It is a defect in the write path, and it is precisely the class the 28,000-test
suite cannot see.

**The oracle matches on exception message text, and the match has been wrong the whole time.** The
allowlist tests `Message.Contains("does not exist in the package")`. The message OpenXml actually
produces is `Part: /xl/workbook.xml doesn't exist in the package.` — *doesn't*, not *does not*. The
branch never fired. A case the harness was explicitly written to ignore has been landing in
`temp/fuzz/artifacts/` as a crash since the day it was written.

**Two targets can only detect a hard crash.** `address` calls three predicates and discards all three
results with `_ =`. A predicate that returns the wrong `bool` is invisible to it. `formula` evaluates
against a freshly constructed empty `XLWorkbook`, so every cell reference resolves to blank and the
coercion, array-argument and error-propagation paths are largely unreachable.

## Premises already disproven

Recorded here because the next reader needs them before starting, not after.

**"Mutating an `.xlsx` corpus exercises XLibur's readers."** It mostly does not. An `.xlsx` is a ZIP
with per-entry CRCs; a flipped byte in a compressed stream is rejected by `System.IO.Packaging` on
checksum before any XLibur reader is reached. The evidence is the corpus itself after a week of runs:

```
23 entries of exactly 5221 bytes   <- the seed's length, in-place byte substitution
 8 entries of 0..478 bytes         <- truncations
 0 entries that are a different valid package
```

libFuzzer never learned to construct a structurally different workbook, because almost every mutation
dies at the ZIP layer and returns no new coverage to guide it.

**"The allowlist describes what the harness ignores."** It described what its author intended to
ignore. One entry never matched (above), so the loop tripped on the same input within seconds of every
run, and the corpus never diversified. Six crash artifacts accumulated across three separate days and
all six are two bugs.

**"The harness runs."** It did not, and had not since it was written. Under libFuzzer the process
died during module initialisation, before a single input was fuzzed:

```
TypeInitializationException -> NullReferenceException
  at XLibur.Excel.LoadOptions.get_DefaultFontEngine()
  at SixLaborsV1FontBootstrap.Register()
  at ModuleInit.Initialize()
```

`XLibur.Fonts.SixLabors.V1` carries a `[ModuleInitializer]`; registering reads
`LoadOptions.DefaultFontEngine`, which is inside the assembly SharpFuzz has rewritten, and the
rewritten code needs a trace buffer that `Fuzzer.LibFuzzer.Run` is what allocates. **A reference is
enough — the branch need never execute**, so a `Register()` call in a `Main` branch that never runs
still kills the process, because JIT-compiling `Main` loads the assembly. Hence `NoInlining` on the
methods that touch it.

**The failure is near-silent, which is why it survived.** libFuzzer sees only an exit code, then
waits indefinitely for a target that is already gone, **ignoring its own `-max_total_time`**. One
600-second run sat for 25 minutes with no output, no corpus growth, and no crash artifacts, looking
from the outside exactly like a slow run. **Replay cannot catch it**, because replay runs an
uninstrumented build — so every green result obtained through replay says nothing about whether
fuzzing works. `fuzz.ps1` now runs libFuzzer under a watchdog for this reason.

## Solution

### The oracle splits by phase

Loading is where garbage is legitimately rejected; saving is not. A workbook that loaded is one XLibur
has claimed to understand, so the write path gets a nearly empty allowlist.

| Phase | Permitted | Rationale |
|---|---|---|
| `new XLWorkbook(stream)` | a narrow named list of rejection types | garbage in, deliberate rejection out |
| `SaveAs(stream)` | `IOException`, `OutOfMemoryException` only | anything else is a write-path defect |
| re-load of what was just saved | nothing | see below |

`OutOfMemoryException` is permitted but **recorded** rather than silently dropped: an input under
1 MB that exhausts memory in the writer is worth knowing about even though it is not a crash.

**No allowlist entry may match on exception message text.** This rule is not a style preference; it is
the direct lesson of the defect above. Where a rejection cannot be identified by type, the fix is to
give it a type — see the next section.

### The library stops leaking its dependency's exceptions

The reason the allowlist reached for a message match is that OpenXml signals "this package has no
workbook part" with a bare `System.InvalidOperationException`, a type XLibur itself throws for genuine
bugs. Allowing it wholesale during load would blind the harness; matching its message is banned.

So the fix belongs in the library, not the oracle: `XLWorkbook.Load(Stream)` catches the OpenXml
exceptions that escape it and rethrows them as `PartStructureException`. A caller of
`new XLWorkbook(stream)` should not have to catch a `DocumentFormat.OpenXml` type either, so the
harness's requirement and the public contract turn out to be the same requirement.

`PartStructureException` is widened to carry it. Its name already promises package structure; only its
doc comment was narrow, claiming the type is "thrown from parser when there is a problem with data in
XML". A missing package part is not malformed XML — there is no XML to be malformed. The doc comment
is corrected and a `MissingPart(string partName)` factory added, so the message names the part.

### Round-trip: XLibur must be able to read what XLibur wrote

After a successful save, the harness loads the produced bytes again. Nothing is permitted to throw.
This is a genuine invariant and costs one constructor call. It catches silent corruption in the write
path, which no exception-based oracle ever will — the failure mode behind D4 and D10.

Full byte-stability comparison (save → load → save → compare) is deliberately **not** attempted; this
codebase has known-legitimate sources of byte variance and it would drown iteration one in false
findings.

### A structure-aware workbook target

Added alongside the blind mutation target, not replacing it — blind mutation found D27 precisely
because it produces structurally broken packages.

The structure-aware target decodes the fuzz bytes into a *generator* that always emits a valid ZIP
with a valid `[Content_Types].xml`, so every input reaches XLibur's readers. What the bytes vary:

- sheet count, names (including duplicates), cell references, values and types
- relationship IDs, including dangling ones
- style indices, including out-of-range
- `dimension` refs that disagree with the actual cell extents
- shared-string indices past the end of the table

That band is chosen because it is where the open defect register already clusters — internally
inconsistent but well-formed XML. Pivot caches, conditional formatting and data validation are
deliberately out of scope for this spec; each needs its own generator and should be earned with
results from this one.

### The other two targets get oracles

`address` gains a round-trip oracle: where an address parses, render it back to a string, re-parse,
and require the two to agree. Plus a consistency check across the three predicates. A predicate
returning the wrong `bool` becomes visible.

`formula` evaluates against a small fixed populated sheet — numbers, text, dates, booleans, an error
value, a blank, and a small range — so references resolve to real values of every type and the
coercion surface is reachable.

### The tool becomes reproducible

- A seed corpus is committed at `XLibur.Fuzz/corpus/<target>/`, replacing `fuzz.ps1`'s inline
  `Copy-Item`/`WriteAllText` seeding. A fresh clone starts where the author started.
- `fuzz.ps1` gains `-Replay <path>`: run the published harness over a file or directory and print
  phase, exception type and stack per input. Triage is the step between Finding and Defect and it will
  happen dozens of times; it belongs in the tool rather than in a throwaway console project.

## Vocabulary

This spec uses **Rejection**, **Finding** and **Defect** in the senses fixed in
[`CONTEXT.md`](../CONTEXT.md). The distinction that matters most here: a defect in the *harness* is
not a Defect in the register and does not take a `D` number. The message-match bug above is a harness
defect, recorded in this spec's Results and nowhere else.

## Acceptance criteria

1. `XLibur.Fuzz/` and `fuzz.ps1` are tracked, and the harness builds under
   `TreatWarningsAsErrors=true` with nullable enabled.
2. No allowlist entry anywhere in the harness inspects `Exception.Message`. Verifiable by grep.
3. The oracle distinguishes load, save and re-load phases, and the save phase permits at most
   `IOException` and `OutOfMemoryException`.
4. `new XLWorkbook(stream)` throws no `DocumentFormat.OpenXml` exception type for any input in the
   committed seed corpus or in `temp/fuzz/artifacts/`.
5. A regression test covers a package with a valid `[Content_Types].xml` and no workbook part, and it
   asserts the thrown type — the package is **synthesised in the test**, not committed as a binary
   artifact.
6. The structure-aware target produces, in a 600-second run, a corpus containing at least one input
   that is a *different valid package* from any seed. This is the criterion the old target could never
   have met; it is the point of the change.
7. `fuzz.ps1 -Replay <dir>` reproduces the six artifacts in `temp/fuzz/artifacts/` and reports two
   distinct failures, not six.
8. `dotnet test XLibur.Tests/XLibur.Tests.csproj` is green on both `net8.0` and `net10.0`.

**Criterion 6 is the one at risk** and should be reported rather than worked around if it proves
unreachable. Recent rounds have twice produced acceptance criteria that were arithmetically
impossible (spec 26 criterion 8, spec 28 criterion 6); the correct response is to say so.

## Out of scope, deliberately

- **A model-mutation target** — fuzz bytes decoding to a sequence of API operations (insert row,
  delete column, clear range, merge) on a loaded workbook, then save and reload. This is the
  highest-value target on the list, aimed straight at the shifter and sheet-listener code where D15,
  D16, D17 and D26 all live. It is deferred because its oracle is genuinely hard — "did insert-row do
  the right thing?" has no cheap answer — and doing it badly now would burn the idea.
- **Wide generator scope** — pivot caches, conditional formatting, data validation.
- **Splitting `formula` into parser and evaluator targets.** The stack trace already identifies the
  layer.
- **Byte-stability round-trip comparison.** See above.

## Notes for whoever runs the loop

- **This spec has no completion state.** The harness completes; the fuzzing does not. Do not give it a
  status cell that can flip to ✅ — that is exactly the lie §3.6 of the conductor `CLAUDE.md` warns
  about. Findings are recorded as ordinary `DEFECTS.md` entries from D27 onward.
- **Expect the generator grammar to flood.** Dangling relationship IDs and out-of-range style indices
  are the inputs most likely to produce a hundred instances of one Finding. If that happens, say so and
  dial the generator back deliberately — do not quietly narrow the grammar and report the run as clean.
- **One Defect per PR.** The harness lands first as its own PR; each Defect after it is a `fix:` with
  its own regression test and CHANGELOG line.
