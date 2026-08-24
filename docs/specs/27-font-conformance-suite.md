# Spec 27 — One font conformance module, run by every adapter

**Area:** Testing · Architecture · **Correctness (untested seam)**
**Effort:** M (~4–5 days)
**Dependencies:** None. Almost entirely test-project work — no production file is modified.
**Status:** Proposed. **Spec 34 hard-depends on this one** — see Conflicts.

## Goal

Make the three `IXLFontEngine` adapters assert the same numbers. Today each is tested alone,
against monotonicity checks that all three would pass while disagreeing by 20%, and nothing
anywhere compares two of them.

## Why this spec exists

`IXLFontEngine` is a five-method port with three implementations:

| Adapter | File | Lines | Font library |
|---|---|---:|---|
| SixLabors v1 (`DefaultFontEngine`) | `XLibur.Fonts.SixLabors.V1/DefaultFontEngine.cs` | 306 | SixLabors.Fonts 1.0.1 |
| SixLabors v2 (`SixLaborsFontEngine`) | `XLibur.Fonts.SixLabors/SixLaborsFontEngine.cs` | 277 | SixLabors.Fonts 2.1.3 |
| SkiaSharp (`SkiaSharpFontEngine`) | `XLibur.Fonts.SkiaSharp/SkiaSharpFontEngine.cs` | 314 | SkiaSharp 4.151.1 |

Three implementations exist because of one constraint: **SixLabors.Fonts 2.x changed to the Six
Labors Split License and the repo is pinned to 1.0.1 (Apache 2.0).** That is why the port was cut
(`docs/font-architecture.md:7`), why a v2 package ships separately, and why an MIT-licensed
SkiaSharp adapter was added and made the default. The port is a license artefact, and the price of
a license artefact is that three implementations now have to agree by hand.

### The two adapter test suites are near-identical copies

`XLibur.Fonts.SixLabors.Tests/SixLaborsFontEngineTests.cs` is 431 lines with 31 `[Test]`.
`XLibur.Fonts.SkiaSharp.Tests/SkiaSharpFontEngineTests.cs` is 434 lines with 31 `[Test]`.

Normalising the engine name away and diffing them:

```bash
sed 's/SixLabors/ENGINE/g;s/SkiaSharp/ENGINE/g' XLibur.Fonts.SixLabors.Tests/SixLaborsFontEngineTests.cs > /tmp/a.cs
sed 's/SixLabors/ENGINE/g;s/SkiaSharp/ENGINE/g' XLibur.Fonts.SkiaSharp.Tests/SkiaSharpFontEngineTests.cs > /tmp/b.cs
diff /tmp/a.cs /tmp/b.cs | grep -c '^[<>]'
  -> 13
```

**13 differing lines out of 865.** Re-measured today against `1b41cadd`; the number is exact.
(The `sed` here reads two tracked files and writes to `/tmp` — it does not use `-i` and does not
rewrite anything tracked.)

### What the 13 lines are is worse than the fact that there are only 13

Nine of the thirteen are the two assertions that carried a number. They were weakened, not ported:

| Metric | `XLibur.Tests/Graphics/FontTests.cs` (V1) | SixLabors v2 suite | SkiaSharp suite |
|---|---|---|---|
| TestFontA `"A"` @ 20pt, 120 DPI | `IsEqualTo(31.25d).Within(0.0001)` (`:91`) | `IsEqualTo(31.25d).Within(1.0)` (`:251`) | `IsGreaterThan(0)` (`:251`) |
| TestFontB `"B"` @ 30pt, 96 DPI | `IsEqualTo(25d).Within(0.0001)` (`:105`) | `IsEqualTo(25d).Within(1.5)` (`:263`) | `IsGreaterThan(0)` (`:265`) |

The tolerance went from 0.0004% to 3.2% at v2, then to nothing at SkiaSharp. The v2 comment says
why: *"v2 may have slightly different measurement than v1"* — a guess, widened to cover it. The
SkiaSharp copy dropped the assertion entirely and replaced it with a positivity check plus a
`IsNotEqualTo(widthFallback)`.

Mechanically:

```bash
grep -c "Within(" XLibur.Fonts.SixLabors.Tests/SixLaborsFontEngineTests.cs   # -> 2
grep -c "Within(" XLibur.Fonts.SkiaSharp.Tests/SkiaSharpFontEngineTests.cs   # -> 0
```

**The shipped default engine has zero numeric assertions on any of its five metrics.** What it has
instead is 17 `IsGreaterThan(0)` calls
(`grep -c "IsGreaterThan(0)" XLibur.Fonts.SkiaSharp.Tests/SkiaSharpFontEngineTests.cs`).

### Nothing anywhere compares two adapters

```bash
grep -rl "SkiaSharpFontEngine" --include=*.cs . | xargs grep -l "SixLaborsFontEngine"
  -> (no output)
```

Verified. No file in the repository names two font engines. The remaining 29 tests in each suite
are monotonicity and positivity checks — `GetTextWidth_LongerTextIsWider`,
`GetTextWidth_LargerFontIsWider`, `GetDescent_ReturnsPositiveValue`,
`GetGlyphBox_DescentIsPositive`. All three adapters pass all of them while producing different
numbers, because "longer text is wider" is true of any correct font engine and of several
incorrect ones.

CI runs the four suites as four separate steps
(`.github/workflows/build-and-test.yml:101`, `:110`, `:119`, `:127`). None of them crosses an
adapter boundary.

### The shipped default never runs the core suite

`XLibur.Tests/XLibur.Tests.csproj:41` references `XLibur.Fonts.SixLabors.V1` and no other font
package. `XLibur.Tests/TestInfrastructure.cs:31` calls `SixLaborsV1FontBootstrap.Register()` in the
assembly hook. So the whole autofit / row-height suite — 45 `AdjustToContents` call sites across 9
files — runs against V1.

`XLibur/Graphics/DefaultFontEngineProbe.cs:30` makes `XLibur.Fonts.SkiaSharp` the engine a
zero-config consumer actually gets. `XLibur.Report.Tests` already references SkiaSharp
(`XLibur.Report.Tests.csproj:29`), so the two test suites in this repo run under two different font
engines, and the one the library ships by default is the one the core suite never exercises.

### The 0% drift claim rests on a spike, not on CI

`docs/font-architecture.md:63`:

> A measurement spike comparing the two libraries on Carlito (Calibri-compatible), TestFontA, and
> TestFontB found **0% metric drift** across width, descent, height, and max-digit-width

Read and checked. Nothing in the tree defends it. The doc's own Test Strategy table (`:148-154`)
says as much — "1.0.1 only", "2.1.3 only", "SkiaSharp 3.x" — and `:154` states the test projects
never reference two font packages. The claim is a one-off measurement with no regression gate. If
the number is still 0% today, task 1 says so and pins it. If it is not, task 1 is the first thing
in the repo that would notice.

(Minor doc drift found while checking: `docs/font-architecture.md:152` says SkiaSharp 3.x;
`XLibur.Fonts.SkiaSharp.csproj:47` pins 4.151.1.)

### The precedent already exists, one seam over

`XLibur.Report.Tests/Expressions/BothEnginesTests.cs` — 342 lines, 13 `[Test]` — generates the same
report through the Scriban engine and the DynamicLinq engine and asserts the two workbooks agree
cell for cell over the used range (`:84-91`). Its class remarks state the design goal exactly:

> it says the engine seam is where the whole of the difference lives, which is the property that
> lets the package plug in and out

That is the property `IXLFontEngine` claims and does not test. This spec applies the same idea to
the font port, with one structural difference forced on it by the license constraint — see The
design.

### Why this matters outside the test projects

`GetMaxDigitWidth` and `GetGlyphBox` feed column autofit
(`XLibur/Excel/Columns/XLColumn.cs:195`, `:250`, `:253`) and row autofit
(`XLibur/Excel/Rows/XLRow.cs:331`). `GetMaxDigitWidth` is also the divisor that converts pixels to
the OOXML "number of characters" column-width unit at `XLibur/XLHelper.cs:489` and `:518`, and on
the load path at `XLibur/Excel/XLWorkbook_Load.cs:948`. `GetTextHeight` and `GetTextWidth` size
comment boxes at `XLibur/Excel/IO/VmlDrawingPartWriter.cs:279-280`.

A metric that differs between adapters therefore differs in the `<col width="...">` attribute of a
saved file. That is a user-visible output difference produced by a package choice, and nothing
would catch it.

## Non-goals

- **Not changing `IXLFontEngine`.** That is spec 34, being written in parallel. This spec is its
  prerequisite gate: 34 moves metric computation across all three adapters and has nothing to
  prove it did not change any number until this suite exists.
- **Not upgrading SixLabors.Fonts.** The 1.0.1 pin is a license constraint
  (`docs/font-architecture.md:7`) and is the reason there are three adapters at all.
- **Not a performance spec.** The conformance suite runs on stream-loaded fonts and adds a few
  hundred milliseconds to three test projects. No benchmark is involved.
- **Not reconciling a divergence this spec finds.** Task 1 records it; whether it is fixed here or
  in 34 is decided once the number is known.
- **No production-code change.** The only non-test file touched is `XLibur.Tests.csproj` and its
  two adapter siblings, plus the CI workflow.

## Current state

Verified against the tree at `1b41cadd` (2026-08-24).

- **The port** — `XLibur/Graphics/IXLFontEngine.cs:10-55`. Five methods: `GetTextHeight:15`,
  `GetTextWidth:21`, `GetMaxDigitWidth:27`, `GetDescent:33`, `GetGlyphBox:54`.
- **Core reaches it from six places**, not three as first surveyed:
  - `XLibur/Excel/Columns/XLColumn.cs:190` — acquires the engine for `AdjustToContents`
  - `XLibur/Excel/Rows/XLRow.cs:294` — same for row height
  - `XLibur/Excel/IO/VmlDrawingPartWriter.cs:265` — comment autosize
  - `XLibur/XLHelper.cs:489`, `:518` — `GetMaxDigitWidth` for the NoC↔pixel conversion
    (**correction:** the file is `XLibur/XLHelper.cs`, not `XLibur/Utils/XLHelper.cs`)
  - `XLibur/Excel/XLWorkbook_Load.cs:948` — `GetMaxDigitWidth` on the load path
  - fanning out through `XLibur/Excel/Cells/XLCell.cs:1484` →
    `XLibur/Excel/Cells/XLCellGlyphHelper.cs:21`, `:65`, `:79` for `GetGlyphBox`
- **Adapters** — line counts corrected from the survey, which had the two SixLabors files swapped:
  V1 306, v2 277, SkiaSharp 314.
- **Adapter suites** — `SixLaborsFontEngineTests.cs` 431 lines / 31 `[Test]`;
  `SkiaSharpFontEngineTests.cs` 434 lines / 31 `[Test]`; 13 differing lines of 865 normalised.
- **`SkiaSharpFontBootstrapTests.cs`** — 142 lines, 6 `[Test]`, no counterpart in the other suite.
- **`XLibur.Tests/Graphics/FontTests.cs`** — 192 lines, 12 `[Test]`, V1 only. Carries the only
  tight numeric assertions in the repo: `7.43359375d` exact (`:38`, `:77`, `:127`), `31.25d`
  within `0.0001` (`:91`, `:140`), `25d` within `0.0001` (`:105`).
- **`DummyFont`** — four copies: `FontTests.cs:160`, `SixLaborsFontEngineTests.cs:410`,
  `SkiaSharpFontEngineTests.cs:413`, `SkiaSharpFontBootstrapTests.cs:121`.
- **Test fonts** — six files, three projects × two fonts, all byte-identical:
  `TestFontA.ttf` `21ca718e11aadbd392499ab5e5084bef`, `TestFontB.ttf`
  `40d7a669ff482fec490d9a964da9a68d`.
- **Registration is `??=`** — `SixLaborsV1FontBootstrap.Register()` and
  `SkiaSharpFontBootstrap.Register()` both do `LoadOptions.DefaultFontEngine ??= …`, and V1 also
  auto-registers from `XLibur.Fonts.SixLabors.V1/ModuleInit.cs`. First writer wins. Task 5 depends
  on knowing this.
- **CarlitoBare** — `XLibur.Fonts.SixLabors/` embeds none;
  `XLibur.Fonts.SixLabors.V1.csproj:24-27` and `XLibur.Fonts.SkiaSharp.csproj:30-33` each embed
  their own copy; `XLibur/Graphics/Fonts/*.ttf` are tracked and embedded by nothing
  (`XLibur.csproj:25-26` embeds two XML files and nothing else). All three copies are byte-identical
  (`f2f52428f1f3fe4d3cbca3bae238b775`). See the closing note.

## File structure

```
XLibur.Fonts.Conformance/FontEngineConformance.cs      new — the shared [Test] bodies
XLibur.Fonts.Conformance/GoldenMetrics.cs              new — expected values + per-metric tolerance
XLibur.Fonts.Conformance/ConformanceFixture.cs         new — how a suite supplies its engine/fonts
XLibur.Fonts.Conformance/ConformanceFont.cs            new — the one IXLFontBase stub

XLibur.Fonts.SixLabors.Tests/SixLaborsConformanceTests.cs        new — ~20 lines
XLibur.Fonts.SixLabors.Tests/SixLaborsFontEngineTests.cs         deleted (431 lines)
XLibur.Fonts.SixLabors.Tests/XLibur.Fonts.SixLabors.Tests.csproj modified — linked Compile items

XLibur.Fonts.SkiaSharp.Tests/SkiaSharpConformanceTests.cs        new — ~20 lines
XLibur.Fonts.SkiaSharp.Tests/SkiaSharpFontEngineTests.cs         deleted (434 lines)
XLibur.Fonts.SkiaSharp.Tests/XLibur.Fonts.SkiaSharp.Tests.csproj modified — linked Compile items

XLibur.Tests/Graphics/SixLaborsV1ConformanceTests.cs   new — ~20 lines
XLibur.Tests/Graphics/EngineDriftProbe.cs              new in task 1, deleted in task 4
XLibur.Tests/Graphics/FontTests.cs                     modified — V1-specific tests only
XLibur.Tests/TestInfrastructure.cs                     modified — engine selectable in task 5
XLibur.Tests/XLibur.Tests.csproj                       modified — linked Compile items, + SkiaSharp
.github/workflows/build-and-test.yml                   modified — second core pass under SkiaSharp
```

`XLibur.Fonts.Conformance/` holds `.cs` files and **no `.csproj`**. That is deliberate; see
decision 1.

## The design

### The constraint that shapes everything: two adapters can never share a process

`XLibur.Fonts.SixLabors.V1` depends on SixLabors.Fonts 1.0.1;
`XLibur.Fonts.SixLabors` depends on 2.1.3. A project referencing both gets NuGet's unification to
2.1.3, and the V1-compiled code silently runs against the V2 assembly — the exact failure
`docs/font-architecture.md:9` and `:162` were written to prevent, and the reason `:154` records
that no test project references both.

So a `BothEnginesTests`-shaped comparison is **impossible for the pair that matters most**. Two of
the three adapters can never be loaded into the same process, and no test can ever call both and
subtract.

Cross-adapter agreement therefore has to be established **through shared constants**, not through
shared execution. One module of tests, compiled into three separate test assemblies, asserting the
same golden numbers with the same tolerances. Agreement between adapter A and adapter B follows
from both agreeing with the table, transitively, without them ever meeting.

That is also why this spec is a suite and not a single comparison test: the single comparison test
is only available for the pairs {V1, SkiaSharp} and {v2, SkiaSharp} — SkiaSharp shares no library
with either SixLabors package, so it can be co-referenced with either. Task 1 uses that to
bootstrap the constants.

### Decision 1 — linked source, not a shared project

**Recommended: `<Compile Include="..\XLibur.Fonts.Conformance\*.cs" Link="Conformance\%(Filename)%(Extension)" />` in each of the three test projects.**

Reasons, in order of weight:

1. **TUnit discovery is guaranteed.** TUnit is source-generated: the generator emits registrations
   into the assembly it compiles. Tests declared in a referenced class library register through
   that library's module initializer, which fires only when the CLR first touches a type in it —
   the same lazy-load trap `docs/font-architecture.md:167-168` documents for the font bootstrap,
   and the reason `DefaultFontEngineProbe` exists at all. Linked source compiles the `[Test]`
   bodies into each test assembly, so the generator runs where the runner looks.
2. **Nothing can be packed.** A directory with no `.csproj` cannot appear in `XLibur.slnx`, cannot
   be picked up by the release workflow, and cannot acquire a MinVer version. A shared project
   would need `<IsPackable>false</IsPackable>` and a standing reason not to remove it.
3. **No transitive references.** A shared test project would have to reference `XLibur` core and
   TUnit, and both are already present in all three consumers. Adding an assembly buys nothing.

Cost: three `<Compile Include>` globs to keep in sync, and the file appears three times in an IDE
solution view. Accepted.

**If a 30-minute spike shows TUnit discovers inherited `[Test]` methods from a referenced assembly
reliably on both net8.0 and net10.0, the shared-project form is the better long-term shape and this
decision should be revisited.** It is not the shape to bet the spec on today.

### Decision 2 — the tolerance is the contract

Different rasterisers will not agree to the bit, so the golden table needs a stated tolerance per
metric. The tolerance is not slack for making tests pass; it is the strongest statement the spec is
willing to make about the port, and it is what spec 34 will be measured against.

The existing evidence sets the bar. Working back from the three known constants:

| Constant | Where | Font units |
|---|---|---|
| `31.25` px = TestFontA `"A"` @ 20pt, 120 DPI | `FontTests.cs:91` | `0.9375` em = **15/16** exactly |
| `25` px = TestFontB `"B"` @ 30pt, 96 DPI | `FontTests.cs:105` | `0.625` em = **5/8** exactly |
| `7.43359375` px = CarlitoBare max digit @ 11pt, 96 DPI | `FontTests.cs:38`, `:77` | **1038/2048** em exactly |

Every one is an exact rational fraction of the em square. That is not a coincidence — all three
adapters compute widths as `advance_fu × size / unitsPerEm`, so the pixel value is a rational
number and the only source of difference is floating-point width and rounding order. An adapter
that lands 3.2% away from `31.25` is not rasterising differently; it is reading a different number
out of the font. The v2 suite's `Within(1.0)` was never justified by anything measured.

**The tolerances this spec sets:**

| Metric | Tolerance | Why |
|---|---|---|
| `GetMaxDigitWidth`, `GetDescent`, `GetTextHeight` | **0.5% relative, or 0.01 px, whichever is larger** | Pure `fu × size / upem` arithmetic in all three adapters (`SixLaborsFontEngine.cs:126`, `:133`, `:140`; `SkiaSharpFontEngine.cs:128`, `:135`, `:142`). Anything above float rounding is a different input, not a different renderer. |
| `GetTextWidth` on a multi-glyph string | **1% relative** | The one genuine algorithmic difference: the SixLabors adapters pass `KerningMode.None` to `TextMeasurer.MeasureAdvance` (`SixLaborsFontEngine.cs:150`, `DefaultFontEngine.cs:164`), SkiaSharp calls `SKFont.MeasureText` (`SkiaSharpFontEngine.cs:149`) with the font's default shaping. |
| `GetTextWidth` on a single glyph | **0.5%** | No kerning pair exists, so it reduces to the arithmetic case. |
| `GlyphBox.EmSize`, `GlyphBox.Descent` | **exact** | Both are `Math.Round(..., MidpointRounding.AwayFromZero)` in every adapter (`SkiaSharpFontEngine.cs:167-168`) and are documented as whole numbers (`GlyphBox.cs:37-50`). They are integers. A tolerance here would hide an off-by-one, which is one pixel of row height in every saved file. |
| `GlyphBox.AdvanceWidth` | **0.5%** | Deliberately unrounded (`GlyphBox.cs:13-15`) so widths sum without accumulating error. |

**A metric the three cannot agree on within these is a finding, not a reason to widen the
tolerance.** Record it, and pin the divergence with a per-adapter expected value and a comment
naming the cause — the way `FormulaShifterCorpusTests` carries a separate `LegacyExpected` column
for the nine rows the two shifters disagree on (spec 25). Never widen a shared tolerance to cover
one adapter.

**Two premises here could be wrong, and task 1 is what disproves them:**

- **P1: the 1% kerning tolerance holds** because TestFontA and TestFontB have no `kern` or GPOS
  table. Unverified — the fonts were not inspected. If they do kern, `GetTextWidth` on
  `"Lorem ipsum dolor sit amet"` will diverge by more than 1% and the multi-glyph row needs either
  a kerning-free string or a pinned per-adapter value.
- **P2: `GetDescent` and `GetTextHeight` agree at all.** SixLabors 1.x/2.x read
  `FontMetrics.VerticalMetrics.Ascender/Descender`; SkiaSharp reads `SKFontMetrics.Ascent/Descent`
  (`SkiaSharpFontEngine.cs:182`). The two formulas are algebraically identical once the sign
  convention is normalised — `Ascender − 2·Descender` versus `AscentFu + 2·DescentFu` with
  `AscentFu = −metrics.Ascent` — but only if both libraries read the *same table*. `hhea` and OS/2
  `sTypoAscender` and OS/2 `usWinAscent` are three different numbers in most fonts, and
  `IXLFontEngine.cs:32` records that Excel uses the third one and that neither adapter does. If
  Skia and SixLabors pick differently, descent and height diverge by whatever the font's tables
  disagree by, which for Calibri-class fonts is 10–20%.

If P2 fails, that is a real defect in row heights and comment box sizes, not a test-authoring
problem, and it is recorded in Results whether or not it is fixed here.

### The shared module

```csharp
/// <summary>
/// What one adapter test project supplies so the shared conformance tests can run against it.
/// </summary>
/// <remarks>
/// The engine is created per fixture rather than shared, because two of the three adapters can
/// never be loaded into the same process (SixLabors 1.0.1 and 2.1.3 unify) and so this module is
/// compiled separately into each test assembly rather than referenced from one.
/// </remarks>
internal abstract class ConformanceFixture
{
    /// <summary>An engine loaded with <c>TestFontA</c> as fallback and <c>TestFontB</c> as an extra.</summary>
    internal abstract IXLFontEngine CreateStreamEngine();

    /// <summary>
    /// An engine whose ultimate fallback is the embedded CarlitoBare, or <c>null</c> when the
    /// adapter embeds none. <c>XLibur.Fonts.SixLabors</c> (v2) embeds no fonts, so the CarlitoBare
    /// rows of the golden table are not applicable to it.
    /// </summary>
    internal abstract IXLFontEngine? CreateEmbeddedFallbackEngine();

    /// <summary>Name of the adapter, used in assertion messages and in the drift report.</summary>
    internal abstract string AdapterName();
}
```

```csharp
/// <summary>
/// One row of the golden metric table: an expected value in pixels for a named font at a fixed
/// size and DPI, plus the tolerance the three adapters must meet.
/// </summary>
/// <remarks>
/// Expected values are exact rational fractions of the em square, because every adapter computes
/// <c>advance_fu * size / unitsPerEm</c>. See spec 27, decision 2, for why the tolerances are what
/// they are and why widening one is not an available response to a failure.
/// </remarks>
internal readonly record struct GoldenMetric(
    string Font,
    double SizePt,
    double Dpi,
    string? Text,
    double ExpectedPx,
    double TolerancePx,
    string EmFraction);
```

The 31 existing tests move into `FontEngineConformance` unchanged in intent: the monotonicity
checks stay (they are cheap and they catch a whole class of wiring bug), and each acquires a golden
row where one exists. `DummyFont` collapses from four copies to one `ConformanceFont`.

## Global constraints

- Warnings are errors (`TreatWarningsAsErrors=true` in `Directory.Build.props`); nullable is
  enabled repo-wide — new code must be null-annotated.
- Branch per spec; never commit to main. Commit prefixes `test:` for tasks 1, 3, 4, 5;
  `refactor:` for task 2.
- No compound shell commands (`&&`, `||`, `;`) in agent tool calls.
- **Do not use `sed -i` on tracked files.** `.gitattributes` checks out CRLF and Git Bash's
  `sed -i` rewrites the file as LF, turning a one-line change into a whole-file diff. Use the
  Edit/Write tools; verify with `git diff --numstat` — a file whose changed-line count approaches
  its total line count was rewritten, not edited. (The `sed` in this spec's evidence section writes
  to `/tmp` and never touches a tracked file.)
- **Do not upgrade SixLabors.Fonts.** Newer versions carry the Six Labors Split License. This
  constraint is why three adapters exist and is the whole premise of this spec.
- Tests: `--treenode-filter`, never `--filter`. Exit 5 = invalid runner option; exit 8 = zero tests
  matched. **Never filter at solution level** — name the `.csproj`.
- Pass `-f net10.0` while iterating; run without it before opening the PR, since every test project
  multi-targets `net8.0;net10.0`.
- Build: `dotnet build XLibur/XLibur.csproj -c Release -v q`
- TUnit: `await Assert.That(actual).IsEqualTo(expected)`. Assertions are awaitable and a missing
  `await` passes silently, which is why `XLibur.Tests.csproj:14` promotes CS4014 to an error. The
  two adapter test projects do not set that explicitly — `TreatWarningsAsErrors` covers it, but add
  the explicit `<WarningsAsErrors>$(WarningsAsErrors);CS4014</WarningsAsErrors>` line to both while
  editing them in task 2.

## Work plan

| # | Task | Size | Gate |
|---|---|---|---|
| 1 | Prove the gap: one test, two adapters, all five metrics, numbers recorded | S | New test green; drift table in Results |
| 2 | Extract the conformance module; delete the two copies | M | 31 tests run in both adapter suites; both copies gone |
| 3 | Golden metric table | M | Every metric asserted against a constant in every suite |
| 4 | V1 joins the conformance run | S | Three suites × the same module |
| 5 | Core suite, second pass under SkiaSharp | S | Full `XLibur.Tests` green with SkiaSharp registered |

**Task 1 sizes the risk in tasks 3–5.** If V1 and SkiaSharp already disagree beyond the tolerances
in decision 2, the golden table cannot be written as a single shared constant per metric, and this
spec changes shape. Do not start task 3 before task 1's numbers are recorded.

---

### Task 1 — Prove the gap

`XLibur.Tests` already references V1. SkiaSharp shares no library with SixLabors
(`XLibur.Fonts.SkiaSharp.csproj:41-49` — `XLibur` core, MinVer, SkiaSharp, SkiaSharp native assets,
and nothing else), so adding it creates no unification hazard. That reference is the one task 5
needs permanently, so add it now.

**Files:**
- Create: `XLibur.Tests/Graphics/EngineDriftProbe.cs`
- Modify: `XLibur.Tests/XLibur.Tests.csproj`

**Interfaces:**
- Produces: `V1_and_SkiaSharp_agree_on_every_metric`, and a printed drift table that becomes the
  golden table in task 3.

- [ ] **Step 1: Add the SkiaSharp project reference**

In `XLibur.Tests/XLibur.Tests.csproj`, inside the existing `ItemGroup` at `:39-43`:

```xml
    <ProjectReference Include="..\XLibur.Fonts.SkiaSharp\XLibur.Fonts.SkiaSharp.csproj" />
```

This is the only project in the repo that will reference V1 and SkiaSharp together. It is safe;
`XLibur.Report.Tests` already references SkiaSharp alone and `XLibur.Tests` already references V1
alone, and neither package depends on the other's font library.

Run: `git diff --numstat XLibur.Tests/XLibur.Tests.csproj`
Expected: `1  0  XLibur.Tests/XLibur.Tests.csproj`. Anything near 45 changed lines means the file
was rewritten with LF endings — revert and use the Edit tool.

- [ ] **Step 2: Write the probe**

```csharp
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Fonts.SixLabors.V1;
using XLibur.Fonts.SkiaSharp;
using XLibur.Graphics;

namespace XLibur.Tests.Graphics;

/// <summary>
/// Spec 27 task 1. Two IXLFontEngine adapters, the same stream-loaded fonts, all five metrics,
/// compared directly.
/// </summary>
/// <remarks>
/// <para>
/// This is the only adapter pair that can be compared inside one process. SixLabors.Fonts 1.0.1
/// and 2.1.3 unify under NuGet, so V1 and the v2 package can never be loaded together — which is
/// why spec 27's design establishes agreement through shared constants rather than through a
/// BothEnginesTests-style direct comparison.
/// </para>
/// <para>
/// docs/font-architecture.md:63 claims 0% metric drift between the engines, from a one-off spike.
/// Nothing in CI defends it. This is the first thing in the tree that would notice if it were
/// wrong.
/// </para>
/// </remarks>
public class EngineDriftProbe
{
    private static IXLFontEngine V1() =>
        DefaultFontEngine.CreateOnlyWithFonts(
            TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf"),
            TestHelper.GetStreamFromResource("Fonts.TestFontB.ttf"));

    private static IXLFontEngine Skia() =>
        SkiaSharpFontEngine.CreateOnlyWithFonts(
            TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf"),
            TestHelper.GetStreamFromResource("Fonts.TestFontB.ttf"));

    private static IEnumerable<(string Font, double Size, double Dpi, string Text)> Cases() =>
    [
        ("TestFontA", 20, 120, "A"),
        ("TestFontA", 11,  96, "Test"),
        ("TestFontA", 11,  96, "Lorem ipsum dolor sit amet"),
        ("TestFontA", 10,  96, "0"),
        ("TestFontB", 30,  96, "B"),
        ("TestFontB", 11,  96, "Lorem ipsum dolor sit amet"),
    ];

    /// <summary>
    /// Prints every metric from both adapters side by side with the relative difference, then
    /// asserts each is inside spec 27's tolerance. The printed table is the input to task 3.
    /// </summary>
    [Test]
    public async Task V1_and_SkiaSharp_agree_on_every_metric()
    {
        var v1 = V1();
        var skia = Skia();

        var failures = new List<string>();

        foreach (var (name, size, dpi, text) in Cases())
        {
            var font = new DriftFont(name, size);

            Compare($"{name}/{size}pt/{dpi}dpi TextWidth({text.Length} ch)",
                v1.GetTextWidth(text, font, dpi), skia.GetTextWidth(text, font, dpi),
                text.Length > 1 ? 0.01 : 0.005);

            Compare($"{name}/{size}pt/{dpi}dpi TextHeight",
                v1.GetTextHeight(font, dpi), skia.GetTextHeight(font, dpi), 0.005);

            Compare($"{name}/{size}pt/{dpi}dpi MaxDigitWidth",
                v1.GetMaxDigitWidth(font, dpi), skia.GetMaxDigitWidth(font, dpi), 0.005);

            Compare($"{name}/{size}pt/{dpi}dpi Descent",
                v1.GetDescent(font, dpi), skia.GetDescent(font, dpi), 0.005);

            Span<int> glyph = ['8'];
            var boxV1 = v1.GetGlyphBox(glyph, font, new Dpi(dpi, dpi));
            var boxSkia = skia.GetGlyphBox(glyph, font, new Dpi(dpi, dpi));

            Compare($"{name}/{size}pt/{dpi}dpi GlyphBox.AdvanceWidth",
                boxV1.AdvanceWidth, boxSkia.AdvanceWidth, 0.005);
            Compare($"{name}/{size}pt/{dpi}dpi GlyphBox.EmSize",
                boxV1.EmSize, boxSkia.EmSize, 0.0);
            Compare($"{name}/{size}pt/{dpi}dpi GlyphBox.Descent",
                boxV1.Descent, boxSkia.Descent, 0.0);
        }

        await Assert.That(failures).IsEmpty();

        void Compare(string label, double a, double b, double relativeTolerance)
        {
            var allowed = Math.Max(Math.Abs(a) * relativeTolerance, 0.01);
            var delta = Math.Abs(a - b);
            var pct = a == 0 ? 0 : delta / Math.Abs(a) * 100;

            Console.WriteLine(string.Format(CultureInfo.InvariantCulture,
                "| {0,-52} | {1,12:F6} | {2,12:F6} | {3,7:F3}% |", label, a, b, pct));

            if (delta > allowed)
                failures.Add($"{label}: V1={a}, SkiaSharp={b}, delta={delta} > {allowed}");
        }
    }

    private sealed class DriftFont(string name, double size) : IXLFontBase
    {
        public string FontName { get; set; } = name;
        public double FontSize { get; set; } = size;
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Strikethrough { get; set; }
        public XLFontUnderlineValues Underline { get; set; } = XLFontUnderlineValues.None;
        public XLFontVerticalTextAlignmentValues VerticalAlignment { get; set; }
        public bool Shadow { get; set; }
        public XLColor FontColor { get; set; } = XLColor.Black;
        public XLFontFamilyNumberingValues FontFamilyNumbering { get; set; }
            = XLFontFamilyNumberingValues.NotApplicable;
        public XLFontCharSet FontCharSet { get; set; } = XLFontCharSet.Default;
        public XLFontScheme FontScheme { get; set; }
    }
}
```

`GetGlyphBox.EmSize` and `.Descent` are compared with a relative tolerance of `0.0`, which the
`Math.Max(..., 0.01)` floor turns into "must be within 0.01 px" — they are whole numbers, so that
is exact equality with a float-comparison guard.

- [ ] **Step 3: Run it and capture the table**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/EngineDriftProbe/*"`

Expected: PASS, if `docs/font-architecture.md:63` is still true.

**Record the printed table verbatim in a Results section of this spec, pass or fail.** It is the
first cross-adapter measurement in the repository's history and it is the input to task 3's
constants.

- [ ] **Step 4: If it fails, stop and write it up**

A failure is a **real defect**, not a test-authoring problem. `GetMaxDigitWidth` is the divisor in
`XLHelper.cs:489` and `:518` that converts pixels to the OOXML column-width unit, so a divergence
there changes the `<col width="...">` attribute of every file saved by a consumer who picked a
different font package. `GetGlyphBox.Descent` and `.EmSize` set row heights through
`XLRow.cs:331`.

In that case:

1. Record which metric, which font, which size, and the percentage, in a Results section.
2. Determine which adapter is right, by reading the font's tables directly rather than by trusting
   either engine. Excel's own rule is OS/2 `usWinAscent`/`usWinDescent` (`IXLFontEngine.cs:32`),
   and the interface documents that it is *not* what the adapters do — so "right" may be "neither",
   which is itself the finding.
3. Decide, and say so in the spec, whether the fix belongs here or in spec 34. Either way task 3
   pins the divergence per adapter rather than widening a shared tolerance.
4. `docs/font-architecture.md:63` must be corrected in the same PR. A doc that claims 0% drift
   while CI measures otherwise is worse than no doc.

- [ ] **Step 5: Commit**

```bash
git add XLibur.Tests/Graphics/EngineDriftProbe.cs XLibur.Tests/XLibur.Tests.csproj
```
```bash
git commit -m 'test(fonts): measure metric drift between two adapters (spec 27 task 1)'
```

---

### Task 2 — Extract the conformance module

**Files:**
- Create: `XLibur.Fonts.Conformance/FontEngineConformance.cs`,
  `XLibur.Fonts.Conformance/ConformanceFixture.cs`, `XLibur.Fonts.Conformance/ConformanceFont.cs`
- Create: `XLibur.Fonts.SixLabors.Tests/SixLaborsConformanceTests.cs`,
  `XLibur.Fonts.SkiaSharp.Tests/SkiaSharpConformanceTests.cs`
- Delete: `XLibur.Fonts.SixLabors.Tests/SixLaborsFontEngineTests.cs` (431 lines),
  `XLibur.Fonts.SkiaSharp.Tests/SkiaSharpFontEngineTests.cs` (434 lines)
- Modify: both adapter `.csproj` files

**Interfaces:**
- Produces: `FontEngineConformance`, `ConformanceFixture`, `ConformanceFont`.

- [ ] **Step 1: Create the fixture abstraction and the font stub**

`ConformanceFont` is `DummyFont` with `sealed` added and a primary constructor — one copy replacing
the four at `FontTests.cs:160`, `SixLaborsFontEngineTests.cs:410`,
`SkiaSharpFontEngineTests.cs:413` and `SkiaSharpFontBootstrapTests.cs:121`. Leave the
`SkiaSharpFontBootstrapTests` copy alone for now; that file has no counterpart in the other suite
and is not part of the conformance surface.

The fixture also has to abstract resource loading. Each suite's `TestHelper` resolves
`$"{typeof(TestHelper).Namespace}.Resource.{name}"`, which differs per project, so
`ConformanceFixture` exposes `Stream OpenTestFont(string name)` and each concrete fixture forwards
to its own `TestHelper`. The font files are byte-identical across all three projects
(`21ca718e…`, `40d7a669…`), so the streams are interchangeable.

- [ ] **Step 2: Move the 31 tests**

Take `SkiaSharpFontEngineTests.cs` as the base — it is the later copy and its structure is the
cleaner of the two — and restore the two assertions it dropped:

```csharp
    /// <summary>
    /// TestFontA's "A" advances exactly 15/16 em, so at 20pt and 120 DPI the answer is
    /// 0.9375 * 20 * 120/72 = 31.25 px in any adapter that reads the font correctly.
    /// </summary>
    /// <remarks>
    /// The SkiaSharp copy of this test asserted only IsGreaterThan(0) and the SixLabors v2 copy
    /// widened the tolerance to +/-1.0 px (3.2%) with the comment "v2 may have slightly different
    /// measurement than v1". Neither was measured. Spec 27 decision 2 sets 0.5% and explains why
    /// widening it is not an available response to a failure.
    /// </remarks>
    [Test]
    public async Task CreateOnlyWithFonts_UsesProvidedFallback()
    {
        var engine = Fixture.CreateStreamEngine();
        var width = engine.GetTextWidth("A", new ConformanceFont("Nonexistent Font", 20), 120);

        await Assert.That(width).IsEqualTo(31.25d).Within(0.15625d); // 0.5% of 31.25
    }
```

Keep every monotonicity test. They cost nothing and they catch wiring bugs the golden table cannot
— an engine that returns the fallback font's metrics for a font it should have resolved will still
hit the golden numbers if the fallback happens to be the golden font.

- [ ] **Step 3: Wire the two adapter suites**

```csharp
namespace XLibur.Fonts.SkiaSharp.Tests;

/// <summary>Runs spec 27's shared conformance module against <see cref="SkiaSharpFontEngine"/>.</summary>
public sealed class SkiaSharpConformanceTests : FontEngineConformance
{
    protected override ConformanceFixture Fixture { get; } = new SkiaSharpFixture();

    private sealed class SkiaSharpFixture : ConformanceFixture
    {
        internal override string AdapterName() => "SkiaSharp";

        internal override Stream OpenTestFont(string name) =>
            TestHelper.GetStreamFromResource($"Fonts.{name}");

        internal override IXLFontEngine CreateStreamEngine() =>
            SkiaSharpFontEngine.CreateOnlyWithFonts(
                OpenTestFont("TestFontA.ttf"), OpenTestFont("TestFontB.ttf"));

        internal override IXLFontEngine? CreateEmbeddedFallbackEngine() =>
            SkiaSharpFontBootstrap.CreateDefault();
    }
}
```

The SixLabors v2 fixture returns `null` from `CreateEmbeddedFallbackEngine()` — that package embeds
no fonts (`docs/font-architecture.md:45`, and `XLibur.Fonts.SixLabors.csproj` has no
`EmbeddedResource` item). The CarlitoBare rows of the golden table skip on a `null` fixture.

- [ ] **Step 4: Add the linked Compile items**

To both adapter `.csproj` files:

```xml
  <ItemGroup>
    <!--
      Spec 27's shared font conformance module. Linked source rather than a project reference:
      TUnit's source generator emits test registrations into the assembly it compiles, and tests
      in a referenced library register from a module initializer that only fires once the CLR
      touches a type in it. Compiling the bodies here is what guarantees discovery. It also keeps
      the module out of the solution, so nothing can pack or version it.
    -->
    <Compile Include="..\XLibur.Fonts.Conformance\*.cs"
             Link="Conformance\%(Filename)%(Extension)" />
  </ItemGroup>
```

While in these files, add the CS4014 promotion that `XLibur.Tests.csproj:14` carries and these two
do not:

```xml
    <WarningsAsErrors>$(WarningsAsErrors);CS4014</WarningsAsErrors>
```

Use the Edit tool. Run: `git diff --numstat` on both.
Expected: fewer than 10 changed lines in each. A count near the file's 27 lines means it was
rewritten with LF endings.

- [ ] **Step 5: Delete the two copies**

```bash
git rm XLibur.Fonts.SixLabors.Tests/SixLaborsFontEngineTests.cs
```
```bash
git rm XLibur.Fonts.SkiaSharp.Tests/SkiaSharpFontEngineTests.cs
```

- [ ] **Step 6: Run both suites**

Run: `dotnet test XLibur.Fonts.SixLabors.Tests/XLibur.Fonts.SixLabors.Tests.csproj -f net10.0`
Run: `dotnet test XLibur.Fonts.SkiaSharp.Tests/XLibur.Fonts.SkiaSharp.Tests.csproj -f net10.0`

Expected: PASS, 31 conformance tests in each, plus the 6 `SkiaSharpFontBootstrapTests`.

If either reports **exit 8** the filter matched nothing — but there is no filter here, so exit 8
means TUnit discovered no tests at all and decision 1's premise about linked source is wrong.
Investigate before working around it. Exit 5 would mean an invalid runner option, which cannot
happen on an unfiltered run.

- [ ] **Step 7: Commit**

```bash
git add XLibur.Fonts.Conformance XLibur.Fonts.SixLabors.Tests XLibur.Fonts.SkiaSharp.Tests
```
```bash
git commit -m 'refactor(fonts): one conformance module for two adapter suites (spec 27 task 2)'
```

---

### Task 3 — The golden metric table

The conformance module now runs in two places but still asserts mostly relative properties. This
task gives every metric a constant.

**Files:**
- Create: `XLibur.Fonts.Conformance/GoldenMetrics.cs`
- Modify: `XLibur.Fonts.Conformance/FontEngineConformance.cs`

**Interfaces:**
- Produces: `GoldenMetric`, `GoldenMetrics.StreamFonts`, `GoldenMetrics.CarlitoBare`.

- [ ] **Step 1: Write the table from task 1's output**

```csharp
/// <summary>
/// Fixed expected metrics for the two stream-loaded test fonts. Every adapter must produce these
/// numbers, which is how agreement is established between two adapters that can never be loaded
/// into the same process.
/// </summary>
/// <remarks>
/// Values are exact rational fractions of the em square, because every adapter computes
/// <c>advance_fu * size / unitsPerEm</c>. The tolerance is 0.5% for arithmetic-only metrics and
/// 1% for multi-glyph GetTextWidth, where the SixLabors adapters pass KerningMode.None and
/// SkiaSharp uses SKFont.MeasureText's default shaping. See spec 27 decision 2. A metric the three
/// adapters cannot meet is recorded as a divergence with a per-adapter value, never covered by
/// widening the tolerance here.
/// </remarks>
internal static class GoldenMetrics
{
    internal static readonly GoldenMetric[] StreamFonts =
    [
        // Verbatim from spec 27 task 1's drift table; EmFraction is the ground truth the pixel
        // value is derived from, so a size or DPI can be added without re-measuring.
        new("TestFontA", 20, 120, "A", 31.25d,  0.15625d, "15/16"),
        new("TestFontB", 30,  96, "B", 25.00d,  0.125d,   "5/8"),
        // ... remaining rows filled from task 1's printed table
    ];

    /// <summary>
    /// Metrics for the embedded CarlitoBare fallback. Applies only to the adapters that embed it:
    /// V1 (XLibur.Fonts.SixLabors.V1.csproj:24-27) and SkiaSharp
    /// (XLibur.Fonts.SkiaSharp.csproj:30-33). XLibur.Fonts.SixLabors embeds no fonts, so its
    /// fixture returns null from CreateEmbeddedFallbackEngine and these rows are skipped.
    /// </summary>
    internal static readonly GoldenMetric[] CarlitoBare =
    [
        new("CarlitoBare", 11, 96, "8", 7.43359375d, 0.037d, "1038/2048"),
        // ... remaining rows filled from task 1's printed table
    ];
}
```

`7.43359375` is already asserted exactly against V1 at `FontTests.cs:38`, `:77` and `:127`. It is
the strongest constant in the repo and the one whose divergence would be most visible: it is the
divisor for the OOXML column-width unit.

- [ ] **Step 2: Drive it from the conformance tests**

```csharp
    /// <summary>
    /// Every adapter produces the same pixel value for the same font, size and DPI. This is what
    /// makes the three adapters interchangeable — not agreement measured directly, which is
    /// impossible for the SixLabors pair, but agreement with one shared table.
    /// </summary>
    [Test]
    [MethodDataSource(typeof(GoldenMetrics), nameof(GoldenMetrics.StreamFontRows))]
    public async Task Stream_font_metrics_match_the_golden_table(GoldenMetric metric)
    {
        var engine = Fixture.CreateStreamEngine();
        var font = new ConformanceFont(metric.Font, metric.SizePt);
        var actual = engine.GetTextWidth(metric.Text!, font, metric.Dpi);

        await Assert.That(actual)
            .IsEqualTo(metric.ExpectedPx)
            .Within(metric.TolerancePx)
            .Because($"{Fixture.AdapterName()} must read {metric.Font} '{metric.Text}' as "
                   + $"{metric.EmFraction} em");
    }
```

Repeat for `GetMaxDigitWidth`, `GetDescent`, `GetTextHeight` and the three `GlyphBox` components,
each with its own table and its own tolerance from decision 2. `GlyphBox.EmSize` and
`.Descent` use `IsEqualTo` with no `Within` — they are whole numbers.

Every `Assert.That` here must be `await`ed. A missing `await` makes the test pass while asserting
nothing, which for a spec whose entire purpose is asserting numbers would be a silent total
failure. `CS4014` is promoted to an error in all three projects after task 2 step 4; check the
build does not warn rather than trusting it.

- [ ] **Step 3: Run both suites**

Run: `dotnet test XLibur.Fonts.SixLabors.Tests/XLibur.Fonts.SixLabors.Tests.csproj -f net10.0`
Run: `dotnet test XLibur.Fonts.SkiaSharp.Tests/XLibur.Fonts.SkiaSharp.Tests.csproj -f net10.0`

Expected: PASS.

**A failure here is a finding, not a tolerance problem.** The two adapters that this task covers
were measured against each other in task 1, so a value outside tolerance at this point means the
constant was transcribed wrong or the v2 adapter — which task 1 could not measure, because it
cannot be loaded next to V1 — is the one that differs. If it is the v2 adapter, record it, pin the
v2 value separately with a comment naming the metric and the cause, and leave the shared tolerance
alone.

- [ ] **Step 4: Verify the gate bites**

Temporarily change one `GoldenMetric` row's expected value by 2% and re-run. Expected: FAIL, with
the `.Because` message naming the adapter and the em fraction. Restore the row.

This step is not optional. A golden table that cannot fail is the same thing as the
`IsGreaterThan(0)` assertions it replaces.

- [ ] **Step 5: Commit**

```bash
git add XLibur.Fonts.Conformance
```
```bash
git commit -m 'test(fonts): assert every metric against a golden table (spec 27 task 3)'
```

---

### Task 4 — V1 joins the conformance run

**Files:**
- Create: `XLibur.Tests/Graphics/SixLaborsV1ConformanceTests.cs`
- Delete: `XLibur.Tests/Graphics/EngineDriftProbe.cs`
- Modify: `XLibur.Tests/Graphics/FontTests.cs`, `XLibur.Tests/XLibur.Tests.csproj`

- [ ] **Step 1: Add the linked Compile items and the V1 fixture**

Same `<Compile Include="..\XLibur.Fonts.Conformance\*.cs" …>` block as task 2 step 4, and a
`SixLaborsV1ConformanceTests : FontEngineConformance` with a fixture built on
`DefaultFontEngine.CreateOnlyWithFonts(...)` and `new DefaultFontEngine("NonExistentFallbackFont")`
for the CarlitoBare tier — that is how `FontTests.cs:71` already forces resolution down to the
embedded font.

- [ ] **Step 2: Reduce `FontTests.cs` to what is V1-specific**

The tests that move into the conformance module and can be deleted from `FontTests.cs`:
`CanSpecifyFallbackFontWithoutFileSystem` (`:82`),
`CanSpecifyExtraFontsAsStreamsWithoutFileSystem` (`:96`),
`DefaultFontEngine_CanBeUsedDirectly` (`:122`),
`DefaultFontEngine_CanSpecifyFallbackFontWithoutFileSystem` (`:131`),
`FontEngine_CanBeInjectedViaLoadOptions` (`:144`),
`UseEmbeddedFontWhenFallbackFontIsNotPresent` (`:68`).

The tests that stay, because they are about system fonts or about `DefaultGraphicEngine` rather
than about the port: `CalculatedTextWidth` (`:15`), `CalculatedTextHeight` (`:24`),
`GetMaxDigitWidth` (`:34`), `DescentIsPositive` (`:43`), `NonExistentFontUsesFallback` (`:53`),
`Issue_1916_CanMeasureSpecificArabicText` (`:110`).

The first four of those carry wide tolerances — `Within(100)` on 500, `Within(0.5)` on 3.667 —
whose comments say why (`:28`, `:47`): *"Calibri on Windows vs Carlito fallback on Linux"*. Those
are **font** differences across operating systems, not **engine** differences, and they are
correctly wide. Leave them, and add a comment saying so, so a later reader does not cite them as
precedent for widening a conformance tolerance:

```csharp
    // The wide tolerance here is an OS difference (real Calibri on Windows, the CarlitoBare
    // fallback elsewhere), not an engine difference. Spec 27's conformance tolerances are 0.5%
    // because they use stream-loaded fonts, where the font file is byte-identical everywhere.
```

- [ ] **Step 3: Delete the drift probe**

Task 1's probe is superseded: with V1 in the conformance run, all three adapters assert the same
table, and the probe's direct V1↔SkiaSharp comparison adds nothing the table does not cover. Its
numbers live on in `GoldenMetrics` and in this spec's Results.

```bash
git rm XLibur.Tests/Graphics/EngineDriftProbe.cs
```

Keep the SkiaSharp `ProjectReference` — task 5 needs it.

- [ ] **Step 4: Run all three suites, both frameworks**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj`
Run: `dotnet test XLibur.Fonts.SixLabors.Tests/XLibur.Fonts.SixLabors.Tests.csproj`
Run: `dotnet test XLibur.Fonts.SkiaSharp.Tests/XLibur.Fonts.SkiaSharp.Tests.csproj`

Expected: PASS on net8.0 and net10.0 in all three.

- [ ] **Step 5: Commit**

```bash
git add XLibur.Tests
```
```bash
git commit -m 'test(fonts): run the conformance module against V1 too (spec 27 task 4)'
```

---

### Task 5 — Run the core suite under the shipped default

The conformance module proves the five metrics agree. It does not prove the library works when
driven by SkiaSharp, because no core test has ever done that.

**Files:**
- Modify: `XLibur.Tests/TestInfrastructure.cs:28-33`
- Modify: `.github/workflows/build-and-test.yml`

- [ ] **Step 1: Make the engine selectable**

```csharp
    /// <summary>
    /// Registers the font engine the suite runs under. V1 by default; set
    /// <c>XLIBUR_TEST_FONT_ENGINE=skiasharp</c> for the second CI pass that runs the whole suite
    /// against the engine consumers actually get by default
    /// (XLibur/Graphics/DefaultFontEngineProbe.cs:30).
    /// </summary>
    /// <remarks>
    /// Assigned, not registered. Both bootstraps use <c>??=</c>, and XLibur.Fonts.SixLabors.V1 also
    /// auto-registers from a [ModuleInitializer] the moment the CLR touches a V1 type — so calling
    /// SkiaSharpFontBootstrap.Register() here would silently no-op and the "SkiaSharp pass" would
    /// run V1 while reporting success.
    /// </remarks>
    [Before(HookType.Assembly)]
    public static void GlobalSetup()
    {
        var requested = Environment.GetEnvironmentVariable("XLIBUR_TEST_FONT_ENGINE");
        LoadOptions.DefaultFontEngine =
            string.Equals(requested, "skiasharp", StringComparison.OrdinalIgnoreCase)
                ? Fonts.SkiaSharp.SkiaSharpFontBootstrap.CreateDefault()
                : Fonts.SixLabors.V1.DefaultFontEngine.Instance.Value;

        SetCulture(DefaultCulture);
    }
```

The `??=` trap in step 1's remarks is the single thing most likely to make this task produce a
false green. Verify it directly:

```csharp
    /// <summary>The suite is running under the engine the environment asked for.</summary>
    [Test]
    public async Task The_suite_runs_under_the_requested_font_engine()
    {
        var requested = Environment.GetEnvironmentVariable("XLIBUR_TEST_FONT_ENGINE") ?? "v1";
        var expected = requested == "skiasharp" ? "SkiaSharpFontEngine" : "DefaultFontEngine";

        using var wb = new XLWorkbook();
        await Assert.That(wb.FontEngineTypeNameForTests).IsEqualTo(expected);
    }
```

`XLWorkbook.FontEngine` is `internal` (`XLibur/Excel/XLWorkbook.cs:115`) and `XLibur.Tests` already
reads core internals, so assert on `LoadOptions.DefaultFontEngine!.GetType().Name` instead if no
`InternalsVisibleTo` grant reaches this member — the point is only that the requested engine is the
one in force.

- [ ] **Step 2: Run the whole core suite under SkiaSharp locally**

```bash
XLIBUR_TEST_FONT_ENGINE=skiasharp dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0
```

Expected: PASS.

The suite has 74 absolute `Width`/`Height` assertions, but most are chart and picture geometry
(`ChartAnchorTests.cs:89`, `PictureTests.cs:29`, …) which never reach the font engine. The
font-derived ones are mostly relative — `IsNotEqualTo(widthBefore)`,
`IsGreaterThan(defaultWidth)` (`ColumnsUsedLinqTests.cs:71`, `RichTextColorRoundTripTests.cs:249`,
`RowTests.cs:486`) — so this pass is expected to be green today.

**That expectation is the premise, and a green run is the useful result.** It proves the shipped
default engine can drive the whole library, which nothing has ever shown. A failure is a defect in
the default engine, affecting every zero-config consumer, and belongs in Results with the same
weight task 1's would have had.

Do not soften a core assertion to make this pass. If a font-derived assertion is genuinely
engine-specific, mark it `[Skip]` under the SkiaSharp pass with a comment naming the metric, and
add that metric to `GoldenMetrics` in task 3 so the underlying disagreement is pinned rather than
skipped away.

- [ ] **Step 3: Add the CI pass**

After the existing core-suite step (`.github/workflows/build-and-test.yml:99-106`), matching its
folded-scalar form and its `--configuration Release --no-build --framework net10.0` flags:

```yaml
      # Spec 27 task 5. The step above runs the core suite against XLibur.Fonts.SixLabors.V1; this
      # runs it against XLibur.Fonts.SkiaSharp, the engine a zero-config consumer actually gets
      # (XLibur/Graphics/DefaultFontEngineProbe.cs:30). Without this pass the shipped default has
      # never driven the autofit, row-height or comment-sizing paths.
      # No --coverage here: the second pass duplicates the first's line coverage and would
      # overwrite coverage.xml, which sonar.cs.vscoveragexml.reportsPaths reads.
      - name: Test core under the default font engine
        env:
          XLIBUR_TEST_FONT_ENGINE: skiasharp
        run: >
          dotnet test XLibur.Tests/XLibur.Tests.csproj
          --configuration Release
          --no-build
          --framework net10.0
          --report-trx --report-trx-filename skiasharp-core-test-results.trx
```

- [ ] **Step 4: Commit**

```bash
git add XLibur.Tests/TestInfrastructure.cs .github/workflows/build-and-test.yml
```
```bash
git commit -m 'test(fonts): run the core suite against the shipped default engine (spec 27 task 5)'
```

---

## Acceptance criteria

1. `XLibur.Fonts.Conformance/` exists, contains no `.csproj`, and its `.cs` files are linked into
   exactly three test projects. Gate:
   `grep -rl "XLibur.Fonts.Conformance" --include=*.csproj .` returns 3 paths, and
   `find XLibur.Fonts.Conformance -name "*.csproj"` returns nothing.
2. The two copied suites are gone. Gate:
   `test ! -f XLibur.Fonts.SixLabors.Tests/SixLaborsFontEngineTests.cs` and
   `test ! -f XLibur.Fonts.SkiaSharp.Tests/SkiaSharpFontEngineTests.cs`.
3. No `IsGreaterThan(0)` survives as the *only* assertion on a metric that has a golden row. Gate:
   `grep -c "IsGreaterThan(0)" XLibur.Fonts.Conformance/FontEngineConformance.cs` is at most the
   count of metrics with no golden row, and that count is stated in the file's remarks.
4. Every one of the five `IXLFontEngine` methods has at least one golden row. Gate:
   `grep -c "new GoldenMetric(" XLibur.Fonts.Conformance/GoldenMetrics.cs` is at least 5, and
   `grep -cE "GetTextHeight|GetTextWidth|GetMaxDigitWidth|GetDescent|GetGlyphBox" XLibur.Fonts.Conformance/FontEngineConformance.cs`
   covers all five names.
5. `DummyFont` drops from four declarations to at most two. Gate:
   `grep -rc "class DummyFont" --include=*.cs . | grep -v ":0"` lists at most
   `SkiaSharpFontBootstrapTests.cs` plus the single `ConformanceFont`.
6. Task 3 step 4's gate-bites check was performed and its FAIL observed.
7. `dotnet test` is green on net8.0 and net10.0 for `XLibur.Tests`,
   `XLibur.Fonts.SixLabors.Tests` and `XLibur.Fonts.SkiaSharp.Tests`, each named by `.csproj`,
   never at solution level.
8. `XLibur.Tests` passes with `XLIBUR_TEST_FONT_ENGINE=skiasharp`, and CI runs that pass. Gate:
   `grep -c "XLIBUR_TEST_FONT_ENGINE" .github/workflows/build-and-test.yml` returns at least 1.
9. Task 1's cross-adapter drift table is recorded verbatim in this spec's Results section, whether
   it passed or failed.
10. Any metric on which the adapters disagree beyond decision 2's tolerances is recorded in Results
    with the metric, the font, the size and the percentage, and is pinned per adapter — never
    covered by widening a shared tolerance.
11. `docs/font-architecture.md:63` either still says 0% drift and is now backed by CI, or is
    corrected with the measured number.
12. No production file under `XLibur/` or any `XLibur.Fonts.*` (non-`Tests`) project is modified.
    Gate: `git diff --stat main -- XLibur XLibur.Fonts.SixLabors XLibur.Fonts.SixLabors.V1 XLibur.Fonts.SkiaSharp`
    returns nothing.
13. `git diff --numstat` shows no `.csproj` with a changed-line count near its total line count —
    that is the CRLF-rewrite signature.

## Conflicts

**None in production code.** This spec modifies no file under `XLibur/` and no adapter
implementation. `git diff --stat` against `main` for those directories must be empty (criterion 12).

- **Spec 34 hard-depends on this one.** 34 leaves `IXLFontEngine` unchanged — the three shipped
  adapters stop implementing it directly and implement a narrower mechanism port instead — but it
  moves metric computation across all three adapters. There is currently no test anywhere that would notice if that work
  changed a number in one adapter and not the others, and two of the three adapters can never be
  loaded into the same process to be compared directly. **34 must not start before task 3 lands.**
  Tasks 4 and 5 can run concurrently with 34's early work; tasks 1–3 cannot.
- **Spec 20** (`20-style-key-struct-size.md`) touches `XL*Key.cs`. No overlap.
- **Specs 22, 23, 24, 25** are the architecture-deepening set. 22 is chart IO, 23 is style facades,
  24 is the worksheet element load, 25 is the formula shifter. None touches the font port, any
  adapter, or any of the three font test projects.
- **Spec 18 task 5** owns per-sheet load cost in `XLWorkbook_Load`. This spec reads
  `XLWorkbook_Load.cs:948` as evidence and modifies nothing there.
- **`XLibur.Tests.csproj`** gains one `ProjectReference` and one `ItemGroup`. Any other spec
  editing that file should land first or rebase; the change is three lines.

Spec 25's shape is the model for what to do if task 1 finds a divergence: it does not reconcile the
nine rows on which the two formula shifters disagree, it records them in the data and asserts both
columns on every run. A font metric the three adapters cannot agree on gets the same treatment.

## Closing note — three copies of CarlitoBare, one of them dead

Checked while gathering evidence, and confirmed:

- `XLibur/Graphics/Fonts/CarlitoBare-{Regular,Bold,Italic,BoldItalic}.ttf` are **tracked**
  (`git ls-files` lists all four) and **embedded by nothing**. `XLibur.csproj:25-26` declares
  exactly two `EmbeddedResource` items, both XML.
- `XLibur.Fonts.SixLabors.V1.csproj:24-27` embeds its own copy from
  `XLibur.Fonts.SixLabors.V1/Fonts/`, under `LogicalName="XLibur.Graphics.Fonts.…"` — the core
  logical name, which is what `DefaultFontEngine.cs:235` looks up. That is why the core copy looks
  live and is not.
- `XLibur.Fonts.SkiaSharp.csproj:30-33` embeds a third copy under its own logical name.
- All three physical copies are byte-identical (`f2f52428f1f3fe4d3cbca3bae238b775` for
  `CarlitoBare-Regular.ttf`).

So four tracked font files in core appear to be dead weight left behind when the port was cut.
Deleting them is a one-line change with an obvious gate
(`grep -rn "XLibur.Graphics.Fonts" --include=*.csproj .` must still find V1's four `LogicalName`
attributes, since V1 depends on that exact string). It is **not** a task in this spec — it is
unrelated to the conformance seam, and folding it in would make criterion 12 untrue. Raise it
separately.

The same duplication exists for the test fonts: six files, three projects, two distinct fonts, all
byte-identical (`21ca718e11aadbd392499ab5e5084bef`, `40d7a669ff482fec490d9a964da9a68d`). Task 2's
`ConformanceFixture.OpenTestFont` abstraction is what would let those collapse to one copy later;
this spec does not do it, because each test project embeds its resources through its own assembly
and consolidating them is a separate packaging change.
