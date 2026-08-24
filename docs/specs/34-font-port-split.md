# Spec 34 — Split the font port: mechanism below, policy above

**Area:** Architecture · Refactor · **Compat (undeclared error mode)**
**Effort:** M (~5–6 days)
**Dependencies:** **[Spec 27](27-font-conformance-suite.md) is a hard prerequisite, through its task 3.**
This spec moves metric computation across all three shipped adapters and nothing in the tree today
asserts that any two adapters agree. Without 27's shared conformance module
(`XLibur.Fonts.Conformance/`) and its golden metric table (`GoldenMetrics.cs`, 27's task 3) there is
no gate, and task 0 exists solely to refuse to start without it. 27 states the same dependency from
its side. Otherwise file-disjoint from every open spec.
**Status:** Proposed.

## Goal

Draw the font seam at mechanism instead of at policy. `IXLFontEngine` — the **core-facing** port —
does not change. What changes is that the three shipped adapters stop implementing 5 methods of
policy each and implement a 4-member mechanism port, `IXLTypefaceSource`, with the caching, fallback
chain, unit conversion, digit scan and glyph-box assembly moved into one core module, `XLFontMetrics`.

A fourth adapter should cost a typeface lookup, not 300 lines.

## Why this spec exists

**The port is not the problem.** `IXLFontEngine` is 5 methods over 55 lines
(`XLibur/Graphics/IXLFontEngine.cs:10-55`) and core reaches font measurement through three entry
points — column autofit, row autofit, comment autosize. By the deletion test it is a deep module: a
lot of font-library knowledge behind a very small door. Nothing below asks to widen or narrow it.

The problem is **where the line is drawn**. Everything on the far side of that door is policy that has
nothing to do with any font library, and it is written out three times.

### Three adapters, 897 lines, one algorithm

| Adapter | File | Lines |
|---|---|---:|
| SixLabors.Fonts 1.0.1 | `XLibur.Fonts.SixLabors.V1/DefaultFontEngine.cs` | 306 |
| SixLabors.Fonts 2.1.3 | `XLibur.Fonts.SixLabors/SixLaborsFontEngine.cs` | 277 |
| SkiaSharp | `XLibur.Fonts.SkiaSharp/SkiaSharpFontEngine.cs` | 314 |
| | **Total** | **897** |

> **Correction to the survey that prompted this spec.** It recorded V1 at 277 and v2 at 306. The two
> are the other way round: V1 is the 306-line file, v2 the 277-line one. The 897 total is right.

Measured with `diff` (longest-common-subsequence, so these are lines in the same order, not just a
multiset intersection):

```
diff --unchanged-group-format='%=' --old-group-format='' --new-group-format='' \
     --changed-group-format='' A B | wc -l
```

| Pair | Identical lines | Non-blank |
|---|---:|---:|
| V1 ↔ v2 | **229** | 190 |
| V1 ↔ Skia | 150 | — |
| v2 ↔ Skia | 171 | — |
| **All three** | **145** | **109** |

229 of v2's 277 lines are byte-identical to V1 — **83% of the file**. 145 lines are identical across
all three.

> **Second correction.** The survey reported "150 byte-identical across all three." 150 is the
> V1↔Skia *pairwise* count. Chaining the three gives **145** (109 non-blank). The claim survives; the
> number moves by five.

The duplicated 145 lines are not library glue. They are the same eight pieces of policy, transcribed:

| Duplicated policy | V1 | v2 | Skia |
|---|---|---|---|
| `PointsToPixels` | `:275` | `:246` | `:244` |
| The `MetricId` cache key struct | `:277-305` | `:248-276` | `:269-305` |
| `ConcurrentDictionary` face cache | `:34` | `:24` | `:25` |
| `ConcurrentDictionary` max-digit-width cache | `:41` | `:27` | `:28` |
| `GetMaxDigitWidth` | `:143-148` | `:129-134` | `:131-136` |
| `GetTextHeight` | `:150-156` | `:136-142` | `:138-143` |
| `GetDescent` (public + private overload) | `:132-141` | `:118-127` | `:120-129` |
| `GetGlyphBox` assembly tail | `:194-200` | `:176-182` | `:162-168` |
| The 0–9 digit scan | `:250-273` | `:221-244` | `:224-233` |

`GetMaxDigitWidth` is identical in all three, to the character:

```csharp
    public double GetMaxDigitWidth(IXLFontBase font, double dpiX)
    {
        var metricId = new MetricId(font);
        var maxDigitWidth = _maxDigitWidths.GetOrAdd(metricId, _calculateMaxDigitWidth);
        return PointsToPixels(maxDigitWidth * font.FontSize, dpiX);
    }
```

Not one line of that mentions SixLabors or Skia.

`PointsToPixels` is a fourth-time offender: core already has the same conversion at
`XLibur/XLHelper.cs:444`, written the other way round (`points * dpi / 72d` rather than
`points / 72d * dpi`). V1's copy at `:275` is `internal static` where v2's and Skia's are
`private static` — and grep finds **no user of that `internal`** anywhere in the tree. It is
vestigial visibility on a duplicated constant.

### The fallback divergence — the strongest evidence, and not what the survey said

`IXLFontEngine` documents **zero error modes**. Read all five doc comments
(`IXLFontEngine.cs:12-54`): no `<exception>` tag, no statement of what happens when a font is not
installed. Yet all three adapters end their resolution chain differently, and which one a caller gets
depends on which NuGet package is referenced.

**V1 — never fails.** `DefaultFontEngine.LoadFont` (`:219-230`):

```csharp
        if (!_fontCollection.Value.TryGet(metricId.Name, out var fontFamily) &&
            !_fontCollection.Value.TryGet(_fallbackFont, out fontFamily))
        {
            fontFamily = _fontCollection.Value.Get(EmbeddedFontName);
        }
```

`EmbeddedFontName` is the hardcoded `"CarlitoBare"` (`:24`), and `AddEmbeddedFont` is called
unconditionally from **every** constructor (`:60`, `:81`). There is no configuration of V1 in which
this chain can fail. An unknown font silently gets Calibri-compatible metrics.

**v2 — throws.** `SixLaborsFontEngine.LoadFont` (`:201-219`):

```csharp
            if (_embeddedFontName is not null)
                fontFamily = _fontCollection.Value.Get(_embeddedFontName);
            else
                throw new InvalidOperationException(
                    $"Font '{metricId.Name}' not found, and fallback font '{_fallbackFont}' is also not available. " +
                    "Consider providing a fallback font stream or using system fonts.");
```

`_embeddedFontName` is `null` in both public constructors (`:43`, `:91`) and is set only by the
four-argument constructor at `:55`. **`XLibur.Fonts.SixLabors` ships no font assets at all** — the
package directory contains one `.cs` file and a README, no `Fonts/` folder, which
`docs/font-architecture.md:45` records as deliberate. So no ordinary configuration of v2 reaches V1's
behaviour; `new SixLaborsFontEngine("Arial")` on a box without Arial throws.

**Skia — throws, with one extra step.** `ResolveTypeface` (`:185-208`) is a four-level chain ending in
the same `InvalidOperationException` at `:205-207`, with the same message text as v2. Its extra step
is `TryMatchSystemFont` (`:210-222`), which rejects a system match whose family name is not exactly
equal:

```csharp
        var match = SKFontManager.Default.MatchFamily(familyName, style);
        if (match is not null && string.Equals(match.FamilyName, familyName, StringComparison.OrdinalIgnoreCase))
```

`SKFontManager.MatchFamily` substitutes rather than failing, so without this check every unknown font
would silently resolve to whatever Skia picks. SixLabors' `TryGet` is already exact, so this is Skia
**compensating for a library difference**, not expressing a different policy.

> **Third correction.** The survey read Skia's divergence as "requires an exact family-name match
> against system fonts," citing `:210-222`. That range is the compensation, and it is mechanism. The
> actual terminal behaviour is at `:205-207` and it **agrees with v2**. The odd one out is V1.

So the real shape is narrower and more useful than "three policies":

| | Terminal behaviour when nothing resolves | Ships an embedded family? |
|---|---|---|
| V1 | CarlitoBare metrics, always | Yes — 4 `.ttf`, added unconditionally |
| v2 | `InvalidOperationException` | **No** — package has no font assets |
| Skia, constructed directly | `InvalidOperationException` | Yes, but only if the caller names it |
| Skia, `CreateDefault()` (the zero-config path) | CarlitoBare metrics | Yes — `SkiaSharpFontBootstrap.cs:35`, `:76` |

**The chains are one chain with a configuration difference**, plus one piece of library
compensation. That reframing is the most valuable output of writing this spec down, and it makes
task 2 far smaller than it looked — but it does not make the divergence go away. It is still
observable from identical user code, still undeclared by the port, and still — critically — **pinned
by nothing**:

```
grep -rn "InvalidOperationException" XLibur.Fonts.SixLabors.Tests XLibur.Fonts.SkiaSharp.Tests XLibur.Tests/Graphics
  -> no matches
```

Both packages have a `#region Fallback behavior` (`SixLaborsFontEngineTests.cs:210`,
`SkiaSharpFontEngineTests.cs:210`) and both build their engine with
`CreateOnlyWithFonts(TestFontA.ttf)`, whose fallback family is present. They exercise level 1 of the
chain and never reach the terminal branch. **No test in the repository asserts what happens when a
font cannot be resolved, in any of the three packages.**

### Thread-safety is undeclared and universally assumed

`IXLFontEngine` says nothing about concurrency. All three adapters nevertheless use
`ConcurrentDictionary` for both caches and hoist the factory delegate into a field so `GetOrAdd` does
not allocate a closure per call (`V1:36,43`; `v2:25,28`; `Skia:26,29`). That is three independent
authors arriving at the same undeclared invariant. It belongs to whoever owns the caching, which
after this spec is one core module that can state it once.

### The same five signatures are declared four times

| Declaration | Location |
|---|---|
| `IXLGraphicEngine` (the pre-split port) | `XLibur/Graphics/IXLGraphicEngine.cs:25,31,37,43,64` |
| `IXLFontEngine` (the port) | `XLibur/Graphics/IXLFontEngine.cs:15,21,27,33,54` |
| `DefaultGraphicEngine` (forwards to the font engine) | `XLibur/Graphics/DefaultGraphicEngine.cs:61,66,68,73,76` |
| `GraphicEngineFontAdapter` (forwards to the graphic engine) | `XLibur/Graphics/GraphicEngineFontAdapter.cs:12,15,18,21,24` |

`GraphicEngineFontAdapter` is 14 lines of pure pass-through and shallow by the deletion test — but it
**earns its keep**. It is the only thing that keeps a user who implemented `IXLGraphicEngine` before
`IXLFontEngine` existed working, since `IXLGraphicEngine` carries the same five signatures
(`docs/font-architecture.md:179-181`). It is reached at `XLWorkbook.cs:1007`. **This spec does not
touch it.** Deleting it is a breaking change dressed as a cleanup.

### Where core actually reads the engine

`XLWorkbook.FontEngine` (`XLibur/Excel/XLWorkbook.cs:115`, `internal`) is read at **six** sites:

| Site | Purpose |
|---|---|
| `XLibur/Excel/Columns/XLColumn.cs:190` | `AdjustToContents` — column autofit |
| `XLibur/Excel/Rows/XLRow.cs:294` | `AdjustToContents` — row autofit |
| `XLibur/Excel/IO/VmlDrawingPartWriter.cs:265` | `EstimateAutoSizeHeight` — comment autosize |
| `XLibur/Excel/XLWorkbook_Load.cs:948` | `CalculateColumnWidth` on load |
| `XLibur/XLHelper.cs:489` | `NoCToPixels` |
| `XLibur/XLHelper.cs:518` | `ConvertWidthToNoC` |

The last three are all `GetMaxDigitWidth`. `IXLFontEngine` is then threaded as a parameter into
`XLColumn.CalculateMinColumnWidth:223`, `XLRow.CalculateMinRowHeight:314`,
`VmlDrawingPartWriter.CountWrappedLines:302`, `XLCell.GetGlyphBoxes:1484` and
`XLCellGlyphHelper.cs:21,43`.

> **Fourth correction.** The survey said three call sites and cited `XLibur/Utils/XLHelper.cs`. There
> are six reads, and `XLHelper.cs` sits at `XLibur/XLHelper.cs` — there is no `Utils` directory.

Core throws when no engine can be resolved (`XLWorkbook.cs:994-1005`), after a four-level resolution
order ending in `DefaultFontEngineProbe.TryResolveDefault()`. There is no silent no-op adapter. That
is correct and this spec keeps it.

### Why the probe exists (read before touching registration)

`docs/font-architecture.md:77-81` records that a `[ModuleInitializer]` in the font package cannot
work for zero-config usage: `new XLWorkbook()` touches only core types, so the font assembly is never
loaded and its initializer never fires — which is also why auto-registration via the `XLibur.Bundle`
meta-package (which has no code of its own) did not work. Core is the one assembly guaranteed to be
loaded, so discovery has to originate there. **This spec changes nothing about registration.**

### The test suites are duplicated too

`XLibur.Fonts.SixLabors.Tests/SixLaborsFontEngineTests.cs` (431 lines) and
`XLibur.Fonts.SkiaSharp.Tests/SkiaSharpFontEngineTests.cs` (434 lines) share **416 identical lines**
by raw `diff`. Spec 27 measures it better — normalising the engine name away leaves **13 differing
lines out of 865** — and owns the fix: both files are deleted in its task 2 and replaced by one
shared conformance module. That is the same disease as this spec's, one layer up, and it is why 27
comes first: two near-identical suites that never compare their two subjects to each other are
exactly the gate this refactor needs and does not have.

## Non-goals

- **Not changing `IXLFontEngine`.** The core-facing port keeps its five methods and its signatures.
  The six core read sites and any third-party implementer are untouched. If this spec ends up wanting
  to change `IXLFontEngine`, it has gone wrong — stop and re-scope.
- **Not deleting `GraphicEngineFontAdapter`.** See above.
- **Not upgrading SixLabors.Fonts.** V1 stays on 1.0.1 (Apache 2.0). The v2 package stays on 2.1.3
  and stays a separate package. The license split is *why* three adapters exist and is the constraint
  this spec works inside.
- **Not adding a fourth adapter.** A stated win is that a fourth would now cost a typeface lookup
  rather than 300 lines. Proving that by writing one is out of scope.
- **Not changing registration, `DefaultFontEngineProbe`, or the bootstraps** beyond what converting an
  adapter's constructor mechanically requires.
- **No change to the public constructors or factory methods** of any of the three engines. Every
  `CreateOnlyWithFonts` / `CreateWithFontsAndSystemFonts` / named-fallback constructor keeps its
  signature and its semantics.
- **No metric change.** Every number the golden table records must come out the same, save for the
  one decision task 2 names.

## Current state

Verified against the tree at `1b41cadd` (2026-08-24). Every line number below was read from the file,
not carried forward.

**Core (unchanged by this spec except where noted):**

- `XLibur/Graphics/IXLFontEngine.cs:10-55` — the port. `GetTextHeight:15`, `GetTextWidth:21`,
  `GetMaxDigitWidth:27`, `GetDescent:33`, `GetGlyphBox:54`. No `<exception>` tag anywhere.
- `XLibur/Graphics/IXLGraphicEngine.cs:25,31,37,43,64` — the same five signatures, pre-split.
- `XLibur/Graphics/DefaultGraphicEngine.cs:9` implements both; forwards at `:61,66,68,73,76`.
- `XLibur/Graphics/GraphicEngineFontAdapter.cs:10-26` — 14 lines of pass-through, **kept**.
- `XLibur/Graphics/DefaultFontEngineProbe.cs:43` — `TryResolveDefault`.
- `XLibur/Excel/XLWorkbook.cs:115` — `internal IXLFontEngine FontEngine`; `:994-1007` resolution;
  throw at `:1001-1005`.
- `XLibur/XLHelper.cs:444` — a fourth copy of `PointsToPixels`.

**V1 — `XLibur.Fonts.SixLabors.V1/DefaultFontEngine.cs`, 306 lines:**

- `EmbeddedFontName = "CarlitoBare"` `:24`; `FontMetricSize = 16f` `:26`
- caches `:34`, `:41`; hoisted factories `:36`, `:43`
- `GetDescent` `:132-136` + private overload `:138-141`
- `GetMaxDigitWidth` `:143-148` · `GetTextHeight` `:150-156` · `GetTextWidth` `:158-167`
- `GetGlyphBox` `:170-201`, assembly tail `:194-200`
- `GetMetrics` `:203-207` · `GetFont` `:209-217` · `LoadFont` `:219-230`
- `AddEmbeddedFont` `:232-248` · `CalculateMaxDigitWidth` `:250-273`
- `PointsToPixels` `:275` (`internal`, no users) · `MetricId` `:277-305`
- `Fonts/` holds four `CarlitoBare-*.ttf`; resource path `XLibur.Graphics.Fonts.CarlitoBare-{0}.ttf`
  (`:235`)

**v2 — `XLibur.Fonts.SixLabors/SixLaborsFontEngine.cs`, 277 lines:**

- `_embeddedFontName` `:22`, null in both public ctors (`:43`, `:91`), set only by the four-arg ctor
  `:55`
- caches `:24`, `:27`
- `GetDescent` `:118-122` + `:124-127` · `GetMaxDigitWidth` `:129-134` · `GetTextHeight` `:136-142`
- `GetTextWidth` `:144-153` · `GetGlyphBox` `:156-183`, tail `:176-182`
- `GetMetrics` `:185-189` · `GetFont` `:191-199` · `LoadFont` `:201-219`, throw `:212-215`
- `CalculateMaxDigitWidth` `:221-244` · `PointsToPixels` `:246` · `MetricId` `:248-276`
- **No font assets in the package.**

**Skia — `XLibur.Fonts.SkiaSharp/SkiaSharpFontEngine.cs`, 314 lines:**

- `_streamFonts` `:20`, `_useSystemFonts` `:21`, `_embeddedFontName` `:23`, caches `:25`, `:28`
- `GetDescent` `:120-124` + `:126-129` · `GetMaxDigitWidth` `:131-136` · `GetTextHeight` `:138-143`
- `GetTextWidth` `:145-151` · `GetGlyphBox` `:154-169`, tail `:162-168`
- `GetFont` `:171-174` · `LoadFont` `:176-183` · `ResolveTypeface` `:185-208`, throw `:205-207`
- `TryMatchSystemFont` `:210-222`, exact-name check `:213`
- `CalculateMaxDigitWidth` `:224-233` · `AddTypeface` `:235-242` · `PointsToPixels` `:244`
- `FontEntry` `:246-267` · `MetricId` `:269-305` · `FontStyleKind` `:307-313`
- `SkiaSharpFontBootstrap.cs:35` names CarlitoBare; `:76` wires it into `CreateDefault()`

**Tests (as they stand at `1b41cadd`; spec 27 replaces the first three):**

- `XLibur.Tests/Graphics/FontTests.cs` — V1 only (V1 has no test project of its own;
  `docs/font-architecture.md:150`)
- `XLibur.Fonts.SixLabors.Tests/SixLaborsFontEngineTests.cs` — 431 lines, **deleted by 27 task 2**
- `XLibur.Fonts.SkiaSharp.Tests/SkiaSharpFontEngineTests.cs` — 434 lines, **deleted by 27 task 2**
- After 27: `XLibur.Fonts.Conformance/{FontEngineConformance,GoldenMetrics,ConformanceFixture,
  ConformanceFont}.cs`, linked into three thin per-package suites
  (`SixLaborsConformanceTests.cs`, `SkiaSharpConformanceTests.cs`,
  `XLibur.Tests/Graphics/SixLaborsV1ConformanceTests.cs`).
- **No test in the tree asserts the terminal fallback branch of any adapter**, before or after 27 —
  27's scope is metric agreement, not error modes. That gap is this spec's task 2.

**Benchmarks:** there is **no autofit benchmark**. `grep -rln AdjustToContents XLibur.Benchmarks`
returns nothing. Task 6 has to write one.

## File structure

```
XLibur/Graphics/IXLTypefaceSource.cs        new — the adapter-facing mechanism port (4 members)
XLibur/Graphics/XLTypeface.cs               new — opaque face handle + XLTypefaceMetrics + XLTypefaceStyle
XLibur/Graphics/XLFontFaceKey.cs            new — the MetricId struct, moved into core
XLibur/Graphics/XLFontMetrics.cs            new — IXLFontEngine over IXLTypefaceSource; all the policy

XLibur/Graphics/IXLFontEngine.cs            UNCHANGED
XLibur/Graphics/GraphicEngineFontAdapter.cs UNCHANGED
XLibur/Graphics/DefaultGraphicEngine.cs     UNCHANGED
XLibur/Excel/XLWorkbook.cs                  UNCHANGED

XLibur.Fonts.SixLabors.V1/DefaultFontEngine.cs      modified — 306 -> ~150; becomes IXLTypefaceSource
XLibur.Fonts.SixLabors/SixLaborsFontEngine.cs       modified — 277 -> ~140; becomes IXLTypefaceSource
XLibur.Fonts.SkiaSharp/SkiaSharpFontEngine.cs       modified — 314 -> ~170; becomes IXLTypefaceSource
XLibur.Fonts.SkiaSharp/SkiaSharpFontBootstrap.cs    modified — wraps the source in XLFontMetrics
XLibur.Fonts.SixLabors.V1/SixLaborsV1FontBootstrap.cs modified — same

CHANGELOG.md                                modified — task 2's behaviour decision
docs/font-architecture.md                   modified — the two-level port
```

No file deleted. No public type removed.

## The design

### Two ports, two audiences

```
                     core: XLColumn, XLRow, VmlDrawingPartWriter, XLHelper (6 read sites)
                                            |
                                   IXLFontEngine          <- UNCHANGED. public. 5 methods.
                                            |                 The seam for someone replacing the
                                            |                 whole of font policy, and for the
                                            |                 pre-split IXLGraphicEngine adapter.
                          +-----------------+------------------+
                          |                                    |
                   XLFontMetrics                     (third-party IXLFontEngine)
              caching / fallback / units /
              digit scan / glyph assembly
                          |
                  IXLTypefaceSource          <- NEW. public. 4 members. Mechanism only.
                          |                     The seam a fourth shipped adapter sits at.
       +------------------+------------------+
       |                  |                  |
  V1 (SixLabors 1)   v2 (SixLabors 2)   SkiaSharp
```

**A fourth adapter sits at `IXLTypefaceSource`.** That is the answer to "where do I plug in a new font
library." `IXLFontEngine` remains available and remains supported, but implementing it means opting
out of the shared policy — the right choice only for someone who disagrees with the policy itself.

### `IXLTypefaceSource` — 4 members

```csharp
using System;

namespace XLibur.Graphics;

/// <summary>
/// Resolves font families to typefaces and reads raw advances from them. This is the
/// <b>mechanism</b> half of the font seam: everything a font library can do that XLibur cannot do
/// for itself, and nothing else.
/// </summary>
/// <remarks>
/// <para>
/// Caching, the fallback chain, point-to-pixel conversion, the 0-9 digit scan and glyph-box
/// assembly all live above this interface in <see cref="XLFontMetrics"/>. An implementation must
/// not do any of them.
/// </para>
/// <para>
/// <b>Units.</b> Both advance methods return <i>font design units</i>, to be divided by
/// <see cref="XLTypefaceMetrics.UnitsPerEm"/>. No kerning is applied.
/// </para>
/// <para>
/// <b>Thread safety.</b> Implementations must be safe for concurrent calls from multiple threads.
/// <see cref="XLFontMetrics"/> calls <see cref="TryResolve"/> from inside
/// <c>ConcurrentDictionary.GetOrAdd</c>, which may invoke it concurrently for the same key.
/// </para>
/// <para>
/// <b>Failure.</b> <see cref="TryResolve"/> returns <c>false</c> for a family it cannot supply. It
/// must not throw for an unknown family and must not substitute a different family — deciding what
/// happens next is <see cref="XLFontMetrics"/>'s job, and it cannot decide if a substitution has
/// already happened silently.
/// </para>
/// </remarks>
public interface IXLTypefaceSource
{
    /// <summary>Resolve a family name and style to a typeface, or return <c>false</c>.</summary>
    bool TryResolve(string familyName, XLTypefaceStyle style, out XLTypeface? typeface);

    /// <summary>Read <c>unitsPerEm</c>, ascent and descent from a resolved typeface.</summary>
    XLTypefaceMetrics GetMetrics(XLTypeface typeface);

    /// <summary>Total advance of <paramref name="text"/> in font design units, no kerning.</summary>
    double GetTextAdvance(XLTypeface typeface, string text);

    /// <summary>Total advance of one grapheme cluster in font design units.</summary>
    double GetClusterAdvance(XLTypeface typeface, ReadOnlySpan<int> codePoints);
}
```

Supporting types:

```csharp
/// <summary>An opaque, adapter-owned handle to a resolved typeface.</summary>
/// <remarks>
/// The handle is created by the adapter and passed back to it unread. <see cref="XLFontMetrics"/>
/// caches it and never inspects <see cref="Handle"/>. Skia's existing <c>FontEntry</c>
/// (SkiaSharpFontEngine.cs:246-267) is this type in adapter-private form.
/// </remarks>
public sealed class XLTypeface
{
    public XLTypeface(object handle) => Handle = handle;
    public object Handle { get; }
}

/// <summary>
/// The three vertical numbers XLibur needs from a typeface, in font design units.
/// </summary>
/// <remarks>
/// <b>Sign convention: ascent and descent are both positive.</b> This is stated because the
/// adapters currently disagree — SixLabors' <c>VerticalMetrics.Descender</c> is negative and both
/// SixLabors adapters negate it (V1 :140, v2 :126), while Skia stores it positive
/// (SkiaSharpFontEngine.cs:182) and its height formula adds where the others subtract
/// (Skia :142 vs V1 :154). Same answer, opposite conventions, neither written down. This tag is
/// where it gets written down.
/// </remarks>
public readonly record struct XLTypefaceMetrics(int UnitsPerEm, double AscentFu, double DescentFu);

/// <summary>Regular / Bold / Italic / BoldItalic. The style half of the cache key.</summary>
public enum XLTypefaceStyle { Regular, Bold, Italic, BoldItalic }
```

### `XLFontMetrics` — the policy, once

```csharp
/// <summary>
/// Implements <see cref="IXLFontEngine"/> over an <see cref="IXLTypefaceSource"/>. Owns every
/// decision that is not a font-library call: the fallback chain, both caches, point-to-pixel
/// conversion, the 0-9 digit scan and glyph-box assembly.
/// </summary>
/// <remarks>
/// <b>Thread safety.</b> Safe for concurrent use. Both caches are <see cref="ConcurrentDictionary{TKey,TValue}"/>
/// and the <c>GetOrAdd</c> factories are hoisted into fields so no closure is allocated per lookup —
/// the shape all three adapters independently arrived at before this type existed.
/// </remarks>
public sealed class XLFontMetrics : IXLFontEngine
{
    private const double FontMetricSize = 16d;

    private readonly IXLTypefaceSource _source;
    private readonly string _fallbackFamily;
    private readonly string? _lastResortFamily;

    private readonly ConcurrentDictionary<XLFontFaceKey, ResolvedFace> _faces = new();
    private readonly Func<XLFontFaceKey, ResolvedFace> _resolve;

    private readonly ConcurrentDictionary<XLFontFaceKey, double> _maxDigitWidths = new();
    private readonly Func<XLFontFaceKey, double> _calculateMaxDigitWidth;

    /// <param name="source">The font library.</param>
    /// <param name="fallbackFamily">Tried when the requested family does not resolve.</param>
    /// <param name="lastResortFamily">
    /// Tried when the fallback does not resolve either. When <c>null</c>, an unresolvable font
    /// throws — see the remarks on <see cref="XLFontResolutionException"/>.
    /// </param>
    public XLFontMetrics(IXLTypefaceSource source, string fallbackFamily, string? lastResortFamily = null);

    private readonly record struct ResolvedFace(XLTypeface Typeface, XLTypefaceMetrics Metrics);
}
```

`ResolvedFace` caches the metrics alongside the handle so `GetMetrics` is called once per
`(family, style)` rather than once per measurement — matching what Skia's `FontEntry` already does
and what V1/v2 currently pay a property read for on every call (`V1:203-207`, `v2:185-189`).

`XLFontFaceKey` is `MetricId` moved into core:

```csharp
/// <summary>
/// Cache key for a resolved typeface: family name plus style. Moved into core from three
/// byte-identical private copies (V1 :277-305, v2 :248-276, Skia :269-305).
/// </summary>
public readonly record struct XLFontFaceKey(string Name, XLTypefaceStyle Style)
{
    public static XLFontFaceKey From(IXLFontBase font) => new(font.FontName, StyleOf(font));
    // ...
}
```

### Constant answer to design question 2 — is one cache key enough?

Yes. All three adapters key on exactly `(FontName, Style)` where `Style` is a four-value enum
(`V1:279-283`, `v2:250-254`, `Skia:271-275`), and all three hash it identically
(`(Name.GetHashCode() * 397) ^ (int)Style`). Skia's `MetricId` additionally *derives* an `SkStyle`
property (`:281-287`), but that is a projection of `Style`, not part of the key. One key works for
all three, unchanged.

What **does** change is cache **lifetime**. Today each cache lives on the adapter instance, so it is
per-engine. After this spec it lives on the `XLFontMetrics` instance — which is also per-engine,
because each package's bootstrap constructs one `XLFontMetrics` wrapping one source. The lifetime is
therefore unchanged, and task 5 must confirm it: an `XLFontMetrics` must never be shared across
`IXLTypefaceSource` instances, because the key does not include the source. Enforce it structurally —
`XLFontMetrics` takes its source in the constructor and exposes no setter.

The cache **value** changes from `Font` / `FontEntry` to `ResolvedFace`, one extra reference hop per
lookup. Task 6 measures it.

### The unit contract, and the precision risk it carries

`IXLTypefaceSource` returns advances in font design units. Skia already produces that directly —
`GetGlyphBox` at `:157` constructs `new SKFont(entry.Typeface, entry.UnitsPerEm)` for exactly this
reason. SixLabors does not: both SixLabors adapters measure at `FontMetricSize = 16f` and scale
(`V1:161-166`, `v2:147-152`).

**This is a real risk and it is the reason task 1 gates on the golden table.** A `float` advance
measured at size 16 and multiplied by `unitsPerEm / 16` is not bit-identical to the integral font-unit
advance. Column widths round to integer pixels (`XLColumn.cs:195`, `XLHelper.cs:489`), so a boundary
case can flip by one pixel.

**Escape hatch, decided in advance:** if the golden table moves in task 1, do **not** re-baseline it.
Change the contract instead — have `GetTextAdvance` return the advance at `FontMetricSize`, exactly as
today, and document `FontMetricSize` on the interface. That keeps the arithmetic bit-identical at the
cost of a less clean unit. The clean unit is worth having; it is not worth a metric change.

## Global constraints

- **Warnings are errors** (`TreatWarningsAsErrors=true`); nullable enabled. Every new type above must
  be null-annotated — `out XLTypeface? typeface` in particular.
- **Branch per spec; never commit to main.** Commit prefixes `refactor:` / `fix:` / `test:` / `perf:`.
- **No compound shell commands** (`&&`, `||`, `;`) in agent tool calls.
- **Do not use `sed -i` on tracked files.** `.gitattributes` checks out CRLF and Git Bash's `sed -i`
  rewrites the file as LF, turning a one-line change into a whole-file diff. Use the Edit/Write tools.
  Verify with `git diff --numstat`: a file whose changed-line count is near its total line count was
  rewritten, not edited. This spec deletes ~450 lines across three files, so the numstat check is not
  optional — it is the only way to tell a real deletion from a line-ending rewrite.
- **Do not upgrade SixLabors.Fonts.** Newer versions carry the Six Labors Split License. This is why
  three font packages exist and it constrains this spec directly: V1 stays on 1.0.1, v2 stays a
  separate package, and neither may be collapsed into the other however similar their adapters become
  after this refactor.
- Test filtering uses `--treenode-filter`, never `--filter`. Exit 5 = invalid option; exit 8 = zero
  tests matched. **Never filter at solution level** — name the `.csproj`.
- Pass `-f net10.0` while iterating; run without it before opening the PR.
- Build: `dotnet build XLibur/XLibur.csproj -c Release -v q`
- Tests use **TUnit**: `await Assert.That(actual).IsEqualTo(expected)`. Assertions are awaitable and a
  missing `await` silently passes. `[Test]`, `[Arguments(...)]`, `[MethodDataSource(...)]`. The suite
  is serial (`[assembly: NotInParallel]`).

## Work plan

| # | Task | Size | Gate |
|---|---|---|---|
| 0 | Confirm spec 27's conformance suite is in place and green | XS | Suite runs against all three adapters; golden table exists |
| 1 | `IXLTypefaceSource`, `XLFontMetrics`, V1 converted | M | Conformance suite green, golden table unmoved |
| 2 | Decide and implement the unified fallback; record it | S | New tests pin the terminal branch in all three packages; CHANGELOG entry |
| 3 | Convert v2 | S | Conformance suite green |
| 4 | Convert Skia | S | Conformance suite green |
| 5 | Delete the duplicated policy; confirm the counts | S | Duplication measurement re-run; grep gates pass |
| 6 | Benchmark the autofit path | S | Within noise of baseline, or reverted |

Tasks 3 and 4 are independent of each other once 2 lands. Everything else is ordered.

---

### Task 0 — Refuse to start without spec 27

This spec relocates metric computation across three adapters that ship to three different NuGet
packages. The only thing that can catch a metric change is a test that measures the same font through
more than one adapter and compares. **No such test exists at `1b41cadd`.**

**Files:**
- Create: none. This task writes no code.

- [ ] **Step 1: Confirm 27's task 3 has landed — not just been written**

The spec file exists at `docs/specs/27-font-conformance-suite.md`. That is not the gate. The gate is
its **task 3**, the golden metric table, which is what turns "all three adapters run the same tests"
into "all three adapters assert the same numbers." 27's own Conflicts section says it: *"34 must not
start before task 3 lands. Tasks 4 and 5 can run concurrently with 34's early work; tasks 1–3
cannot."*

Run: `ls XLibur.Fonts.Conformance/`
Expected: `FontEngineConformance.cs`, `GoldenMetrics.cs`, `ConformanceFixture.cs`,
`ConformanceFont.cs`. If `GoldenMetrics.cs` is missing, 27 task 3 has not landed and **this spec stops
here.**

Run: `git log --oneline -30`
Confirm the commits for 27 tasks 1–3 are merged.

27's task 4 (V1 joining the conformance run) is not strictly required to start, but this spec
converts V1 first, so **starting before 27 task 4 means converting the one adapter the conformance
module does not yet cover.** Wait for task 4 unless there is a reason not to; if you do start early,
task 1's gate falls back to `XLibur.Tests/Graphics/FontTests.cs` alone, which is weaker.

- [ ] **Step 2: Run the conformance suite against all three adapters, green**

Run: `dotnet test XLibur.Fonts.SixLabors.Tests/XLibur.Fonts.SixLabors.Tests.csproj -f net10.0`
Run: `dotnet test XLibur.Fonts.SkiaSharp.Tests/XLibur.Fonts.SkiaSharp.Tests.csproj -f net10.0`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/*Conformance*/*"`
Expected: PASS on all three. The third exits 8 if 27 task 4 has not landed — that is the "V1 not yet
covered" case from step 1, not a failure.

- [ ] **Step 3: Record the golden table's current values**

Copy the constants from `XLibur.Fonts.Conformance/GoldenMetrics.cs` into this spec as a Results
appendix, verbatim, **before touching any source**, together with the per-metric tolerances 27
attaches to each row. Tasks 1–5 compare against this copy, not against a re-read of the file — a
re-read cannot catch a value this spec's own work changed.

- [ ] **Step 4: Verify the gate bites**

In `XLibur.Fonts.SkiaSharp/SkiaSharpFontEngine.cs:150`, temporarily change
`advance / FontMetricSize * font.FontSize` to `advance / FontMetricSize * font.FontSize * 1.001`.

Run: `dotnet test XLibur.Fonts.SkiaSharp.Tests/XLibur.Fonts.SkiaSharp.Tests.csproj -f net10.0`
Expected: **FAIL** on a golden-table row. Restore the line.

0.1% is deliberately below 27's stated per-metric tolerances for cross-rasteriser drift. If the suite
does not notice it, check whether the tolerance on the width rows is wider than 0.1%; if it is,
perturb by 2% instead and record that this spec's effective resolution is the tolerance, not zero.
**If even 2% passes, the suite cannot gate this spec** — report it to spec 27 and stop.

- [ ] **Step 5: Commit the appendix**

```bash
git add docs/specs/34-font-port-split.md
git commit -m 'docs(specs): pin spec 27 golden metrics as spec 34 baseline (spec 34 task 0)'
```

---

### Task 1 — The two ports, and V1 on the new one

**Files:**
- Create: `XLibur/Graphics/IXLTypefaceSource.cs`
- Create: `XLibur/Graphics/XLTypeface.cs`
- Create: `XLibur/Graphics/XLFontFaceKey.cs`
- Create: `XLibur/Graphics/XLFontMetrics.cs`
- Modify: `XLibur.Fonts.SixLabors.V1/DefaultFontEngine.cs`
- Modify: `XLibur.Fonts.SixLabors.V1/SixLaborsV1FontBootstrap.cs:33`

**Interfaces:**
- Produces: `IXLTypefaceSource` (4 members), `XLTypeface`, `XLTypefaceMetrics`, `XLTypefaceStyle`,
  `XLFontFaceKey`, `XLFontMetrics : IXLFontEngine`.
- Unchanged: `IXLFontEngine`, `GraphicEngineFontAdapter`, `DefaultGraphicEngine`, `XLWorkbook`.

- [ ] **Step 1: Write the four core types**

Exactly as given under **The design**. Points to get right, each of which is a thing an adapter got
wrong somewhere today:

- `TryResolve` must be documented as **not throwing** for an unknown family and **not substituting**.
  Skia's `MatchFamily` substitutes; the adapter's exact-name check at `:213` moves into Skia's
  `TryResolve` in task 4 and is the reason this sentence is in the contract.
- `XLTypefaceMetrics` documents both ascent and descent as **positive**. Both SixLabors adapters must
  negate `VerticalMetrics.Descender` inside `GetMetrics`, as they do today at `V1:140` / `v2:126`.
- Thread safety is stated on both `IXLTypefaceSource` and `XLFontMetrics`.
- `XLFontMetrics` hoists both `GetOrAdd` factories into readonly fields. All three adapters already
  do this (`V1:36,43`; `v2:25,28`; `Skia:26,29`); do not lose it.

- [ ] **Step 2: Move the policy into `XLFontMetrics`, verbatim**

Take the bodies from V1, which is the reference implementation, and change only what the new types
force:

```csharp
    public double GetMaxDigitWidth(IXLFontBase font, double dpiX)
    {
        var key = XLFontFaceKey.From(font);
        var maxDigitWidth = _maxDigitWidths.GetOrAdd(key, _calculateMaxDigitWidth);
        return PointsToPixels(maxDigitWidth * font.FontSize, dpiX);
    }

    public double GetTextHeight(IXLFontBase font, double dpiY)
    {
        var face = GetFace(font);
        return PointsToPixels(
            (face.Metrics.AscentFu + 2 * face.Metrics.DescentFu) * font.FontSize
            / face.Metrics.UnitsPerEm, dpiY);
    }

    public double GetDescent(IXLFontBase font, double dpiY)
        => PointsToPixels(GetFace(font) is var f
            ? f.Metrics.DescentFu * font.FontSize / f.Metrics.UnitsPerEm : 0, dpiY);

    private static double PointsToPixels(double points, double dpi) => points / 72d * dpi;
```

Note the sign flip in `GetTextHeight`: V1 writes `Ascender - 2 * Descender` with a negative
`Descender` (`:154`), Skia writes `AscentFu + 2 * DescentFu` with a positive one (`:142`). Under the
new positive-descent contract, Skia's form is the correct one. This is the single most likely place
to introduce a sign bug; the golden table is what catches it.

The 0–9 digit scan (`V1:250-273`) and the glyph-box assembly tail (`V1:194-200`) move verbatim, with
the per-codepoint advance coming from `_source.GetClusterAdvance`.

- [ ] **Step 3: Convert V1 to `IXLTypefaceSource`**

`DefaultFontEngine` keeps its class name, its namespace, its public constructors, `Instance`,
`CreateOnlyWithFonts` and `CreateWithFontsAndSystemFonts` — all four signatures are public API and
all four must still return an `IXLFontEngine`. Internally the class splits:

```csharp
/// <summary>
/// SixLabors.Fonts 1.0.1 typeface source. Mechanism only: family resolution and raw advances.
/// Caching, fallback, unit conversion and the digit scan live in <see cref="XLFontMetrics"/>.
/// </summary>
internal sealed class SixLaborsV1TypefaceSource : IXLTypefaceSource
{
    private readonly Lazy<IReadOnlyFontCollection> _fontCollection;

    public bool TryResolve(string familyName, XLTypefaceStyle style, out XLTypeface? typeface)
    {
        if (!_fontCollection.Value.TryGet(familyName, out var family))
        {
            typeface = null;
            return false;
        }

        typeface = new XLTypeface(family.CreateFont(16f, ToFontStyle(style)));
        return true;
    }

    public XLTypefaceMetrics GetMetrics(XLTypeface typeface)
    {
        var metrics = ((Font)typeface.Handle).FontMetrics;
        return new XLTypefaceMetrics(
            metrics.UnitsPerEm,
            metrics.VerticalMetrics.Ascender,
            -metrics.VerticalMetrics.Descender);   // contract: positive
    }
    // GetTextAdvance, GetClusterAdvance
}
```

and the public `DefaultFontEngine` becomes a thin composition:

```csharp
public class DefaultFontEngine : IXLFontEngine
{
    private const string EmbeddedFontName = "CarlitoBare";
    private readonly XLFontMetrics _metrics;

    public DefaultFontEngine(string fallbackFont)
    {
        // ... build the collection exactly as today, AddEmbeddedFont included
        _metrics = new XLFontMetrics(source, fallbackFont, EmbeddedFontName);
    }

    public double GetTextHeight(IXLFontBase font, double dpiY) => _metrics.GetTextHeight(font, dpiY);
    // ... four more one-line forwards
}
```

**V1's terminal behaviour is preserved exactly** by passing `EmbeddedFontName` as
`lastResortFamily` — `AddEmbeddedFont` runs in every constructor, so CarlitoBare is always present
and the chain still cannot fail. Nothing about V1 changes observably in this task. That is the point:
task 1 must be a pure refactor so that task 2's decision is isolated.

Update `SixLaborsV1FontBootstrap.cs:33` only if `Instance` moved; it should not have to.

- [ ] **Step 4: Run the golden table**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/*Font*/*"`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS, with **every value in task 0's golden table identical**.

If a value moved, do not re-baseline. Diagnose in this order: (a) the descent sign, (b) the advance
unit — apply the escape hatch under **The unit contract** and re-run, (c) `GetOrAdd` returning a
different face because the key changed. Record which it was.

- [ ] **Step 5: Commit**

```bash
git add XLibur/Graphics/IXLTypefaceSource.cs XLibur/Graphics/XLTypeface.cs XLibur/Graphics/XLFontFaceKey.cs XLibur/Graphics/XLFontMetrics.cs XLibur.Fonts.SixLabors.V1/DefaultFontEngine.cs
git commit -m 'refactor(fonts): split the font port into mechanism and policy, V1 first (spec 34 task 1)'
```

---

### Task 2 — Decide the fallback, and say so out loud

**This is a behaviour decision, not a mechanical move.** Defining fallback once means picking one
answer for what "font not found" means. Each of the current answers is observable from identical user
code.

**Files:**
- Modify: `XLibur/Graphics/XLFontMetrics.cs`
- Create: `XLibur/Graphics/XLFontResolutionException.cs` (if option B or C is chosen)
- Modify: `CHANGELOG.md`
- Modify: `docs/font-architecture.md`
- Create: `XLibur.Fonts.Conformance/FallbackConformance.cs` — the shared shape
- Modify: `XLibur.Fonts.SixLabors.Tests/SixLaborsConformanceTests.cs`,
  `XLibur.Fonts.SkiaSharp.Tests/SkiaSharpConformanceTests.cs`,
  `XLibur.Tests/Graphics/SixLaborsV1ConformanceTests.cs` — each supplies its own expected terminal
  behaviour

**Do not add these to `SixLaborsFontEngineTests.cs` / `SkiaSharpFontEngineTests.cs`.** Spec 27 task 2
deletes both files. Add them alongside 27's conformance module, following its
`ConformanceFixture` pattern — the terminal behaviour is exactly the kind of thing that *should* be
one shared test body parameterised by what each package expects, and 27 did not cover it because its
scope is metric agreement, not error modes.

- [ ] **Step 1: Establish what each package does today, with a test**

Before deciding anything, pin the current terminal behaviour. Nothing does today. The three bodies
below are written per-package for clarity; fold them into one conformance body with an expectation on
the fixture if that reads better against 27's actual shape.

```csharp
/// <summary>
/// What happens when neither the requested family nor the fallback resolves. Spec 34 task 2
/// unifies this; these tests are what make the change visible rather than silent.
/// </summary>
[Test]
public async Task Unresolvable_font_falls_through_to_the_embedded_family()
{
    // V1: never throws. AddEmbeddedFont runs in every constructor.
    var engine = new DefaultFontEngine("NoSuchFallbackFamily12345");
    var unknown = new DummyFont("NoSuchFamily98765", 11);
    var carlito = new DummyFont("CarlitoBare", 11);

    await Assert.That(engine.GetTextWidth("Test", unknown, 96))
        .IsEqualTo(engine.GetTextWidth("Test", carlito, 96));
}
```

`DummyFont` is the pre-27 stub; after 27 it is `ConformanceFont`
(`XLibur.Fonts.Conformance/ConformanceFont.cs`). Use whichever is live.

And, in `XLibur.Fonts.SixLabors.Tests`:

```csharp
[Test]
public async Task Unresolvable_font_throws_when_no_embedded_family_is_configured()
{
    var engine = new SixLaborsFontEngine("NoSuchFallbackFamily12345");
    var unknown = new DummyFont("NoSuchFamily98765", 11);

    await Assert.That(() => engine.GetTextWidth("Test", unknown, 96))
        .Throws<InvalidOperationException>();
}
```

and the same shape in `XLibur.Fonts.SkiaSharp.Tests`, plus one asserting that
`SkiaSharpFontBootstrap.CreateDefault()` does **not** throw, because it configures CarlitoBare
(`SkiaSharpFontBootstrap.cs:35,76`). That last one guards the zero-config path every `XLibur.Bundle`
consumer gets, and it is the single test in this spec whose failure would be a shipped outage.

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/*Conformance*/*"`
Run: `dotnet test XLibur.Fonts.SixLabors.Tests/XLibur.Fonts.SixLabors.Tests.csproj -f net10.0`
Run: `dotnet test XLibur.Fonts.SkiaSharp.Tests/XLibur.Fonts.SkiaSharp.Tests.csproj -f net10.0`
Expected: PASS. These pass on unmodified behaviour — they are characterization, not a change.

**If any of these fails, the premise in "Why this spec exists" is wrong and that is a real result.**
Record what actually happened and revise the table in this spec before continuing.

- [ ] **Step 2: Choose one of A, B or C — and write the choice into this spec**

| | Behaviour when nothing resolves | Who breaks | Cost |
|---|---|---|---|
| **A** | Always fall through to an embedded metric-only family | v2 and Skia users who rely on the throw to catch a missing font in CI | `XLibur.Fonts.SixLabors` must start shipping four `.ttf` assets it deliberately does not ship today (`docs/font-architecture.md:45`) |
| **B** | Always throw `XLFontResolutionException` | **V1 users** — every unknown font that silently measured as CarlitoBare now throws | V1 must stop passing a last-resort family, a visible regression for its zero-config path |
| **C** | Throw unless a last-resort family is configured; each package configures what it configures today | Nobody | The port states one rule and the packages differ by *declared configuration* rather than by three hand-written chains |

**Recommendation: C.** It is what the evidence actually supports. The three chains turn out to be one
chain with two configuration values (`lastResortFamily` present or absent) plus one piece of library
compensation that belongs below the seam. C makes that structure explicit, keeps every package's
observable behaviour, and — the part that matters — moves the divergence from "undocumented accident
in three files" to "one parameter, documented on one constructor, pinned by step 1's tests."

C is not a dodge. The *port* now states exactly one rule; what differs is a value each package
chooses and this spec records:

| Package | `lastResortFamily` | Terminal behaviour |
|---|---|---|
| V1, every constructor | `"CarlitoBare"` | falls through |
| v2, every constructor except the 4-arg one | `null` | throws |
| v2, 4-arg constructor | caller's `embeddedFontName` | falls through |
| Skia, direct construction | `null` | throws |
| Skia, `CreateDefault()` — the zero-config path | `"CarlitoBare"` | falls through |

If B is chosen instead, it is a **breaking change for `XLibur.Fonts.SixLabors.V1`** and needs the
CHANGELOG entry below plus a migration note. If A is chosen, it is a **breaking change for
`XLibur.Fonts.SixLabors` and `XLibur.Fonts.SkiaSharp`** and additionally requires adding font assets
to the v2 package, which is a packaging change outside this spec's file list — re-scope before
starting.

Whichever is chosen, **write the decision and its reasoning into this spec's Results section.** A
spec that leaves this open has not done its job.

- [ ] **Step 3: Implement the chain in `XLFontMetrics`, once**

```csharp
    private ResolvedFace Resolve(XLFontFaceKey key)
    {
        if (_source.TryResolve(key.Name, key.Style, out var typeface)
            || _source.TryResolve(_fallbackFamily, key.Style, out typeface)
            || (_lastResortFamily is not null
                && _source.TryResolve(_lastResortFamily, key.Style, out typeface)))
        {
            return new ResolvedFace(typeface!, _source.GetMetrics(typeface!));
        }

        throw new XLFontResolutionException(key.Name, _fallbackFamily, _lastResortFamily);
    }
```

Three lines of chain replacing `V1:219-230`, `v2:201-219` and `Skia:185-208`.

`XLFontResolutionException` derives from `InvalidOperationException` so step 1's tests keep passing
and so no existing `catch (InvalidOperationException)` in a consumer breaks. Give it `FamilyName`,
`FallbackFamily` and `LastResortFamily` properties — the current message concatenates them into a
string a caller cannot inspect (`v2:212-215`, `Skia:205-207`).

- [ ] **Step 4: Document the error mode on `IXLFontEngine`**

`IXLFontEngine` gains **no member and no signature change** — but its doc comments gain an
`<exception>` tag on each of the five methods naming `XLFontResolutionException`. This is the one
edit to that file the spec permits, and it is documentation only:

```csharp
    /// <exception cref="XLFontResolutionException">
    /// The font could not be resolved and the engine has no last-resort family configured.
    /// Implementations that always resolve never throw this.
    /// </exception>
```

Adding the tag is the whole point of the spec's title: the error mode existed, in three different
forms, and nowhere in the contract.

- [ ] **Step 5: CHANGELOG**

Under `## Unreleased` in `CHANGELOG.md:15`, following the repo's existing heading style
(`### ⚠️ Breaking Changes`, first used at `:29`). For option C, which breaks nothing, use a plain
entry instead:

```markdown
### 🏗️ Architecture

- **The font seam splits into mechanism and policy.** `IXLFontEngine` is unchanged — the same five
  methods, the same signatures, and any third-party implementation keeps working. What changed is
  below it: the three shipped adapters now implement a 4-member `IXLTypefaceSource` (resolve a
  family, read `unitsPerEm`/ascent/descent, advance a string, advance a cluster) and the caching,
  fallback chain, unit conversion, digit scan and glyph-box assembly they each carried a copy of now
  live once in `XLFontMetrics`. 145 lines were byte-identical across all three engines; they are now
  written once. A new font backend costs a typeface lookup rather than 300 lines.
- **"Font not found" is now a stated error mode.** `IXLFontEngine` documented none, and the three
  engines disagreed: V1 always fell through to embedded CarlitoBare metrics, while v2 and directly
  constructed SkiaSharp engines threw. Both behaviours are preserved and are now a declared
  constructor parameter, `lastResortFamily`, rather than three hand-written chains. The exception is
  `XLFontResolutionException`, which derives from `InvalidOperationException`, so existing catch
  blocks are unaffected.
```

If B or A was chosen, that second bullet moves under `### ⚠️ Breaking Changes` and must name which
packages change behaviour and what a user does about it.

- [ ] **Step 6: Update `docs/font-architecture.md`**

Add the two-level port diagram to the "Interface Separation" section (`:13-25`) and state which seam
a new backend sits at. Correct `:15` — "five methods" is still right for `IXLFontEngine`, but the
document should now say there are two ports and which is which.

- [ ] **Step 7: Commit**

```bash
git add XLibur/Graphics/XLFontMetrics.cs XLibur/Graphics/XLFontResolutionException.cs XLibur/Graphics/IXLFontEngine.cs CHANGELOG.md docs/font-architecture.md
git commit -m 'fix(fonts): define what "font not found" means, once (spec 34 task 2)'
```

---

### Task 3 — Convert the SixLabors v2 adapter

**Files:**
- Modify: `XLibur.Fonts.SixLabors/SixLaborsFontEngine.cs` — 277 lines, target ~140

**Interfaces:**
- Produces: `SixLaborsTypefaceSource : IXLTypefaceSource` (internal).
- Unchanged public surface: `SixLaborsFontEngine(string)`, `SixLaborsFontEngine(string, string,
  string[], Assembly)`, `CreateOnlyWithFonts`, `CreateWithFontsAndSystemFonts`.

- [ ] **Step 1: Extract the source**

The v2 source is V1's source with `AddEmbeddedFont` removed. `TryResolve` is `_fontCollection.Value
.TryGet(familyName, out var family)` — one level, no fallback, no throw. `GetMetrics`,
`GetTextAdvance` and `GetClusterAdvance` are byte-identical to V1's, because both packages call the
same SixLabors API surface; the only reason they are two files is the license pin.

**Do not merge the two packages.** They must stay separate — see Global constraints. The near-total
duplication between `SixLaborsV1TypefaceSource` and `SixLaborsTypefaceSource` after this task is
deliberate and required: it is the smallest possible pair of files that can compile against two
mutually incompatible assembly versions. Note it in a comment on both so a later reader does not
"tidy" them together.

- [ ] **Step 2: Compose**

```csharp
    public SixLaborsFontEngine(string fallbackFont)
    {
        // ... unchanged validation and collection setup
        _metrics = new XLFontMetrics(source, fallbackFont, lastResortFamily: null);
    }

    public SixLaborsFontEngine(string fallbackFont, string embeddedFontName,
        string[] embeddedFontResourcePaths, Assembly assembly)
    {
        // ... unchanged
        _metrics = new XLFontMetrics(source, fallbackFont, embeddedFontName);
    }
```

`_embeddedFontName` (`:22`) disappears as a field; it becomes the third constructor argument.

- [ ] **Step 3: Run**

Run: `dotnet build XLibur.Fonts.SixLabors/XLibur.Fonts.SixLabors.csproj -c Release -v q`
Run: `dotnet test XLibur.Fonts.SixLabors.Tests/XLibur.Fonts.SixLabors.Tests.csproj -f net10.0`
Expected: PASS, golden table unmoved, and task 2 step 1's throw test still throwing.

- [ ] **Step 4: Commit**

```bash
git add XLibur.Fonts.SixLabors/SixLaborsFontEngine.cs
git commit -m 'refactor(fonts): SixLabors v2 becomes a typeface source (spec 34 task 3)'
```

---

### Task 4 — Convert the SkiaSharp adapter

**Files:**
- Modify: `XLibur.Fonts.SkiaSharp/SkiaSharpFontEngine.cs` — 314 lines, target ~170
- Modify: `XLibur.Fonts.SkiaSharp/SkiaSharpFontBootstrap.cs:76`

- [ ] **Step 1: Move the library compensation below the seam, and only that**

Skia's `TryResolve` absorbs `_streamFonts` lookup (`:188-189`), `TryMatchSystemFont` (`:210-222`
including the exact-name check at `:213`) and `AddTypeface` (`:235-242`). It does **not** absorb the
fallback family or the embedded family — those are levels 3 and 4 of `ResolveTypeface` (`:196-203`)
and they move up into `XLFontMetrics`.

```csharp
    public bool TryResolve(string familyName, XLTypefaceStyle style, out XLTypeface? typeface)
    {
        if (_streamFonts.TryGetValue(familyName, out var streamTypeface))
        {
            typeface = Wrap(streamTypeface);
            return true;
        }

        // SKFontManager.MatchFamily substitutes rather than failing, so an inexact match is a
        // silent wrong font. IXLTypefaceSource.TryResolve must not substitute — deciding what
        // happens next belongs to XLFontMetrics, which cannot decide if this already lied.
        if (_useSystemFonts && TryMatchSystemFont(familyName, ToSkStyle(style), out var systemTypeface))
        {
            typeface = Wrap(systemTypeface);
            return true;
        }

        typeface = null;
        return false;
    }
```

`FontEntry` (`:246-267`) is superseded by `XLTypeface` + `XLTypefaceMetrics`. Its four members map
one-to-one: `Typeface` → `XLTypeface.Handle`, and `UnitsPerEm` / `AscentFu` / `DescentFu` →
`XLTypefaceMetrics`, with the sign convention already matching (`:182` stores both positive).
`FontStyleKind` (`:307-313`) is superseded by `XLTypefaceStyle`.

**`SKTypeface` disposal:** `TryMatchSystemFont` disposes a rejected match (`:219`). An accepted one is
held by the cache for the engine's lifetime, as today. Do not add disposal to the new path — the
current behaviour is to leak the accepted typefaces for the process lifetime, which is correct for a
cache and is not this spec's to change.

- [ ] **Step 2: Compose, preserving `CreateDefault()`**

`SkiaSharpFontBootstrap.CreateDefault()` (`:65-77`) must keep configuring CarlitoBare as
`lastResortFamily`. This is the zero-config path that every consumer of `XLibur.Bundle` gets; if it
starts throwing, `new XLWorkbook()` breaks on any machine missing the requested font.

Run: `dotnet test XLibur.Fonts.SkiaSharp.Tests/XLibur.Fonts.SkiaSharp.Tests.csproj -f net10.0 --treenode-filter "/*/*/SkiaSharpFontBootstrapTests/*"`
Expected: PASS.

- [ ] **Step 3: Run everything**

Run: `dotnet build XLibur.Fonts.SkiaSharp/XLibur.Fonts.SkiaSharp.csproj -c Release -v q`
Run: `dotnet test XLibur.Fonts.SkiaSharp.Tests/XLibur.Fonts.SkiaSharp.Tests.csproj -f net10.0`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS, golden table unmoved.

- [ ] **Step 4: Commit**

```bash
git add XLibur.Fonts.SkiaSharp/SkiaSharpFontEngine.cs XLibur.Fonts.SkiaSharp/SkiaSharpFontBootstrap.cs
git commit -m 'refactor(fonts): SkiaSharp becomes a typeface source (spec 34 task 4)'
```

---

### Task 5 — Delete the duplicated policy and confirm the numbers

Tasks 1, 3 and 4 each left their adapter composing `XLFontMetrics`. This task removes what is now
dead and measures the result against the numbers this spec opened with.

**Files:**
- Modify: all three adapter files
- Modify: `docs/specs/34-font-port-split.md` — Results

- [ ] **Step 1: Delete, and check the numstat**

Remove from all three: `PointsToPixels`, `MetricId` (including Skia's `FontStyleKind`), both
`ConcurrentDictionary` fields and their hoisted factories, `GetMaxDigitWidth`, `GetTextHeight`,
`GetDescent`'s private overload, `CalculateMaxDigitWidth`, and the glyph-box assembly tail.

Run: `git diff --numstat`
Expected: three rows with large deletion counts and **small insertion counts**. A row whose insertions
are close to the file's total line count means `sed -i` or an editor rewrote the file's line endings —
`.gitattributes` checks these out CRLF. Revert and redo with the Edit tool.

- [ ] **Step 2: Re-measure the duplication**

```bash
diff --unchanged-group-format='%=' --old-group-format='' --new-group-format='' --changed-group-format='' XLibur.Fonts.SixLabors.V1/DefaultFontEngine.cs XLibur.Fonts.SixLabors/SixLaborsFontEngine.cs | wc -l
```

Baseline at `1b41cadd`: 229. Expected after: **under 90.** The residue is the two SixLabors sources,
which must stay separate for the license pin, plus each engine's unchanged public constructors and
factory methods.

```bash
wc -l XLibur.Fonts.SixLabors.V1/DefaultFontEngine.cs XLibur.Fonts.SixLabors/SixLaborsFontEngine.cs XLibur.Fonts.SkiaSharp/SkiaSharpFontEngine.cs
```

Baseline: 306 / 277 / 314 = 897. Expected: **under 500 total.**

- [ ] **Step 3: Grep gates**

```bash
grep -rn "PointsToPixels" XLibur.Fonts.SixLabors.V1 XLibur.Fonts.SixLabors XLibur.Fonts.SkiaSharp --include=*.cs
```
Expected: no output.

```bash
grep -rn "struct MetricId" XLibur.Fonts.SixLabors.V1 XLibur.Fonts.SixLabors XLibur.Fonts.SkiaSharp --include=*.cs
```
Expected: no output.

```bash
grep -rn "ConcurrentDictionary" XLibur.Fonts.SixLabors.V1 XLibur.Fonts.SixLabors XLibur.Fonts.SkiaSharp --include=*.cs
```
Expected: no output — every cache now lives in `XLFontMetrics`. Skia's plain
`Dictionary<string, SKTypeface> _streamFonts` (`:20`) is not a cache and stays; it is populated once
in the constructor and read-only afterwards.

- [ ] **Step 4: Confirm the cache lifetime did not change**

Run: `grep -n "new XLFontMetrics" XLibur.Fonts.SixLabors.V1 XLibur.Fonts.SixLabors XLibur.Fonts.SkiaSharp -r`
Every construction site must pass a freshly built source. If any `XLFontMetrics` is shared across two
sources, the `(Name, Style)` key is ambiguous and metrics will cross-contaminate between engines.
`XLFontMetrics` exposes no setter for its source, so this should be structurally impossible — confirm
it is.

- [ ] **Step 5: Full suite, both frameworks**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj`
Run: `dotnet test XLibur.Fonts.SixLabors.Tests/XLibur.Fonts.SixLabors.Tests.csproj`
Run: `dotnet test XLibur.Fonts.SkiaSharp.Tests/XLibur.Fonts.SkiaSharp.Tests.csproj`
Expected: PASS on net8.0 and net10.0.

- [ ] **Step 6: Commit**

```bash
git add XLibur.Fonts.SixLabors.V1/DefaultFontEngine.cs XLibur.Fonts.SixLabors/SixLaborsFontEngine.cs XLibur.Fonts.SkiaSharp/SkiaSharpFontEngine.cs docs/specs/34-font-port-split.md
git commit -m 'refactor(fonts): delete the duplicated font policy from all three adapters (spec 34 task 5)'
```

---

### Task 6 — Benchmark the autofit path

Text measurement is on the autofit path, and this spec adds one interface hop between the cache and
the font library plus one reference hop inside the cached value. Both are per-lookup, and a lookup
happens per distinct `(family, style)` — but `GetTextWidth` and `GetClusterAdvance` are per **cell**
and per **grapheme**, and those now cross the new interface.

There is no autofit benchmark in the repository today, so this task writes one first.

**Files:**
- Create: `XLibur.Benchmarks/AutofitBenchmarks.cs`

- [ ] **Step 1: Write the benchmark**

```csharp
/// <summary>
/// Column and row autofit over a grid of mixed-length strings. Spec 34 moves text measurement's
/// caching across a new interface boundary; this is what would show it.
/// </summary>
/// <remarks>
/// Two shapes deliberately: many distinct fonts exercises the resolution cache, one font exercises
/// the per-cell measurement path. The second is the one at risk — the first pays the new
/// indirection once per (family, style), the second once per cell.
/// </remarks>
[MemoryDiagnoser]
public class AutofitBenchmarks
{
    private const int Rows = 5_000;
    private const int Columns = 10;

    [Benchmark] public void AdjustColumnsToContents_SingleFont() { /* ... */ }
    [Benchmark] public void AdjustColumnsToContents_ManyFonts() { /* ... */ }
    [Benchmark] public void AdjustRowsToContents_SingleFont() { /* ... */ }
}
```

Model it on `XLibur.Benchmarks/SheetGeometryBenchmarks.cs`, which is the closest existing shape and
already references `XLibur.Fonts.SixLabors.V1`.

- [ ] **Step 2: Measure the merge-base**

```
git stash
dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- --filter '*Autofit*'
```

Run it **three times**. Record all three.

- [ ] **Step 3: Measure the branch, three runs**

Same command.

- [ ] **Step 4: Compare medians**

**A single pair of runs proves nothing on this machine.** Spec 19 measured the noise floor at
**4.5–9%** on the write benchmarks (`19-benchmark-hotspot-survey.md:63`, `:873`, `:1197`) and records
that it is worse elsewhere; run-to-run swings approaching 40% have been observed on this hardware for
workloads that touch the filesystem. Autofit does not, so expect the lower band — but establish the
floor rather than assuming it: the three merge-base runs from step 2 *are* the noise measurement.
Compare median-of-three to median-of-three and treat anything inside the measured floor, or ±15%,
whichever is larger, as noise.

Allocation is the reliable signal here, not time. Expected: **zero change in allocated bytes.** The
new types are one `XLTypeface` and one `ResolvedFace` per `(family, style)`, both cached; the
per-cell path allocates nothing new. **An allocation delta on the per-cell path is a defect**, most
likely a `GetOrAdd` factory that was not hoisted into a field (`V1:36,43` is the pattern) and is now
capturing a closure per call.

- [ ] **Step 5: Decision rule**

If the per-cell single-font benchmark regresses more than 15% on median-of-three, **revert task 5's
deletions and re-scope.** The precedent is spec 21, which implemented its task 1, measured a 60%
regression on the enclosing walk, and reverted — and whose headline finding is that the interface
dispatch it was written to remove had already been devirtualised by dynamic PGO and was never the
cost. The same could be true here in reverse: a new interface call that PGO cannot devirtualise
because three implementations are live in the process.

Record the numbers either way. A measured "no change" is the result this task exists to produce.

- [ ] **Step 6: Commit**

```bash
git add XLibur.Benchmarks/AutofitBenchmarks.cs docs/specs/34-font-port-split.md
git commit -m 'perf(fonts): benchmark the autofit path across the new font seam (spec 34 task 6)'
```

---

## Acceptance criteria

1. **`IXLFontEngine`'s five signatures are unchanged.** Gate:
   `git diff 1b41cadd -- XLibur/Graphics/IXLFontEngine.cs` shows only added `<exception>` doc
   comments — no changed, added or removed member.
2. **`GraphicEngineFontAdapter.cs` is byte-identical to `1b41cadd`.** Gate:
   `git diff --numstat 1b41cadd -- XLibur/Graphics/GraphicEngineFontAdapter.cs` returns nothing.
3. **`IXLTypefaceSource` has at most 4 members.** Gate: count of `;`-terminated member declarations in
   `XLibur/Graphics/IXLTypefaceSource.cs` is ≤ 4.
4. **All three adapters implement `IXLTypefaceSource` and none implements policy.** Gate:
   `grep -rn "PointsToPixels\|struct MetricId\|ConcurrentDictionary" XLibur.Fonts.SixLabors.V1 XLibur.Fonts.SixLabors XLibur.Fonts.SkiaSharp --include=*.cs`
   returns nothing.
5. **Duplication is down.** Gate: the V1↔v2 `diff` common-line count is **< 90** (baseline 229), and
   `wc -l` over the three adapter files totals **< 500** (baseline 897).
6. **Every value in spec 27's golden metric table is unchanged**, or the one exception is named in
   Results with the reason and a CHANGELOG entry.
7. **The terminal fallback branch is pinned in all three packages** by a test that reaches it — at
   least one test per package that constructs an engine whose fallback family does not resolve, plus
   one asserting `SkiaSharpFontBootstrap.CreateDefault()` does not throw. Baseline: zero such tests
   exist. The tests live alongside spec 27's conformance module, not in the two files 27 deletes.
8. **The fallback decision is recorded** in this spec's Results section, naming which of A/B/C was
   chosen, which packages change behaviour, and the CHANGELOG entry that says so.
9. **`XLFontResolutionException` derives from `InvalidOperationException`** so existing catch blocks
   are unaffected. Gate: `grep -n "class XLFontResolutionException" XLibur/Graphics/XLFontResolutionException.cs`
   shows `: InvalidOperationException`.
10. **No public constructor, factory method or type of any font package is removed or resignatured.**
    Gate: every `public static IXLFontEngine CreateOnlyWithFonts` and
    `CreateWithFontsAndSystemFonts` still present in all three files, unchanged.
11. **Registration is untouched.** Gate:
    `git diff --numstat 1b41cadd -- XLibur/Graphics/DefaultFontEngineProbe.cs XLibur/Excel/XLWorkbook.cs XLibur/Excel/LoadOptions.cs`
    returns nothing.
12. **Autofit allocation is unchanged**; time within ±15% on median-of-three, or task 5 reverted per
    task 6's decision rule. Numbers in Results either way.
13. **Full suite green on net8.0 and net10.0** across `XLibur.Tests`,
    `XLibur.Fonts.SixLabors.Tests` and `XLibur.Fonts.SkiaSharp.Tests`.
14. **No `sed -i` line-ending damage.** Gate: no file in `git diff --numstat` against the merge-base
    has an insertion count within 10% of its total line count unless it is a new file.

## Conflicts

- **[Spec 27](27-font-conformance-suite.md) is a hard prerequisite, not a conflict.** Both specs
  state the dependency; task 0 enforces it. The two share test files — 27 creates
  `XLibur.Fonts.Conformance/` and rewrites all three font test projects, and this spec's task 2
  adds a fallback body to that module. **Run 27 through at least task 3 first**, ideally task 4; this
  spec then adds to its module rather than competing with it. 27's tasks 4 and 5 may run concurrently
  with this spec's tasks 1–2.

  **One correction to 27.** Its Conflicts section says *"34 changes `IXLFontEngine`."* It does not.
  `IXLFontEngine`'s five signatures are unchanged and criterion 1 gates that; the only edit to that
  file is adding `<exception>` doc tags in task 2. 27 was written before this spec existed and its
  sentence describes an earlier sketch. The dependency it derives from that sentence still holds —
  metric computation does move across all three adapters — so nothing about the ordering changes.

- **Everything else: disjoint.** Checked against every spec now in `docs/specs/` (01–30, 34). No spec
  other than 27 touches `XLibur/Graphics/`, `XLibur.Fonts.SixLabors.V1/`, `XLibur.Fonts.SixLabors/`
  or `XLibur.Fonts.SkiaSharp/`. Verified by grep across all spec files:

  - **Spec 28** mentions `XLibur.Tests/Graphics/FontTests.cs` once (`28-single-style-decoder.md:100`),
    as an example of a file that is *not* in its scope. No overlap.
  - **Specs 20 and 23** work in `XLibur/Excel/Styles/`. `IXLFontBase` is read by this spec and
    modified by neither.
  - **Spec 24** works in `WorksheetElementReader.cs` and `XLWorkbook_Load.cs`; **spec 18 task 5** in
    `LoadWorksheetElements`. This spec reads `XLWorkbook_Load.cs:948` as evidence and edits nothing
    there.
  - **Spec 25** is in `XLCellFormulaShifter*.cs`; **22** in chart IO; **15/16/17** in DrawingML and
    `PictureWriter.cs`; **26**, **29**, **30** in the grid axis, write-path resolvers and array
    application respectively.

  The one shared file is `CHANGELOG.md`, which every spec eventually edits. Conflicts there are
  textual and trivial.

- **Adjacent but deliberately untouched, three of them:**
  - `XLibur/XLHelper.cs:444` holds a fourth copy of `PointsToPixels`. Core's copy has core callers
    (`XLRow.cs:299` via `PixelsToPoints`, and the width conversions at `:489`/`:518`); consolidating
    core's unit helpers is separate, smaller work.
  - The four `CarlitoBare-*.ttf` under `XLibur/Graphics/Fonts/` are tracked and embedded by nothing —
    spec 27's closing note establishes this and explicitly declines to delete them. This spec
    declines too, for the same reason: it is unrelated to the seam and would make criterion 11
    ("registration is untouched") harder to read.
  - `XLibur.Fonts.SixLabors.V1/ModuleInit.cs` retains V1's module initializer for backward
    compatibility (`docs/font-architecture.md:168`). Untouched.
