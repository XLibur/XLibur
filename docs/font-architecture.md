# Font Engine Architecture

## Problem

XLibur originally embedded all font operations (text measurement, glyph metrics, digit width calculation) directly in `DefaultGraphicEngine`, tightly coupled to SixLabors.Fonts 1.0.1. This created three problems:

1. **License lock-in**: SixLabors.Fonts 2.x changed to the Six Labors Split License (commercial restrictions for organizations over $1M revenue). The project is pinned to 1.0.1 (Apache 2.0) and cannot upgrade.
2. **No font library choice**: Users who needed SixLabors.Fonts 2.x features, or wanted an entirely different font library, had no path — replacing font handling required re-implementing image handling too, since `IXLGraphicEngine` conflated both concerns.
3. **DLL version conflict**: Shipping a v2 SixLabors package alongside the core (v1) would cause NuGet to unify to v2 at runtime, silently running v1-compiled code against the v2 assembly.

## Design

### Interface Separation

Font operations were extracted from `IXLGraphicEngine` into a new `IXLFontEngine` interface with five methods:

```text
GetTextHeight    — text height in pixels (EMHeight + descent)
GetTextWidth     — text width in pixels (no padding)
GetMaxDigitWidth — widest 0-9 digit in pixels (OOXML column width unit)
GetDescent       — font descent in pixels (positive value)
GetGlyphBox      — glyph bounding box for a grapheme cluster
```

`IXLGraphicEngine` was left unchanged (no breaking change). `DefaultGraphicEngine` implements both interfaces, delegating font methods to an `IXLFontEngine` instance.

### Package Structure

```text
XLibur (core)
├── IXLFontEngine               — font engine interface
├── IXLGraphicEngine            — graphic engine interface (unchanged)
├── DefaultGraphicEngine        — image handling + font delegation (requires IXLFontEngine)
├── GraphicEngineFontAdapter    — wraps legacy IXLGraphicEngine as IXLFontEngine
└── LoadOptions.FontEngine      — injection point

XLibur.Fonts.SixLabors.V1
├── DefaultFontEngine            — SixLabors.Fonts 1.0.1 implementation
├── SixLaborsV1FontBootstrap     — explicit registration API
├── CarlitoBare embedded fonts   — Calibri-compatible metric-only fallback
└── ModuleInit                   — auto-registers if assembly is loaded

XLibur.Fonts.SixLabors
├── SixLaborsFontEngine          — SixLabors.Fonts 2.1.3 implementation
└── (no embedded fonts)          — uses system fonts or stream-provided fonts

XLibur.Fonts.SkiaSharp
├── SkiaSharpFontEngine          — SkiaSharp (MIT) implementation
└── (no embedded fonts)          — uses system fonts or stream-provided fonts
```

The core assembly has **zero dependency on any font library**. Each font package depends only on XLibur core and its respective font library. No circular dependencies.

### SkiaSharp engine (license-friendly alternative)

`XLibur.Fonts.SkiaSharp` provides `SkiaSharpFontEngine`, a third `IXLFontEngine` implementation backed by [SkiaSharp](https://github.com/mono/SkiaSharp). It mirrors the public surface of `SixLaborsFontEngine` (fallback-name constructor, embedded-font constructor, `CreateOnlyWithFonts`, `CreateWithFontsAndSystemFonts`) and is selected the same way:

```csharp
using var wb = new XLWorkbook(new LoadOptions { FontEngine = new SkiaSharpFontEngine("Arial") });
```

**Why add it:** SkiaSharp is MIT-licensed (wrapping BSD-licensed Skia), with no commercial-revenue restrictions — unlike SixLabors.Fonts 2.x (Six Labors Split License). A measurement spike comparing the two libraries on Carlito (Calibri-compatible), TestFontA, and TestFontB found **0% metric drift** across width, descent, height, and max-digit-width: SkiaSharp's `SKFontMetrics.Ascent/Descent` resolves to the same vertical font units as SixLabors' `VerticalMetrics`, so parity comes for free without parsing the OS/2 table by hand.

**Trade-off — native dependency:** Unlike the pure-managed SixLabors.Fonts engines, SkiaSharp wraps native Skia and ships per-platform native binaries. The package includes `SkiaSharp.NativeAssets.Linux.NoDependencies`, which needs no system `fontconfig`/`freetype`, so stream-loaded fonts work in headless and serverless environments. Consumers who do not reference this package take on no native dependency.

**Implementation note:** SkiaSharp has no `FontCollection.TryGet(name)` equivalent, so the engine keeps its own `name → SKTypeface` dictionary for stream-loaded fonts and falls through to `SKFontManager.Default` for system fonts. Glyph metrics are read via `SKFont.MeasureText` and `SKFontMetrics`; the `SKTypeface.GetGlyph`/`SKFont.ContainsGlyph` lookup APIs proved unreliable for stream-loaded typefaces and are not used.

### Registration

The core cannot reference V1 (that would create a circular dependency). Instead, each font engine package provides an explicit public bootstrap API that consumers call at application startup:

```csharp
// In Program.cs or application startup — register the V1 font engine:
SixLaborsV1FontBootstrap.Register();
```

This sets `LoadOptions.DefaultFontEngine` — the single global default. `DefaultGraphicEngine` requires an `IXLFontEngine` as a constructor parameter; it has no factory-based or string-based constructors. The V1 package also includes a `[ModuleInitializer]` that calls `Register()` automatically when the assembly is loaded — but explicit registration is preferred for clarity and predictability.

If no font engine is registered and no `FontEngine` is configured, an `InvalidOperationException` is thrown with a clear message explaining what to do.

### Injection and Resolution

Users configure font engines through `LoadOptions`:

```csharp
// Register V1 at startup, then use XLibur normally
SixLaborsV1FontBootstrap.Register();
using var wb = new XLWorkbook();

// Or use SixLabors.Fonts 2.x explicitly
var options = new LoadOptions
{
    FontEngine = new SixLaborsFontEngine("Microsoft Sans Serif")
};
using var wb = new XLWorkbook(options);

// Or use stream-based fonts (Blazor, serverless)
var engine = SixLaborsFontEngine.CreateOnlyWithFonts(fontStream);
var options = new LoadOptions { FontEngine = engine };
```

Resolution order in `XLWorkbook` constructor:

```
FontEngine =
    loadOptions.FontEngine                         // explicit per-workbook
    ?? LoadOptions.DefaultFontEngine               // global static default
    ?? (GraphicEngine as IXLFontEngine)            // if graphic engine implements it
    ?? new GraphicEngineFontAdapter(GraphicEngine)  // wrap legacy graphic engine
```

The `GraphicEngineFontAdapter` ensures backward compatibility: a custom `IXLGraphicEngine` that doesn't implement `IXLFontEngine` still has its font methods used (since `IXLGraphicEngine` has the same 5 method signatures).

When a `FontEngine` is explicitly provided but no `GraphicEngine` is set, the workbook creates `new DefaultGraphicEngine(fontEngine)` — avoiding the V1 assembly load for the graphic engine's font needs.

### Versioning

Each package has independent MinVer versioning with distinct tag prefixes:

| Package | Tag prefix | Example |
|---------|-----------|---------|
| XLibur | `v` | `v0.106.0` |
| XLibur.Fonts.SixLabors.V1 | `fonts-sixlabors-1-v` | `fonts-sixlabors-1-v1.0.0` |
| XLibur.Fonts.SixLabors | `fonts-sixlabors-v` | `fonts-sixlabors-v0.1.0` |
| XLibur.Fonts.SkiaSharp | `fonts-skiasharp-v` | `fonts-skiasharp-v0.1.0` |

The release workflow triggers on any of these tag patterns. `--skip-duplicate` on NuGet push means unchanged packages are not re-published.

### Bold/Italic Font Style

`MetricId` (the font cache key) encodes font name + style (Regular/Bold/Italic/BoldItalic). `LoadFont` passes the style to `FontFamily.CreateFont(size, style)` so bold and italic variants get correct metrics. This was a pre-existing bug in the original `DefaultGraphicEngine` that was fixed during the extraction.

### Test Strategy

| Test project | References | SixLabors version | What it tests |
|---|---|---|---|
| XLibur.Tests | XLibur + V1 | 1.0.1 only | DefaultFontEngine, DefaultGraphicEngine, font injection via LoadOptions |
| XLibur.Fonts.SixLabors.Tests | XLibur.Fonts.SixLabors (transitive XLibur) | 2.1.3 only | SixLaborsFontEngine with embedded test fonts (no system fonts needed for CI) |
| XLibur.Fonts.SkiaSharp.Tests | XLibur.Fonts.SkiaSharp (transitive XLibur) | SkiaSharp 3.x | SkiaSharpFontEngine with embedded test fonts (no system fonts needed for CI) |

The test projects never reference both V1 and the v2 package, eliminating DLL version conflicts entirely. The SixLabors test project uses stream-based `TestFontA.ttf` as its default engine so tests pass on Linux CI without system fonts.

## Key Decisions and Rationale

**Why not remove SixLabors.Fonts from the ecosystem entirely?**
SixLabors.Fonts 1.0.1 is battle-tested, Apache 2.0 licensed, and provides accurate metrics for hundreds of fonts. A from-scratch font parser would be a massive effort with questionable benefit.

**Why a separate V1 package instead of keeping SixLabors in core?**
If both V1 (1.0.1) and V2 (2.1.3) packages exist and a consumer installs both, NuGet unifies to 2.1.3. With SixLabors in core, *every* consumer who adds the V2 package gets the unified version — the V1 code runs against V2 silently. By extracting V1, consumers choose one or the other. The core is version-agnostic.

**Why explicit bootstrap instead of automatic assembly scanning?**
An earlier design had the core probing for V1's DLL at runtime via `AssemblyLoadContext.LoadFromAssemblyPath` and `RuntimeHelpers.RunModuleConstructor`. This was fragile — it relied on `AppContext.BaseDirectory` being correct, assembly probing paths, and forcing module initializers. The explicit `SixLaborsV1FontBootstrap.Register()` call is one line, obvious, stable, and gives consumers full control over initialization order. The module initializer remains as a convenience fallback but is not the primary design.

**Why does `DefaultGraphicEngine` only accept `IXLFontEngine` (no string constructor)?**
The original `DefaultGraphicEngine(string fallbackFont)` constructor, `Instance` singleton, and `CreateOnlyWithFonts`/`CreateWithFontsAndSystemFonts` factory methods all internally created a `DefaultFontEngine` — which lives in the V1 assembly. Since the core can't reference V1 (circular dependency), those constructors required an internal factory delegation system (`DefaultFontEngineFactory`) bridged via `InternalsVisibleTo`. This created two parallel global registration paths that could get out of sync. Removing the factory-dependent constructors eliminates the entire internal factory system. Consumers who need a `DefaultGraphicEngine` pass in the font engine explicitly: `new DefaultGraphicEngine(new DefaultFontEngine("Arial"))`. Most consumers don't create `DefaultGraphicEngine` directly — they call `SixLaborsV1FontBootstrap.Register()` and use `new XLWorkbook()`.

**Breaking changes from shipped API:**
- `DefaultGraphicEngine(string)` — removed (use `DefaultGraphicEngine(IXLFontEngine)`)
- `DefaultGraphicEngine.Instance` — removed (use `SixLaborsV1FontBootstrap.Register()` + `new XLWorkbook()`)
- `DefaultGraphicEngine.CreateOnlyWithFonts(...)` — removed (use `DefaultFontEngine.CreateOnlyWithFonts(...)` then `new DefaultGraphicEngine(fontEngine)`)
- `DefaultGraphicEngine.CreateWithFontsAndSystemFonts(...)` — removed (same pattern)

**Why `GraphicEngineFontAdapter`?**
A user who implemented `IXLGraphicEngine` before `IXLFontEngine` existed has the same 5 font method signatures on their type — they just don't implement the new interface. The adapter wraps their graphic engine and delegates the font calls, preserving their measurement behavior without requiring them to change code.
