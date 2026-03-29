# Font Engine Architecture

## Problem

XLibur originally embedded all font operations (text measurement, glyph metrics, digit width calculation) directly in `DefaultGraphicEngine`, tightly coupled to SixLabors.Fonts 1.0.1. This created three problems:

1. **License lock-in**: SixLabors.Fonts 2.x changed to the Six Labors Split License (commercial restrictions for organizations over $1M revenue). The project is pinned to 1.0.1 (Apache 2.0) and cannot upgrade.
2. **No font library choice**: Users who needed SixLabors.Fonts 2.x features, or wanted an entirely different font library, had no path — replacing font handling required re-implementing image handling too, since `IXLGraphicEngine` conflated both concerns.
3. **DLL version conflict**: Shipping a v2 SixLabors package alongside the core (v1) would cause NuGet to unify to v2 at runtime, silently running v1-compiled code against the v2 assembly.

## Design

### Interface Separation

Font operations were extracted from `IXLGraphicEngine` into a new `IXLFontEngine` interface with five methods:

```
GetTextHeight    — text height in pixels (EMHeight + descent)
GetTextWidth     — text width in pixels (no padding)
GetMaxDigitWidth — widest 0-9 digit in pixels (OOXML column width unit)
GetDescent       — font descent in pixels (positive value)
GetGlyphBox      — glyph bounding box for a grapheme cluster
```

`IXLGraphicEngine` was left unchanged (no breaking change). `DefaultGraphicEngine` implements both interfaces, delegating font methods to an `IXLFontEngine` instance.

### Package Structure

```
XLibur (core)
├── IXLFontEngine               — font engine interface
├── IXLGraphicEngine            — graphic engine interface (unchanged)
├── DefaultGraphicEngine        — image handling + font delegation
├── FontEngineProvider          — static registry for font engine factories
├── GraphicEngineFontAdapter    — wraps legacy IXLGraphicEngine as IXLFontEngine
└── LoadOptions.FontEngine      — injection point

XLibur.Fonts.SixLabors.V1
├── DefaultFontEngine           — SixLabors.Fonts 1.0.1 implementation
├── CarlitoBare embedded fonts  — Calibri-compatible metric-only fallback
└── ModuleInit                  — auto-registers with FontEngineProvider

XLibur.Fonts.SixLabors
├── SixLaborsFontEngine         — SixLabors.Fonts 2.1.3 implementation
└── (no embedded fonts)         — uses system fonts or stream-provided fonts
```

The core assembly has **zero dependency on SixLabors.Fonts**. Each font package depends only on XLibur core and its respective SixLabors.Fonts version. No circular dependencies.

### Assembly Discovery

The core cannot reference V1 (that would create a circular dependency). Instead:

1. **V1 registers itself** via `[ModuleInitializer]` which sets `FontEngineProvider.FromFallbackFont` and `FontEngineProvider.FromStreams` factory delegates.
2. **FontEngineProvider lazily discovers V1** when first needed — it probes `AppContext.BaseDirectory` and `Assembly.Location` for `XLibur.Fonts.SixLabors.V1.dll`, loads it via `AssemblyLoadContext.Default.LoadFromAssemblyPath`, then calls `RuntimeHelpers.RunModuleConstructor` to trigger the module initializer.
3. If V1 is not present and no `FontEngine` is configured, an `InvalidOperationException` is thrown with a clear message.

### Injection and Resolution

Users configure font engines through `LoadOptions`:

```csharp
// Use the default (V1 auto-discovered)
using var wb = new XLWorkbook();

// Use SixLabors.Fonts 2.x
var options = new LoadOptions
{
    FontEngine = new SixLaborsFontEngine("Microsoft Sans Serif")
};
using var wb = new XLWorkbook(options);

// Use stream-based fonts (Blazor, serverless)
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

The release workflow triggers on any of these tag patterns. `--skip-duplicate` on NuGet push means unchanged packages are not re-published.

### Bold/Italic Font Style

`MetricId` (the font cache key) encodes font name + style (Regular/Bold/Italic/BoldItalic). `LoadFont` passes the style to `FontFamily.CreateFont(size, style)` so bold and italic variants get correct metrics. This was a pre-existing bug in the original `DefaultGraphicEngine` that was fixed during the extraction.

### Test Strategy

| Test project | References | SixLabors version | What it tests |
|---|---|---|---|
| XLibur.Tests | XLibur + V1 | 1.0.1 only | DefaultFontEngine, DefaultGraphicEngine, font injection via LoadOptions |
| XLibur.Fonts.SixLabors.Tests | XLibur.Fonts.SixLabors (transitive XLibur) | 2.1.3 only | SixLaborsFontEngine with embedded test fonts (no system fonts needed for CI) |

The test projects never reference both V1 and the v2 package, eliminating DLL version conflicts entirely. The SixLabors test project uses stream-based `TestFontA.ttf` as its default engine so tests pass on Linux CI without system fonts.

## Key Decisions and Rationale

**Why not remove SixLabors.Fonts from the ecosystem entirely?**
SixLabors.Fonts 1.0.1 is battle-tested, Apache 2.0 licensed, and provides accurate metrics for hundreds of fonts. A from-scratch font parser would be a massive effort with questionable benefit.

**Why a separate V1 package instead of keeping SixLabors in core?**
If both V1 (1.0.1) and V2 (2.1.3) packages exist and a consumer installs both, NuGet unifies to 2.1.3. With SixLabors in core, *every* consumer who adds the V2 package gets the unified version — the V1 code runs against V2 silently. By extracting V1, consumers choose one or the other. The core is version-agnostic.

**Why module initializer + RuntimeHelpers instead of a direct reference?**
XLibur core defines `IXLFontEngine`. V1 implements it. V1 must reference XLibur for the interface. If XLibur also referenced V1, that's a circular dependency MSBuild won't allow. The module initializer pattern breaks the cycle: V1 registers itself at load time, XLibur discovers it lazily.

**Why `RuntimeHelpers.RunModuleConstructor`?**
`Assembly.Load` and `AssemblyLoadContext.LoadFromAssemblyPath` load assembly metadata but don't execute the module initializer (the CLR's `.cctor` for `<Module>`). `RuntimeHelpers.RunModuleConstructor` explicitly invokes it, ensuring the factory delegates are registered before first use.

**Why `GraphicEngineFontAdapter`?**
A user who implemented `IXLGraphicEngine` before `IXLFontEngine` existed has the same 5 font method signatures on their type — they just don't implement the new interface. The adapter wraps their graphic engine and delegates the font calls, preserving their measurement behavior without requiring them to change code.
