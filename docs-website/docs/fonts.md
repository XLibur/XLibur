---
id: fonts
title: Fonts and Font Engines
sidebar_label: Fonts
description: Choose between the SkiaSharp and SixLabors font engines, register them globally or per workbook, and load fonts from streams in headless environments.
---

# Fonts and Font Engines

XLibur needs to *measure* text — how wide is `"Total revenue"` in 11pt Calibri? — to size
columns with `AdjustToContents()`, calculate row heights, and lay out glyphs. That measurement
requires a font library.

Unlike ClosedXML, which bakes SixLabors.Fonts into its core assembly, **XLibur's core assembly
has no font dependency at all**. The font engine ships as a separate, swappable package, so you
can pick a font library whose licence suits you — and library authors who never measure text do
not inherit a font dependency they did not ask for.

## The short version

Install `XLibur.Bundle` and everything works with no configuration:

```sh
dotnet add package XLibur.Bundle
```

```csharp
using var workbook = new XLWorkbook();   // font engine resolved automatically
workbook.Worksheets.Add("Data").Columns().AdjustToContents();
```

`XLibur.Bundle` = `XLibur` + `XLibur.Fonts.SkiaSharp`. The SkiaSharp engine (MIT) is
auto-registered by the core the first time you create a workbook — there is no startup call.

:::warning
If you install the bare `XLibur` package with **no** font engine, creating a workbook throws an
`InvalidOperationException` telling you which package to add. This is deliberate: it is how the
core stays font-library-agnostic.
:::

## Available engines

| Package | Font library | Licence | Notes |
|---|---|---|---|
| `XLibur.Fonts.SkiaSharp` | SkiaSharp | MIT | **Default.** Auto-registers. Ships native binaries. |
| `XLibur.Fonts.SixLabors.V1` | SixLabors.Fonts 1.x | Apache 2.0 | Pure-managed. Matches ClosedXML 0.105's engine exactly. |
| `XLibur.Fonts.SixLabors` | SixLabors.Fonts 2.x | Six Labors Split License | Commercial restrictions above $1M revenue. |

Measurement parity between SkiaSharp and SixLabors was verified at **0% metric drift** across
width, descent, height, and max-digit-width — switching engines does not change column widths.

### Choosing between them

- **SkiaSharp (default)** — MIT, no revenue restrictions, resolves system fonts, and includes
  an embedded metric-only Calibri-compatible fallback so measurement works in headless and
  serverless environments with no fonts installed. Trade-off: it wraps native Skia and ships
  per-platform native binaries.
- **SixLabors.Fonts 1.x** — pure managed, no native dependency. Choose this if you are
  migrating from ClosedXML and want byte-identical behaviour, or if native binaries are a
  problem in your deployment.
- **SixLabors.Fonts 2.x** — only if you specifically need 2.x features and the Split License
  is acceptable for your organisation.

## Registering an engine

Resolution happens in three layers, checked cheapest-first:

1. **Per workbook** — `LoadOptions.FontEngine`
2. **Global explicit** — `LoadOptions.DefaultFontEngine`, usually set by a package bootstrap
3. **Auto-registered default** — the core reflectively locates `XLibur.Fonts.SkiaSharp`

Anything set explicitly always beats the auto-registered default.

### Global registration at startup

Call the bootstrap once, before any workbook is created:

```csharp
using XLibur.Fonts.SixLabors.V1;

// In Program.cs
SixLaborsV1FontBootstrap.Register();

// Every workbook from here on uses SixLabors.Fonts 1.x
using var workbook = new XLWorkbook();
```

```csharp
using XLibur.Fonts.SkiaSharp;

// Force the default engine at a specific point in startup
SkiaSharpFontBootstrap.Register();
```

Both bootstraps use `??=`, so the **first** registration wins and repeat calls are no-ops. If
you need to override a registration that already happened, assign directly:

```csharp
using XLibur.Excel;
using XLibur.Fonts.SixLabors;

LoadOptions.DefaultFontEngine = new SixLaborsFontEngine("Arial");
```

### Per-workbook registration

Pass an engine through `LoadOptions`. This overrides whatever global default is in place, for
this workbook only:

```csharp
using XLibur.Excel;
using XLibur.Fonts.SkiaSharp;

var options = new LoadOptions
{
    FontEngine = new SkiaSharpFontEngine("Arial"),
};

using var workbook = new XLWorkbook(options);
```

The same options object works when loading an existing file:

```csharp
using var workbook = new XLWorkbook("Report.xlsx", options);
```

## The fallback font

Every engine constructor takes a **fallback font name** — the font used when a workbook asks
for a typeface that is not installed:

```csharp
new SkiaSharpFontEngine("Arial");
new SkiaSharpFontEngine("Microsoft Sans Serif");   // the default engine's choice
new SixLaborsFontEngine("Segoe UI");
new DefaultFontEngine("Microsoft Sans Serif");     // SixLabors 1.x
```

Pick something metrically close to what your workbooks actually use. A wildly different
fallback produces wildly different column widths.

## Headless and containerised environments

Docker images, Azure Functions, and AWS Lambda typically have no system fonts. Two options:

### Rely on the embedded fallback

`XLibur.Fonts.SkiaSharp` and `XLibur.Fonts.SixLabors.V1` both embed *CarlitoBare* — a
metric-only, Calibri-compatible font. Because it matches Calibri's metrics, column widths come
out correct for the default Excel font even with nothing installed on the box. This is the
zero-config path and needs no code.

```csharp
// Works in an empty container with XLibur.Bundle installed
using var workbook = new XLWorkbook();
workbook.Worksheets.Add("Data").Columns().AdjustToContents();
```

The SkiaSharp package also references `SkiaSharp.NativeAssets.Linux.NoDependencies`, so it
needs no system `fontconfig` or `freetype`.

### Supply fonts from streams

When the workbook uses a specific corporate typeface, load the font files yourself. Both
engines expose two factories:

```csharp
using XLibur.Excel;
using XLibur.Fonts.SkiaSharp;

await using var fallback = File.OpenRead("fonts/Inter-Regular.ttf");
await using var bold = File.OpenRead("fonts/Inter-Bold.ttf");

// Only these fonts — the system font collection is ignored entirely
var engine = SkiaSharpFontEngine.CreateOnlyWithFonts(fallback, bold);

var options = new LoadOptions { FontEngine = engine };
using var workbook = new XLWorkbook(options);
```

```csharp
// These fonts first, then fall through to whatever the system has
var engine = SkiaSharpFontEngine.CreateWithFontsAndSystemFonts(fallback, bold);
```

`CreateOnlyWithFonts` is the deterministic choice for servers: the same input produces the same
column widths on every machine, regardless of what is installed.

Embedding the fonts as assembly resources avoids shipping loose files:

```csharp
var assembly = typeof(Program).Assembly;

using var fallback = assembly.GetManifestResourceStream("MyApp.Fonts.Inter-Regular.ttf")!;
using var bold = assembly.GetManifestResourceStream("MyApp.Fonts.Inter-Bold.ttf")!;

var engine = SkiaSharpFontEngine.CreateOnlyWithFonts(fallback, bold);
```

The same API exists on `SixLaborsFontEngine` (2.x) and `DefaultFontEngine` (1.x):

```csharp
using XLibur.Fonts.SixLabors.V1;

var engine = DefaultFontEngine.CreateOnlyWithFonts(fallback, bold);
```

## Trimming and AOT

The zero-config path finds the default engine by reflection (`Assembly.Load`), which is
invisible to the trimmer. If you publish trimmed or AOT-compiled, register the engine
explicitly so the assembly is rooted:

```csharp
using XLibur.Fonts.SkiaSharp;

SkiaSharpFontBootstrap.Register();   // a direct reference the trimmer can see
```

## When fonts matter — and when they don't

Text measurement is only used for layout that depends on glyph size:

| Operation | Needs a font engine |
|---|---|
| `Columns().AdjustToContents()` | Yes |
| `Rows().AdjustToContents()` | Yes |
| Automatic row height for wrapped text | Yes |
| Reading and writing cell values | No |
| Formulas and evaluation | No |
| Styles, tables, pivot tables | No |
| Explicit `Column.Width` / `Row.Height` | No |

If your generated files always set explicit widths, the font engine never runs — but the
package must still be present, because a workbook cannot be constructed without one.

## Font properties on cells

Choosing a *typeface* for a cell is a styling concern, separate from the engine that measures
it — see [Styling](./styling.md#font):

```csharp
ws.Cell("A1").Style.Font.FontName = "Segoe UI";
ws.Cell("A1").Style.Font.FontSize = 14;
ws.Cell("A1").Style.Font.Bold = true;

// Non-Latin scripts
ws.Cell("A2").Style
    .Font.SetFontName("Arabic Typesetting")
    .Font.SetFontCharSet(XLFontCharSet.Arabic);

// Follow the workbook theme's heading/body fonts
ws.Cell("A3").Style.Font.FontScheme = XLFontScheme.Major;
```

:::note
Setting `FontName = "Inter"` writes that name into the file. Excel renders it if the machine
opening the file has Inter installed; XLibur measures it if the *generating* machine's font
engine can resolve it. These are two independent concerns.
:::

## Reference

The full design — interface separation, package structure, and why registration works the way
it does — is documented in
[docs/font-architecture.md](https://github.com/XLibur/XLibur/blob/main/docs/font-architecture.md)
in the repository.

## Where to next

- [Styling](./styling.md) — font properties on cells and ranges
- [Getting Started](./getting-started.md) — installation and package choice
