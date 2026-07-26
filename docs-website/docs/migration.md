---
id: migration
title: Migration from ClosedXML
sidebar_label: Migration from ClosedXML
sidebar_position: 3
description: Move an existing ClosedXML project to XLibur — namespace changes and the one packaging difference, font engine configuration.
---

# Migration from ClosedXML

XLibur was forked from [ClosedXML v0.105.0](https://github.com/ClosedXML/ClosedXML/), and the
public API surface is largely unchanged. To migrate:

1. Install the NuGet package (see [Getting Started](./getting-started.md#installation)).
2. Replace `using ClosedXML` namespace references with `using XLibur`.

## Font engine configuration (different from ClosedXML)

This is the one area where XLibur's packaging differs from the base ClosedXML package. ClosedXML
bundles [SixLabors.Fonts](https://github.com/SixLabors/Fonts) directly into its core assembly for text
measurement (column auto-fit, row heights, glyph metrics). XLibur instead keeps the **core assembly
free of any font library** and ships the font engine as a **separate, swappable package**. This lets
you pick a font library with a license that suits you and avoids forcing a font dependency on library
authors who don't need one.

What this means when migrating:

- **Install `XLibur.Bundle` (or `XLibur` + `XLibur.Fonts.SkiaSharp`) and no code changes are needed.**
  The default [SkiaSharp](https://github.com/mono/SkiaSharp) engine (MIT-licensed) is **auto-registered
  by XLibur core the first time you create a workbook** — there is no startup call to add:

  ```csharp
  using var wb = new XLWorkbook(); // font engine resolved automatically
  ```

  The default resolves system fonts and falls back to an embedded, metric-only Calibri-compatible font,
  so text measurement works even in headless/serverless environments with no system fonts installed.

- **If you install the bare `XLibur` package with no font engine**, creating a workbook throws an
  `InvalidOperationException` telling you to add a font engine package. This is intentional — it's how
  the core stays font-library-agnostic.

- **To choose a different engine**, install its package and either register it at startup (it takes
  precedence over the auto-registered default) or pass it per workbook:

  | Package | Font library | License | Notes |
  |---|---|---|---|
  | `XLibur.Fonts.SkiaSharp` | SkiaSharp | MIT | **Default.** Auto-registers; ships native binaries. |
  | `XLibur.Fonts.SixLabors.V1` | SixLabors.Fonts 1.x | Apache 2.0 | Pure-managed; matches ClosedXML 0.105's engine. |
  | `XLibur.Fonts.SixLabors` | SixLabors.Fonts 2.x | Six Labors Split License | Commercial restrictions over $1M revenue. |

  ```csharp
  // Override globally at startup (e.g. keep ClosedXML's SixLabors 1.x behavior):
  SixLaborsV1FontBootstrap.Register();

  // Or override per workbook:
  var options = new LoadOptions { FontEngine = new SkiaSharpFontEngine("Arial") };
  using var wb = new XLWorkbook(options);
  ```

Resolution order for the font engine is: `LoadOptions.FontEngine` (per workbook) →
`LoadOptions.DefaultFontEngine` (explicitly registered global) → the auto-registered default engine.
See [Fonts and Font Engines](./fonts.md) for the full picture, including headless environments and
loading fonts from streams, or
[docs/font-architecture.md](https://github.com/XLibur/XLibur/blob/main/docs/font-architecture.md)
for the design rationale.
