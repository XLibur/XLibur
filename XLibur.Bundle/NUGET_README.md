# XLibur.Bundle

A convenience meta-package that installs [XLibur](https://www.nuget.org/packages/XLibur) together with its default font engine, [XLibur.Fonts.SkiaSharp](https://www.nuget.org/packages/XLibur.Fonts.SkiaSharp) (SkiaSharp, MIT).

This package contains no code of its own — it exists purely to pull both dependencies in via a single `PackageReference`.

## Install

```bash
dotnet add package XLibur.Bundle
```

That's all. The default font engine is auto-registered by XLibur core the first time you create a workbook — you do **not** need to call any startup method:

```csharp
using var wb = new XLWorkbook();
```

The default engine uses system fonts where available and falls back to an embedded metric-only Calibri-compatible font, so text measurement works even in headless and serverless environments with no system fonts installed.

## When to use this

Install `XLibur.Bundle` if you want the recommended, license-safe defaults and don't want to think about font engines.

## Overriding the default

To use a different engine, register it at startup (it takes precedence over the auto-registered default) or pass it per-workbook:

```csharp
// e.g. SixLabors.Fonts 2.x, per workbook
var options = new LoadOptions { FontEngine = new SixLaborsFontEngine("Microsoft Sans Serif") };
using var wb = new XLWorkbook(options);
```

## When to install the pieces separately

- You want a different font engine (e.g. `XLibur.Fonts.SixLabors.V1` for SixLabors.Fonts 1.x, or `XLibur.Fonts.SixLabors` for 2.x), in which case install `XLibur` plus the engine of your choice.
- You're a library author and don't want to force a font-engine choice on downstream consumers — depend only on `XLibur`.

## Documentation

For full documentation, source, and contribution guidelines, visit the [GitHub repository](https://github.com/XLibur/XLibur).
