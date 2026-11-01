# XLibur.Fonts.SkiaSharp

An optional font engine for [XLibur](https://www.nuget.org/packages/XLibur) that uses [SkiaSharp](https://github.com/mono/SkiaSharp) for text measurement and font metrics. SkiaSharp is MIT-licensed, making it a permissive alternative to the SixLabors.Fonts engines.

## Installation

```sh
dotnet add package XLibur.Fonts.SkiaSharp
```

## Usage

This is the **default** XLibur font engine. Once this package is referenced, XLibur core
auto-registers it the first time you create a workbook — no startup call is required:

```csharp
using XLibur.Excel;

// Zero-config: SkiaSharp is discovered and registered automatically.
using var wb = new XLWorkbook();
```

The auto-registered default resolves system fonts and falls back to an embedded metric-only
Calibri-compatible font (`CarlitoBare`), so text measurement works even in headless and serverless
environments that have no system fonts installed.

### Explicit / custom configuration

```csharp
using XLibur.Excel;
using XLibur.Fonts.SkiaSharp;

// Register explicitly at startup (optional — overrides an earlier default):
SkiaSharpFontBootstrap.Register();

// Use system fonts with a named fallback
var fontEngine = new SkiaSharpFontEngine("Arial");
var options = new LoadOptions { FontEngine = fontEngine };
using var wb = new XLWorkbook(options);

// Or use stream-based fonts (useful for Blazor, serverless, etc.)
using var fontStream = File.OpenRead("MyFont.ttf");
var engine = SkiaSharpFontEngine.CreateOnlyWithFonts(fontStream);
var options2 = new LoadOptions { FontEngine = engine };
using var wb2 = new XLWorkbook(options2);
```

## Native dependency

SkiaSharp wraps the native Skia graphics library, so this package brings native binaries per platform (unlike the pure-managed SixLabors.Fonts engines). On Linux it includes `SkiaSharp.NativeAssets.Linux.NoDependencies`, which needs no system `fontconfig`/`freetype` — stream-loaded fonts work in headless and serverless environments.

## License

This package is licensed under MIT. SkiaSharp is also MIT-licensed (wrapping the BSD-licensed Skia engine).
