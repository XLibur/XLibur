# XLibur.Fonts.PinataLabs

An optional font engine for [XLibur](https://www.nuget.org/packages/XLibur) that uses [PinataLabs.Typography](https://github.com/PinataLabs/Typography) for text measurement and font metrics.

PinataLabs.Typography is a maintained fork of SixLabors.Fonts 1.x, under the Apache 2.0 license. This package replaces `XLibur.Fonts.SixLabors.V1`, which is no longer published: it has the same API and gives the same measurements, and its dependency still receives updates.

## Installation

```sh
dotnet add package XLibur.Fonts.PinataLabs
```

## Usage

XLibur's default font engine is SkiaSharp. This engine is opt-in, so register it at startup:

```csharp
using XLibur.Excel;
using XLibur.Fonts.PinataLabs;

// Make PinataLabs the default engine for every workbook:
PinataLabsFontBootstrap.Register();
using var wb = new XLWorkbook();

// Or use system fonts with a named fallback, for one workbook:
var options = new LoadOptions { FontEngine = new PinataLabsFontEngine("Arial") };
using var wb2 = new XLWorkbook(options);

// Or use stream-based fonts only (useful for Blazor, serverless, etc.):
using var fontStream = File.OpenRead("MyFont.ttf");
var engine = PinataLabsFontEngine.CreateOnlyWithFonts(fontStream);
using var wb3 = new XLWorkbook(new LoadOptions { FontEngine = engine });
```

The engine resolves system fonts and falls back to an embedded metric-only Calibri-compatible font (`CarlitoBare`), so text measurement works even where no system fonts are installed.

## Migrating from XLibur.Fonts.SixLabors.V1

Replace the package reference, then rename the types:

| XLibur.Fonts.SixLabors.V1 | XLibur.Fonts.PinataLabs |
| --- | --- |
| `using XLibur.Fonts.SixLabors.V1;` | `using XLibur.Fonts.PinataLabs;` |
| `DefaultFontEngine` | `PinataLabsFontEngine` |
| `SixLaborsV1FontBootstrap.Register()` | `PinataLabsFontBootstrap.Register()` |

The constructor and the `Instance`, `CreateOnlyWithFonts` and `CreateWithFontsAndSystemFonts` members are unchanged. Unlike V1, this package has no module initializer: call `PinataLabsFontBootstrap.Register()` explicitly if you relied on V1 registering itself when its assembly loaded.

## License

This package is licensed under MIT. PinataLabs.Typography is licensed under Apache 2.0.
