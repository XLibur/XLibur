# XLibur.Fonts.SixLabors

An optional font engine for [XLibur](https://www.nuget.org/packages/XLibur) that uses [SixLabors.Fonts](https://github.com/SixLabors/Fonts) 2.x for text measurement and font metrics.

## Installation

```
dotnet add package XLibur.Fonts.SixLabors
```

## Usage

```csharp
using XLibur.Excel;
using XLibur.Fonts.SixLabors;

// Use system fonts with a named fallback
var fontEngine = new SixLaborsFontEngine("Microsoft Sans Serif");
var options = new LoadOptions { FontEngine = fontEngine };
using var wb = new XLWorkbook(options);

// Or use stream-based fonts (useful for Blazor, serverless, etc.)
using var fontStream = File.OpenRead("MyFont.ttf");
var engine = SixLaborsFontEngine.CreateOnlyWithFonts(fontStream);
var options2 = new LoadOptions { FontEngine = engine };
using var wb2 = new XLWorkbook(options2);
```

## License

This package is licensed under MIT. Note that SixLabors.Fonts 2.x uses the [Six Labors Split License](https://github.com/SixLabors/Fonts/blob/main/LICENSE), which has different terms for commercial use.
