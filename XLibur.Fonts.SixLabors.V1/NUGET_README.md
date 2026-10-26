# XLibur.Fonts.SixLabors.V1

> **Superseded by [XLibur.Fonts.PinataLabs](https://www.nuget.org/packages/XLibur.Fonts.PinataLabs).**
> SixLabors.Fonts 1.x no longer receives updates. XLibur.Fonts.PinataLabs is built on PinataLabs.Typography,
> a maintained Apache 2.0 fork of SixLabors.Fonts 1.x, and gives the same measurements. This package will be
> removed in a future release.

A font engine for [XLibur](https://www.nuget.org/packages/XLibur), providing text measurement and font metrics using [SixLabors.Fonts](https://github.com/SixLabors/Fonts) 1.x (Apache 2.0 licensed).

## Migrating to XLibur.Fonts.PinataLabs

Replace the package reference, then rename the types:

| XLibur.Fonts.SixLabors.V1 | XLibur.Fonts.PinataLabs |
| --- | --- |
| `using XLibur.Fonts.SixLabors.V1;` | `using XLibur.Fonts.PinataLabs;` |
| `DefaultFontEngine` | `PinataLabsFontEngine` |
| `SixLaborsV1FontBootstrap.Register()` | `PinataLabsFontBootstrap.Register()` |

The constructor and the `Instance`, `CreateOnlyWithFonts` and `CreateWithFontsAndSystemFonts` members are unchanged. XLibur.Fonts.PinataLabs has no module initializer: call `PinataLabsFontBootstrap.Register()` explicitly if you relied on this package registering itself when its assembly loaded.

## Upgrading to SixLabors.Fonts 2.x

If you need SixLabors.Fonts 2.x features, install `XLibur.Fonts.SixLabors` and configure:

```csharp
var options = new LoadOptions
{
    FontEngine = new SixLaborsFontEngine("Microsoft Sans Serif")
};
using var wb = new XLWorkbook(options);
```
