# XLibur.Fonts.SixLabors.V1

The default font engine for [XLibur](https://www.nuget.org/packages/XLibur), providing text measurement and font metrics using [SixLabors.Fonts](https://github.com/SixLabors/Fonts) 1.x (Apache 2.0 licensed).

This package is automatically included as a dependency of XLibur. You do not need to install it separately unless you want to reference `DefaultFontEngine` directly.

## Upgrading to SixLabors.Fonts 2.x

If you need SixLabors.Fonts 2.x features, install `XLibur.Fonts.SixLabors` and configure:

```csharp
var options = new LoadOptions
{
    FontEngine = new SixLaborsFontEngine("Microsoft Sans Serif")
};
using var wb = new XLWorkbook(options);
```
