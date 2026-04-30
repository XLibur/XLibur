# XLibur.Bundle

A convenience meta-package that installs [XLibur](https://www.nuget.org/packages/XLibur) together with its default font engine, [XLibur.Fonts.SixLabors.V1](https://www.nuget.org/packages/XLibur.Fonts.SixLabors.V1) (SixLabors.Fonts 1.x, Apache 2.0).

This package contains no code of its own — it exists purely to pull both dependencies in via a single `PackageReference`.

## Install

```
dotnet add package XLibur.Bundle
```

That's all. The default font engine registers itself automatically when the assembly loads.

## When to use this

Install `XLibur.Bundle` if you want the recommended, license-safe defaults and don't want to think about font engines.

## When to install the pieces separately

- You want a different font engine (e.g. `XLibur.Fonts.SixLabors` for SixLabors.Fonts 2.x), in which case install `XLibur` plus the engine of your choice.
- You're a library author and don't want to force a font-engine choice on downstream consumers — depend only on `XLibur`.

## Documentation

For full documentation, source, and contribution guidelines, visit the [GitHub repository](https://github.com/XLibur/XLibur).
