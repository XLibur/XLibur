# XLibur.Fonts.SixLabors.Examples

Example applications demonstrating how to use [SixLabors.Fonts 2.x](https://github.com/SixLabors/Fonts) as the font engine for XLibur.

## Running

```bash
dotnet run --project XLibur.Fonts.SixLabors.Examples
```

Output files are written to the `Created/` directory under the build output.

## Examples

### UsingSixLaborsFontsV2

Shows how to configure `SixLaborsFontEngine` as the font engine for a workbook:

- Create a `SixLaborsFontEngine` with a system font fallback
- Pass it via `LoadOptions.FontEngine` per-workbook
- Set `LoadOptions.DefaultFontEngine` as a global default
- Use `AdjustToContents()` which relies on the font engine for text measurement

## Why a separate project?

`XLibur.Fonts.SixLabors` uses SixLabors.Fonts 2.1.3, which is under the Six Labors Split License. Keeping the examples in a separate project keeps that dependency out of the main examples.
