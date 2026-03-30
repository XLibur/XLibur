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

The core `XLibur` package uses SixLabors.Fonts 1.0.1 (via `XLibur.Fonts.SixLabors.V1`), while `XLibur.Fonts.SixLabors` uses version 2.1.3. Keeping the examples in a separate project avoids mixing both versions in the same dependency graph.
