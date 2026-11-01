using System.Reflection;
using XLibur.Excel;
using XLibur.Graphics;

namespace XLibur.Fonts.SkiaSharp;

/// <summary>
/// Bootstrap helper to register the SkiaSharp font engine with XLibur.
/// </summary>
/// <remarks>
/// <para>
/// SkiaSharp is the <b>default</b> font engine for XLibur. When the <c>XLibur.Fonts.SkiaSharp</c>
/// package (or the <c>XLibur.Bundle</c> meta-package) is referenced, XLibur core auto-registers this
/// engine on first workbook creation — you normally do not need to call <see cref="Register"/> at all.
/// </para>
/// <para>
/// Call <see cref="Register"/> explicitly only if you want to force registration at a specific point in
/// application startup, or override a different default that was registered earlier.
/// </para>
/// </remarks>
/// <example>
/// <code>
/// // Zero-config — nothing to call, SkiaSharp is discovered automatically:
/// using var wb = new XLWorkbook();
///
/// // Or register explicitly at application startup:
/// SkiaSharpFontBootstrap.Register();
/// </code>
/// </example>
public static class SkiaSharpFontBootstrap
{
    /// <summary>
    /// Family name of the embedded CarlitoBare fallback font, as reported by SkiaSharp.
    /// </summary>
    private const string EmbeddedFontName = "CarlitoBare";

    /// <summary>
    /// The system font tried before falling back to the embedded font. Matches the historical
    /// SixLabors default so behavior is consistent across engines on Windows.
    /// </summary>
    private const string DefaultFallbackFont = "Microsoft Sans Serif";

    /// <summary>
    /// Register the SkiaSharp engine as the default font engine.
    /// Sets <see cref="LoadOptions.DefaultFontEngine"/> so all workbooks created without an explicit
    /// font engine use the SkiaSharp implementation. Safe to call multiple times — subsequent calls
    /// (and a default already set by another engine) are preserved, i.e. this call is a no-op if a
    /// default is already registered.
    /// </summary>
    public static void Register()
    {
        LoadOptions.DefaultFontEngine ??= CreateDefault();
    }

    /// <summary>
    /// Create the default SkiaSharp font engine: resolves system fonts, uses
    /// <c>Microsoft Sans Serif</c> as the named fallback, and an embedded metric-only CarlitoBare
    /// (Calibri-compatible) font as the ultimate fallback so measurement never fails on machines
    /// without system fonts.
    /// </summary>
    /// <remarks>
    /// This method is the entry point XLibur core invokes reflectively to auto-register the default
    /// engine. Renaming it or changing its signature is a breaking change for auto-registration.
    /// </remarks>
    public static IXLFontEngine CreateDefault()
    {
        var assembly = typeof(SkiaSharpFontBootstrap).Assembly;
        string[] embeddedFontResourcePaths =
        [
            "XLibur.Fonts.SkiaSharp.Fonts.CarlitoBare-Regular.ttf",
            "XLibur.Fonts.SkiaSharp.Fonts.CarlitoBare-Bold.ttf",
            "XLibur.Fonts.SkiaSharp.Fonts.CarlitoBare-Italic.ttf",
            "XLibur.Fonts.SkiaSharp.Fonts.CarlitoBare-BoldItalic.ttf",
        ];

        return new SkiaSharpFontEngine(DefaultFallbackFont, EmbeddedFontName, embeddedFontResourcePaths, assembly);
    }
}
