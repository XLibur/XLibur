using XLibur.Excel;
using XLibur.Graphics;

namespace XLibur.Fonts.PinataLabs;

/// <summary>
/// Bootstrap helper to register the PinataLabs.Typography font engine with XLibur.
/// </summary>
/// <remarks>
/// <para>
/// This engine is opt-in: XLibur's default engine is SkiaSharp, and referencing this package changes
/// nothing until you register it. Call <see cref="Register"/> once at application startup, before
/// creating any workbooks, or pass an engine to <see cref="LoadOptions.FontEngine"/> per workbook.
/// </para>
/// </remarks>
/// <example>
/// <code>
/// // In Program.cs or application startup:
/// PinataLabsFontBootstrap.Register();
///
/// // Then use XLibur normally:
/// using var wb = new XLWorkbook();
/// </code>
/// </example>
public static class PinataLabsFontBootstrap
{
    /// <summary>
    /// The system font tried before falling back to the embedded font. Matches the historical
    /// SixLabors default so behavior is consistent across engines on Windows.
    /// </summary>
    internal const string DefaultFallbackFont = "Microsoft Sans Serif";

    /// <summary>
    /// The embedded metric-only CarlitoBare (Calibri-compatible) fonts, the ultimate fallback.
    /// </summary>
    internal static readonly string[] EmbeddedFontResourcePaths =
    [
        "XLibur.Fonts.PinataLabs.Fonts.CarlitoBare-Regular.ttf",
        "XLibur.Fonts.PinataLabs.Fonts.CarlitoBare-Bold.ttf",
        "XLibur.Fonts.PinataLabs.Fonts.CarlitoBare-Italic.ttf",
        "XLibur.Fonts.PinataLabs.Fonts.CarlitoBare-BoldItalic.ttf",
    ];

    /// <summary>
    /// Register the PinataLabs.Typography engine as the default font engine.
    /// Sets <see cref="LoadOptions.DefaultFontEngine"/> so all workbooks created without an explicit
    /// font engine use the PinataLabs implementation. Safe to call multiple times — this call is a
    /// no-op if a default is already registered.
    /// </summary>
    public static void Register()
    {
        LoadOptions.DefaultFontEngine ??= CreateDefault();
    }

    /// <summary>
    /// Create the default PinataLabs.Typography font engine: resolves system fonts, uses
    /// <c>Microsoft Sans Serif</c> as the named fallback, and an embedded metric-only CarlitoBare
    /// (Calibri-compatible) font as the ultimate fallback so measurement never fails on machines
    /// without system fonts.
    /// </summary>
    /// <remarks>Returns the shared <see cref="PinataLabsFontEngine.Instance"/>.</remarks>
    public static IXLFontEngine CreateDefault() => PinataLabsFontEngine.Instance.Value;
}
