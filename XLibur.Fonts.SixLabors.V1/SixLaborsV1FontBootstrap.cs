using XLibur.Graphics;

namespace XLibur.Fonts.SixLabors.V1;

/// <summary>
/// Bootstrap helper to register the SixLabors.Fonts v1 font engine with XLibur.
/// Call <see cref="Register"/> once at application startup before creating any workbooks.
/// </summary>
/// <example>
/// <code>
/// // In Program.cs or application startup:
/// SixLaborsV1FontBootstrap.Register();
///
/// // Then use XLibur normally:
/// using var wb = new XLWorkbook();
/// </code>
/// </example>
public static class SixLaborsV1FontBootstrap
{
    /// <summary>
    /// Register the SixLabors.Fonts v1 engine factories with <see cref="DefaultFontEngineFactory"/>.
    /// This enables <see cref="DefaultGraphicEngine"/> and <see cref="DefaultFontEngine"/> to be
    /// created without explicit font engine configuration.
    /// </summary>
    /// <remarks>
    /// Safe to call multiple times — subsequent calls are no-ops.
    /// Also called automatically by the module initializer when this assembly is loaded,
    /// but explicit registration is preferred for clarity.
    /// </remarks>
    public static void Register()
    {
        DefaultFontEngineFactory.FromFallbackFont ??= fallbackFont => new DefaultFontEngine(fallbackFont);
        DefaultFontEngineFactory.FromStreams ??= (stream, useSystemFonts, fontStreams)
            => new DefaultFontEngine(stream, useSystemFonts, fontStreams);
    }
}
