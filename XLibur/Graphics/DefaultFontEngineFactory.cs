using System;
using System.IO;

namespace XLibur.Graphics;

/// <summary>
/// Registration point for font engine factories. Font engine packages register their factories here
/// so that <see cref="DefaultGraphicEngine"/> can create font engines without a direct assembly reference.
/// </summary>
/// <remarks>
/// <para>
/// To register the default SixLabors.Fonts v1 engine, call early in your application startup:
/// <code>
/// XLibur.Fonts.SixLabors.V1.SixLaborsV1FontBootstrap.Register();
/// </code>
/// Or set a font engine directly:
/// <code>
/// LoadOptions.DefaultFontEngine = new DefaultFontEngine("Microsoft Sans Serif");
/// </code>
/// </para>
/// </remarks>
public static class DefaultFontEngineFactory
{
    /// <summary>
    /// Factory that creates an <see cref="IXLFontEngine"/> from a fallback font name.
    /// </summary>
    public static Func<string, IXLFontEngine>? FromFallbackFont { get; set; }

    /// <summary>
    /// Factory that creates an <see cref="IXLFontEngine"/> from font streams.
    /// </summary>
    /// <remarks>Parameters: fallbackFontStream, useSystemFonts, additional fontStreams.</remarks>
    public static Func<Stream, bool, Stream[], IXLFontEngine>? FromStreams { get; set; }

    internal static IXLFontEngine CreateFromFallbackFont(string fallbackFont)
    {
        return FromFallbackFont?.Invoke(fallbackFont)
               ?? throw new InvalidOperationException(
                   "No font engine provider is registered. " +
                   "Call SixLaborsV1FontBootstrap.Register() at application startup, " +
                   "or set DefaultFontEngineFactory.FromFallbackFont / LoadOptions.FontEngine manually. " +
                   "See https://github.com/XLibur/XLibur for setup instructions.");
    }

    internal static IXLFontEngine CreateFromStreams(Stream fallbackFontStream, bool useSystemFonts, Stream[] fontStreams)
    {
        return FromStreams?.Invoke(fallbackFontStream, useSystemFonts, fontStreams)
               ?? throw new InvalidOperationException(
                   "No font engine provider is registered. " +
                   "Call SixLaborsV1FontBootstrap.Register() at application startup, " +
                   "or set DefaultFontEngineFactory.FromStreams / LoadOptions.FontEngine manually. " +
                   "See https://github.com/XLibur/XLibur for setup instructions.");
    }
}
