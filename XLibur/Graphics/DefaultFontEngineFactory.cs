using System;
using System.IO;

namespace XLibur.Graphics;

/// <summary>
/// Internal factory delegates used by <see cref="DefaultGraphicEngine"/> to create font engines
/// without a direct assembly reference to the font engine implementation.
/// </summary>
/// <remarks>
/// This is an internal implementation detail. Users should set <see cref="Excel.LoadOptions.DefaultFontEngine"/>
/// or call the font engine package's bootstrap method (e.g., <c>SixLaborsV1FontBootstrap.Register()</c>).
/// </remarks>
internal static class DefaultFontEngineFactory
{
    /// <summary>
    /// Factory that creates an <see cref="IXLFontEngine"/> from a fallback font name.
    /// </summary>
    internal static Func<string, IXLFontEngine>? FromFallbackFont { get; set; }

    /// <summary>
    /// Factory that creates an <see cref="IXLFontEngine"/> from font streams.
    /// </summary>
    /// <remarks>Parameters: fallbackFontStream, useSystemFonts, additional fontStreams.</remarks>
    internal static Func<Stream, bool, Stream[], IXLFontEngine>? FromStreams { get; set; }

    internal static IXLFontEngine CreateFromFallbackFont(string fallbackFont)
    {
        return FromFallbackFont?.Invoke(fallbackFont)
               ?? throw new InvalidOperationException(
                   "No font engine is registered. " +
                   "Call SixLaborsV1FontBootstrap.Register() at application startup, " +
                   "or set LoadOptions.DefaultFontEngine manually. " +
                   "See https://github.com/XLibur/XLibur for setup instructions.");
    }

    internal static IXLFontEngine CreateFromStreams(Stream fallbackFontStream, bool useSystemFonts, Stream[] fontStreams)
    {
        return FromStreams?.Invoke(fallbackFontStream, useSystemFonts, fontStreams)
               ?? throw new InvalidOperationException(
                   "No font engine is registered. " +
                   "Call SixLaborsV1FontBootstrap.Register() at application startup, " +
                   "or set LoadOptions.DefaultFontEngine manually. " +
                   "See https://github.com/XLibur/XLibur for setup instructions.");
    }
}
