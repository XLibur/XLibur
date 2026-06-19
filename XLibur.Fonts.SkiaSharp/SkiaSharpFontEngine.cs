using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using SkiaSharp;
using XLibur.Excel;
using XLibur.Graphics;

namespace XLibur.Fonts.SkiaSharp;

/// <summary>
/// Implementation of <see cref="IXLFontEngine"/> that uses SkiaSharp for font metrics and text measurement.
/// SkiaSharp is MIT-licensed, providing a permissive alternative to the SixLabors.Fonts engines.
/// </summary>
public class SkiaSharpFontEngine : IXLFontEngine
{
    private const float FontMetricSize = 16f;

    private readonly IReadOnlyDictionary<string, SKTypeface> _streamFonts;
    private readonly bool _useSystemFonts;
    private readonly string _fallbackFont;
    private readonly string? _embeddedFontName;

    private readonly ConcurrentDictionary<MetricId, FontEntry> _fonts = new();
    private readonly Func<MetricId, FontEntry> _loadFont;

    private readonly ConcurrentDictionary<MetricId, double> _maxDigitWidths = new();
    private readonly Func<MetricId, double> _calculateMaxDigitWidth;

    /// <summary>
    /// Initialize a new instance of the engine that resolves fonts from the operating system's installed fonts.
    /// </summary>
    /// <param name="fallbackFont">A name of a font that is used when a font in a workbook is not available.</param>
    public SkiaSharpFontEngine(string fallbackFont)
    {
        if (string.IsNullOrWhiteSpace(fallbackFont))
            throw new ArgumentException("Fallback font name must not be null or whitespace.", nameof(fallbackFont));

        _streamFonts = new Dictionary<string, SKTypeface>(StringComparer.OrdinalIgnoreCase);
        _useSystemFonts = true;
        _fallbackFont = fallbackFont;
        _embeddedFontName = null;
        _loadFont = LoadFont;
        _calculateMaxDigitWidth = CalculateMaxDigitWidth;
    }

    /// <summary>
    /// Initialize a new instance of the engine with embedded fallback fonts from a specified assembly.
    /// </summary>
    /// <param name="fallbackFont">A name of a font that is used when a font in a workbook is not available.</param>
    /// <param name="embeddedFontName">The name of the embedded font family used as ultimate fallback.</param>
    /// <param name="embeddedFontResourcePaths">Resource paths for the embedded font files to load from the calling assembly.</param>
    /// <param name="assembly">The assembly containing the embedded font resources.</param>
    public SkiaSharpFontEngine(string fallbackFont, string embeddedFontName, string[] embeddedFontResourcePaths, Assembly assembly)
    {
        if (string.IsNullOrWhiteSpace(fallbackFont))
            throw new ArgumentException("Fallback font name must not be null or whitespace.", nameof(fallbackFont));
        if (string.IsNullOrWhiteSpace(embeddedFontName))
            throw new ArgumentException("Embedded font name must not be null or whitespace.", nameof(embeddedFontName));
        ArgumentNullException.ThrowIfNull(embeddedFontResourcePaths);
        ArgumentNullException.ThrowIfNull(assembly);

        var streamFonts = new Dictionary<string, SKTypeface>(StringComparer.OrdinalIgnoreCase);
        foreach (var resourcePath in embeddedFontResourcePaths)
        {
            using var stream = assembly.GetManifestResourceStream(resourcePath)
                               ?? throw new ArgumentException($"Embedded resource '{resourcePath}' not found in assembly '{assembly.FullName}'.");
            AddTypeface(streamFonts, stream);
        }

        _streamFonts = streamFonts;
        _useSystemFonts = true;
        _fallbackFont = fallbackFont;
        _embeddedFontName = embeddedFontName;
        _loadFont = LoadFont;
        _calculateMaxDigitWidth = CalculateMaxDigitWidth;
    }

    private SkiaSharpFontEngine(Stream fallbackFontStream, bool useSystemFonts, Stream[] fontStreams)
    {
        ArgumentNullException.ThrowIfNull(fallbackFontStream);
        ArgumentNullException.ThrowIfNull(fontStreams);

        var streamFonts = new Dictionary<string, SKTypeface>(StringComparer.OrdinalIgnoreCase);
        var fallbackName = AddTypeface(streamFonts, fallbackFontStream);
        foreach (var fontStream in fontStreams)
            AddTypeface(streamFonts, fontStream);

        _streamFonts = streamFonts;
        _useSystemFonts = useSystemFonts;
        _fallbackFont = fallbackName;
        _embeddedFontName = null;
        _loadFont = LoadFont;
        _calculateMaxDigitWidth = CalculateMaxDigitWidth;
    }

    /// <summary>
    /// Create a font engine that uses only fallback font and additional fonts passed as streams.
    /// It ignores all system fonts, and that can lead to a decrease of initialization time.
    /// </summary>
    /// <param name="fallbackFontStream">A stream that contains a fallback font.</param>
    /// <param name="fontStreams">Fonts that should be loaded to the engine.</param>
    public static IXLFontEngine CreateOnlyWithFonts(Stream fallbackFontStream, params Stream[] fontStreams)
    {
        return new SkiaSharpFontEngine(fallbackFontStream, false, fontStreams);
    }

    /// <summary>
    /// Create a font engine that uses fallback font and additional fonts passed as streams.
    /// It also uses system fonts.
    /// </summary>
    /// <param name="fallbackFontStream">A stream that contains a fallback font.</param>
    /// <param name="fontStreams">Fonts that should be loaded to the engine.</param>
    public static IXLFontEngine CreateWithFontsAndSystemFonts(Stream fallbackFontStream, params Stream[] fontStreams)
    {
        return new SkiaSharpFontEngine(fallbackFontStream, true, fontStreams);
    }

    public double GetDescent(IXLFontBase font, double dpiY)
    {
        var entry = GetFont(font);
        return GetDescent(font, dpiY, entry);
    }

    private static double GetDescent(IXLFontBase font, double dpiY, FontEntry entry)
    {
        return PointsToPixels(entry.DescentFu * font.FontSize / entry.UnitsPerEm, dpiY);
    }

    public double GetMaxDigitWidth(IXLFontBase fontBase, double dpiX)
    {
        var metricId = new MetricId(fontBase);
        var maxDigitWidth = _maxDigitWidths.GetOrAdd(metricId, _calculateMaxDigitWidth);
        return PointsToPixels(maxDigitWidth * fontBase.FontSize, dpiX);
    }

    public double GetTextHeight(IXLFontBase font, double dpiY)
    {
        var entry = GetFont(font);
        return PointsToPixels(
            (entry.AscentFu + 2 * entry.DescentFu) * font.FontSize / entry.UnitsPerEm, dpiY);
    }

    public double GetTextWidth(string text, IXLFontBase fontBase, double dpiX)
    {
        var entry = GetFont(fontBase);
        using var font = new SKFont(entry.Typeface, FontMetricSize);
        var advance = font.MeasureText(text);
        return PointsToPixels(advance / FontMetricSize * fontBase.FontSize, dpiX);
    }

    /// <inheritdoc />
    public GlyphBox GetGlyphBox(ReadOnlySpan<int> graphemeCluster, IXLFontBase font, Dpi dpi)
    {
        var entry = GetFont(font);
        using var measureFont = new SKFont(entry.Typeface, entry.UnitsPerEm);
        double advanceFu = 0;
        foreach (var codePoint in graphemeCluster)
            advanceFu += measureFont.MeasureText(char.ConvertFromUtf32(codePoint));

        var emInPx = font.FontSize / 72d * dpi.Y;
        var advancePx = PointsToPixels(advanceFu * font.FontSize / entry.UnitsPerEm, dpi.X);
        var descentPx = GetDescent(font, dpi.Y, entry);
        return new GlyphBox(
            (float)advancePx,
            (float)Math.Round(emInPx, MidpointRounding.AwayFromZero),
            (float)Math.Round(descentPx, MidpointRounding.AwayFromZero));
    }

    private FontEntry GetFont(IXLFontBase fontBase)
    {
        return _fonts.GetOrAdd(new MetricId(fontBase), _loadFont);
    }

    private FontEntry LoadFont(MetricId metricId)
    {
        var typeface = ResolveTypeface(metricId);

        using var font = new SKFont(typeface, typeface.UnitsPerEm);
        var metrics = font.Metrics;
        return new FontEntry(typeface, typeface.UnitsPerEm, -metrics.Ascent, metrics.Descent);
    }

    private SKTypeface ResolveTypeface(MetricId metricId)
    {
        // 1. Stream-loaded fonts by family name.
        if (_streamFonts.TryGetValue(metricId.Name, out var streamTypeface))
            return streamTypeface;

        // 2. System fonts (if enabled), only when the family is actually present.
        if (_useSystemFonts && TryMatchSystemFont(metricId.Name, metricId.SkStyle, out var systemTypeface))
            return systemTypeface;

        // 3. Fallback font (stream first, then system).
        if (_streamFonts.TryGetValue(_fallbackFont, out var fallbackTypeface))
            return fallbackTypeface;
        if (_useSystemFonts && TryMatchSystemFont(_fallbackFont, metricId.SkStyle, out var systemFallback))
            return systemFallback;

        // 4. Embedded ultimate fallback.
        if (_embeddedFontName is not null && _streamFonts.TryGetValue(_embeddedFontName, out var embeddedTypeface))
            return embeddedTypeface;

        throw new InvalidOperationException(
            $"Font '{metricId.Name}' not found, and fallback font '{_fallbackFont}' is also not available. " +
            "Consider providing a fallback font stream or using system fonts.");
    }

    private static bool TryMatchSystemFont(string familyName, SKFontStyle style, out SKTypeface typeface)
    {
        var match = SKFontManager.Default.MatchFamily(familyName, style);
        if (match is not null && string.Equals(match.FamilyName, familyName, StringComparison.OrdinalIgnoreCase))
        {
            typeface = match;
            return true;
        }

        match?.Dispose();
        typeface = null!;
        return false;
    }

    private double CalculateMaxDigitWidth(MetricId metricId)
    {
        var entry = _fonts.GetOrAdd(metricId, _loadFont);
        using var font = new SKFont(entry.Typeface, entry.UnitsPerEm);
        float maxWidth = 0;
        for (var c = '0'; c <= '9'; ++c)
            maxWidth = Math.Max(maxWidth, font.MeasureText(c.ToString()));

        return maxWidth / entry.UnitsPerEm;
    }

    private static string AddTypeface(IDictionary<string, SKTypeface> fonts, Stream stream)
    {
        using var data = SKData.Create(stream);
        var typeface = SKTypeface.FromData(data)
                       ?? throw new ArgumentException("The provided stream does not contain a valid font.");
        fonts[typeface.FamilyName] = typeface;
        return typeface.FamilyName;
    }

    private static double PointsToPixels(double points, double dpi) => points / 72d * dpi;

    private sealed class FontEntry
    {
        public FontEntry(SKTypeface typeface, int unitsPerEm, double ascentFu, double descentFu)
        {
            Typeface = typeface;
            UnitsPerEm = unitsPerEm;
            AscentFu = ascentFu;
            DescentFu = descentFu;
        }

        /// <summary>The resolved typeface.</summary>
        public SKTypeface Typeface { get; }

        /// <summary>Font design units per em.</summary>
        public int UnitsPerEm { get; }

        /// <summary>Ascent in font units (positive value).</summary>
        public double AscentFu { get; }

        /// <summary>Descent in font units (positive value).</summary>
        public double DescentFu { get; }
    }

    private readonly struct MetricId : IEquatable<MetricId>
    {
        public MetricId(IXLFontBase fontBase)
        {
            Name = fontBase.FontName;
            Style = GetFontStyle(fontBase);
        }

        public string Name { get; }

        public FontStyleKind Style { get; }

        public SKFontStyle SkStyle => Style switch
        {
            FontStyleKind.BoldItalic => SKFontStyle.BoldItalic,
            FontStyleKind.Bold => SKFontStyle.Bold,
            FontStyleKind.Italic => SKFontStyle.Italic,
            _ => SKFontStyle.Normal
        };

        public bool Equals(MetricId other) => Name == other.Name && Style == other.Style;

        public override bool Equals(object? obj) => obj is MetricId other && Equals(other);

        public override int GetHashCode() => (Name.GetHashCode() * 397) ^ (int)Style;

        private static FontStyleKind GetFontStyle(IXLFontBase fontBase)
        {
            return fontBase switch
            {
                { Bold: true, Italic: true } => FontStyleKind.BoldItalic,
                { Bold: true } => FontStyleKind.Bold,
                { Italic: true } => FontStyleKind.Italic,
                _ => FontStyleKind.Regular
            };
        }
    }

    private enum FontStyleKind
    {
        Regular,
        Bold,
        Italic,
        BoldItalic
    }
}
