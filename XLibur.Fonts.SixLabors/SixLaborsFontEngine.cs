using System;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Reflection;
using SixLabors.Fonts;
using SixLabors.Fonts.Unicode;
using XLibur.Excel;
using XLibur.Graphics;

namespace XLibur.Fonts.SixLabors;

/// <summary>
/// Implementation of <see cref="IXLFontEngine"/> that uses SixLabors.Fonts 2.x for font metrics and text measurement.
/// </summary>
public class SixLaborsFontEngine : IXLFontEngine
{
    private const float FontMetricSize = 16f;

    private readonly Lazy<IReadOnlyFontCollection> _fontCollection;
    private readonly string _fallbackFont;
    private readonly string? _embeddedFontName;

    private readonly ConcurrentDictionary<MetricId, Font> _fonts = new();
    private readonly Func<MetricId, Font> _loadFont;

    private readonly ConcurrentDictionary<MetricId, double> _maxDigitWidths = new();
    private readonly Func<MetricId, double> _calculateMaxDigitWidth;

    /// <summary>
    /// Initialize a new instance of the engine.
    /// </summary>
    /// <param name="fallbackFont">A name of a font that is used when a font in a workbook is not available.</param>
    public SixLaborsFontEngine(string fallbackFont)
    {
        if (string.IsNullOrWhiteSpace(fallbackFont))
            throw new ArgumentException("Fallback font name must not be null or whitespace.", nameof(fallbackFont));

        var fontCollection = new FontCollection();

        _fontCollection = new Lazy<IReadOnlyFontCollection>(fontCollection.AddSystemFonts);
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
    public SixLaborsFontEngine(string fallbackFont, string embeddedFontName, string[] embeddedFontResourcePaths, Assembly assembly)
    {
        if (string.IsNullOrWhiteSpace(fallbackFont))
            throw new ArgumentException("Fallback font name must not be null or whitespace.", nameof(fallbackFont));
        ArgumentNullException.ThrowIfNull(embeddedFontResourcePaths);
        ArgumentNullException.ThrowIfNull(assembly);

        var fontCollection = new FontCollection();
        foreach (var resourcePath in embeddedFontResourcePaths)
        {
            using var stream = assembly.GetManifestResourceStream(resourcePath)
                               ?? throw new ArgumentException($"Embedded resource '{resourcePath}' not found in assembly '{assembly.FullName}'.");
            fontCollection.Add(stream);
        }

        _fontCollection = new Lazy<IReadOnlyFontCollection>(fontCollection.AddSystemFonts);
        _fallbackFont = fallbackFont;
        _embeddedFontName = embeddedFontName;
        _loadFont = LoadFont;
        _calculateMaxDigitWidth = CalculateMaxDigitWidth;
    }

    private SixLaborsFontEngine(Stream fallbackFontStream, bool useSystemFonts, Stream[] fontStreams)
    {
        ArgumentNullException.ThrowIfNull(fallbackFontStream);
        ArgumentNullException.ThrowIfNull(fontStreams);

        var fontCollection = new FontCollection();
        var fallbackFamily = fontCollection.Add(fallbackFontStream);
        foreach (var fontStream in fontStreams)
            fontCollection.Add(fontStream);

        _fontCollection = useSystemFonts
            ? new Lazy<IReadOnlyFontCollection>(fontCollection.AddSystemFonts)
            : new Lazy<IReadOnlyFontCollection>(() => fontCollection);
        _fallbackFont = fallbackFamily.Name;
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
        return new SixLaborsFontEngine(fallbackFontStream, false, fontStreams);
    }

    /// <summary>
    /// Create a font engine that uses fallback font and additional fonts passed as streams.
    /// It also uses system fonts.
    /// </summary>
    /// <param name="fallbackFontStream">A stream that contains a fallback font.</param>
    /// <param name="fontStreams">Fonts that should be loaded to the engine.</param>
    public static IXLFontEngine CreateWithFontsAndSystemFonts(Stream fallbackFontStream, params Stream[] fontStreams)
    {
        return new SixLaborsFontEngine(fallbackFontStream, true, fontStreams);
    }

    public double GetDescent(IXLFontBase font, double dpiY)
    {
        var metrics = GetMetrics(font);
        return GetDescent(font, dpiY, metrics);
    }

    private static double GetDescent(IXLFontBase font, double dpiY, FontMetrics metrics)
    {
        return PointsToPixels(-metrics.VerticalMetrics.Descender * font.FontSize / metrics.UnitsPerEm, dpiY);
    }

    public double GetMaxDigitWidth(IXLFontBase font, double dpiX)
    {
        var metricId = new MetricId(font);
        var maxDigitWidth = _maxDigitWidths.GetOrAdd(metricId, _calculateMaxDigitWidth);
        return PointsToPixels(maxDigitWidth * font.FontSize, dpiX);
    }

    public double GetTextHeight(IXLFontBase font, double dpiY)
    {
        var metrics = GetMetrics(font);
        return PointsToPixels(
            (metrics.VerticalMetrics.Ascender - 2 * metrics.VerticalMetrics.Descender) * font.FontSize /
            metrics.UnitsPerEm, dpiY);
    }

    public double GetTextWidth(string text, IXLFontBase font, double dpiX)
    {
        var fontInstance = GetFont(font);
        var dimensionsPx = TextMeasurer.MeasureAdvance(text, new TextOptions(fontInstance)
        {
            Dpi = 72, // Normalize DPI, so 1px is 1pt
            KerningMode = KerningMode.None
        });
        return PointsToPixels(dimensionsPx.Width / FontMetricSize * font.FontSize, dpiX);
    }

    /// <inheritdoc />
    public GlyphBox GetGlyphBox(ReadOnlySpan<int> graphemeCluster, IXLFontBase font, Dpi dpi)
    {
        var metric = GetMetrics(font);
        var advanceFu = 0;
        foreach (var t in graphemeCluster)
        {
            var containsMetrics = metric.TryGetGlyphMetrics(
                new CodePoint(t),
                TextAttributes.None,
                TextDecorations.None,
                LayoutMode.HorizontalTopBottom,
                ColorFontSupport.None,
                out var glyphs);

            if (!containsMetrics)
                continue;

            advanceFu = glyphs!.Aggregate(advanceFu, (current, glyph) => current + glyph.AdvanceWidth);
        }

        var emInPx = font.FontSize / 72d * dpi.Y;
        var advancePx = PointsToPixels(advanceFu * font.FontSize / metric.UnitsPerEm, dpi.X);
        var descentPx = GetDescent(font, dpi.Y, metric);
        return new GlyphBox(
            (float)advancePx,
            (float)Math.Round(emInPx, MidpointRounding.AwayFromZero),
            (float)Math.Round(descentPx, MidpointRounding.AwayFromZero));
    }

    private FontMetrics GetMetrics(IXLFontBase fontBase)
    {
        var font = GetFont(fontBase);
        return font.FontMetrics;
    }

    private Font GetFont(IXLFontBase fontBase)
    {
        return GetFont(new MetricId(fontBase));
    }

    private Font GetFont(MetricId metricId)
    {
        return _fonts.GetOrAdd(metricId, _loadFont);
    }

    private Font LoadFont(MetricId metricId)
    {
        if (!_fontCollection.Value.TryGet(metricId.Name, out var fontFamily) &&
            !_fontCollection.Value.TryGet(_fallbackFont, out fontFamily))
        {
            if (_embeddedFontName is not null)
            {
                fontFamily = _fontCollection.Value.Get(_embeddedFontName);
            }
            else
            {
                throw new InvalidOperationException(
                    $"Font '{metricId.Name}' not found, and fallback font '{_fallbackFont}' is also not available. " +
                    "Consider providing a fallback font stream or using system fonts.");
            }
        }

        return fontFamily.CreateFont(FontMetricSize, metricId.Style);
    }

    private double CalculateMaxDigitWidth(MetricId metricId)
    {
        var font = GetFont(metricId);
        var metrics = font.FontMetrics;
        var maxWidth = 0;
        for (var c = '0'; c <= '9'; ++c)
        {
            var containsMetrics = metrics.TryGetGlyphMetrics(
                new CodePoint(c),
                TextAttributes.None,
                TextDecorations.None,
                LayoutMode.HorizontalTopBottom,
                ColorFontSupport.None,
                out var glyphMetrics);
            if (!containsMetrics)
                continue;

            var glyphAdvance = glyphMetrics!.Aggregate(0, (current, glyphMetric) => current + glyphMetric.AdvanceWidth);

            maxWidth = Math.Max(maxWidth, glyphAdvance);
        }

        return maxWidth / (double)metrics.UnitsPerEm;
    }

    private static double PointsToPixels(double points, double dpi) => points / 72d * dpi;

    private readonly struct MetricId : IEquatable<MetricId>
    {
        public MetricId(IXLFontBase fontBase)
        {
            Name = fontBase.FontName;
            Style = GetFontStyle(fontBase);
        }

        public string Name { get; }

        public FontStyle Style { get; }

        public bool Equals(MetricId other) => Name == other.Name && Style == other.Style;

        public override bool Equals(object? obj) => obj is MetricId other && Equals(other);

        public override int GetHashCode() => (Name.GetHashCode() * 397) ^ (int)Style;

        private static FontStyle GetFontStyle(IXLFontBase fontBase)
        {
            return fontBase switch
            {
                { Bold: true, Italic: true } => FontStyle.BoldItalic,
                { Bold: true } => FontStyle.Bold,
                { Italic: true } => FontStyle.Italic,
                _ => FontStyle.Regular
            };
        }
    }
}
