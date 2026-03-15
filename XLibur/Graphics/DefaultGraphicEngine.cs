using System;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Reflection;
using XLibur.Excel;
using XLibur.Excel.Drawings;
using SixLabors.Fonts;
using SixLabors.Fonts.Unicode;

namespace XLibur.Graphics;

public class DefaultGraphicEngine : IXLGraphicEngine
{
    /// <summary>
    /// Carlito is a Calibri metric compatible font. This is a version stripped of everything but metric information
    /// to keep the embedded file small. It is reasonably accurate for many alphabets (contains 2531 glyphs). It has
    /// no glyph outlines, no TTF instructions, no substitutions, glyph positioning ect. It is created from Carlito
    /// font through strip-fonts.sh scripts.
    /// </summary>
    private const string EmbeddedFontName = "CarlitoBare";

    private const float FontMetricSize = 16f;

    private readonly ImageInfoReader[] _imageReaders =
    [
        new PngInfoReader(),
        new JpegInfoReader(),
        new GifInfoReader(),
        new TiffInfoReader(),
        new BmpInfoReader(),
        new EmfInfoReader(),
        new WmfInfoReader(),
        new WebpInfoReader(),
        new SvgInfoReader(),
        new PcxInfoReader() // Due to poor magic detection, keep last
    ];

    private readonly Lazy<IReadOnlyFontCollection> _fontCollection;
    private readonly string _fallbackFont;

    /// <summary>
    /// A font loaded font in the size <see cref="FontMetricSize"/>. There is no benefit in having multiple allocated instances, everything is just scaled at the moment.
    /// </summary>
    private readonly ConcurrentDictionary<MetricId, Font> _fonts = new();

    private readonly Func<MetricId, Font> _loadFont;

    /// <summary>
    /// Max digit width as a fraction of Em square. Multiply by font size to get pt size.
    /// </summary>
    private readonly ConcurrentDictionary<MetricId, double> _maxDigitWidths = new();

    private readonly Func<MetricId, double> _calculateMaxDigitWidth;

    /// <summary>
    /// Get a singleton instance of the engine that uses <c>Microsoft Sans Serif</c> as a fallback font.
    /// </summary>
    public static Lazy<DefaultGraphicEngine> Instance { get; } = new(() => new("Microsoft Sans Serif"));

    /// <summary>
    /// Initialize a new instance of the engine.
    /// </summary>
    /// <param name="fallbackFont">A name of a font that is used when a font in a workbook is not available.</param>
    public DefaultGraphicEngine(string fallbackFont)
    {
        if (string.IsNullOrWhiteSpace(fallbackFont))
            throw new ArgumentException(nameof(fallbackFont));

        var fontCollection = new FontCollection();
        AddEmbeddedFont(fontCollection);

        _fontCollection = new Lazy<IReadOnlyFontCollection>(fontCollection.AddSystemFonts);
        _fallbackFont = fallbackFont;
        _loadFont = LoadFont;
        _calculateMaxDigitWidth = CalculateMaxDigitWidth;
    }

    /// <summary>
    /// Initialize a new instance of the engine. The engine will be able to use system fonts and fonts loaded from external sources.
    /// </summary>
    /// <remarks>Useful/necessary for environments without access to the filesystem.</remarks>
    /// <param name="fallbackFontStream">A stream that contains a fallback font.</param>
    /// <param name="useSystemFonts">Should the engine try to use system fonts? If false, system fonts won't be loaded, which can significantly speed up library startup.</param>
    /// <param name="fontStreams">Extra fonts that should be loaded to the engine.</param>
    private DefaultGraphicEngine(Stream fallbackFontStream, bool useSystemFonts, Stream[] fontStreams)
    {
        ArgumentNullException.ThrowIfNull(fallbackFontStream);
        ArgumentNullException.ThrowIfNull(fontStreams);

        var fontCollection = new FontCollection();
        AddEmbeddedFont(fontCollection);
        var fallbackFamily = fontCollection.Add(fallbackFontStream);
        foreach (var fontStream in fontStreams)
            fontCollection.Add(fontStream);

        _fontCollection = useSystemFonts
            ? new Lazy<IReadOnlyFontCollection>(fontCollection.AddSystemFonts)
            : new Lazy<IReadOnlyFontCollection>(() => fontCollection);
        _fallbackFont = fallbackFamily.Name;
        _loadFont = LoadFont;
        _calculateMaxDigitWidth = CalculateMaxDigitWidth;
    }

    /// <summary>
    /// Create a default graphic engine that uses only fallback font and additional fonts passed as streams.
    /// It ignores all system fonts, and that can lead to a decrease of initialization time.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Font is determined by a name and style in the worksheet, but the font name must be mapped to a font file/stream.
    /// System fonts on Windows contain hundreds of font files that have to be checked to find the correct font
    /// file for the font name and style. That means to read hundreds of files and parse data inside them.
    /// Even though SixLabors.Fonts does this only once (lazily too) and stores data in a static variable, it is
    /// an overhead that can be avoided.
    /// </para>
    /// <para>
    /// This factory method is useful in several scenarios:
    /// <list type="bullet">
    ///   <item>Client side Blazor doesn't have access to any system fonts.</item>
    ///   <item>Worksheet contains only a limited number of fonts. It might be enough to just load a few fonts we are</item>
    /// </list>
    /// </para>
    /// </remarks>
    /// <param name="fallbackFontStream">A stream that contains a fallback font.</param>
    /// <param name="fontStreams">Fonts that should be loaded to the engine.</param>
    public static IXLGraphicEngine CreateOnlyWithFonts(Stream fallbackFontStream, params Stream[] fontStreams)
    {
        return new DefaultGraphicEngine(fallbackFontStream, false, fontStreams);
    }

    /// <summary>
    /// Create a default graphic engine that uses only fallback font and additional fonts passed as streams.
    /// It also uses system fonts.
    /// </summary>
    /// <param name="fallbackFontStream">A stream that contains a fallback font.</param>
    /// <param name="fontStreams">Fonts that should be loaded to the engine.</param>
    public static IXLGraphicEngine CreateWithFontsAndSystemFonts(Stream fallbackFontStream, params Stream[] fontStreams)
    {
        return new DefaultGraphicEngine(fallbackFontStream, true, fontStreams);
    }

    public XLPictureInfo GetPictureInfo(Stream imageStream, XLPictureFormat expectedFormat)
    {
        foreach (var imageReader in _imageReaders)
        {
            if (imageReader.TryGetInfo(imageStream, out var dimensions))
                return dimensions;
        }

        throw new ArgumentException("Unable to determine the format of the image.");
    }

    public double GetDescent(IXLFontBase font, double dpiY)
    {
        var metrics = GetMetrics(font);
        return GetDescent(font, dpiY, metrics);
    }

    private double GetDescent(IXLFontBase font, double dpiY, FontMetrics metrics)
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
        // SixLabors.Fonts don't have a way to get a glyph representation of a cluster
        // without a TextRenderer that has unacceptable performance.
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

            // As of SixLabors.Fonts 1.0.0, the TryGetGlyphMetrics method never fails. It returns .notdef glyph 0
            // as a fallback glyph, but it might change in the future.
            if (!containsMetrics)
                continue;

            advanceFu = glyphs!.Aggregate(advanceFu, (current, glyph) => current + glyph.AdvanceWidth);
        }

        var emInPx = font.FontSize / 72d * dpi.X;
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
        // First, try the specified fallback font. On Windows, unknown fonts should use MS Sans Serif
        if (!_fontCollection.Value.TryGet(metricId.Name, out var fontFamily) &&
            !_fontCollection.Value.TryGet(_fallbackFont, out fontFamily))
        {
            // If not present, e.g., it's unlikely to be present on Linux, use embedded font as an ultimate fallback.
            fontFamily = _fontCollection.Value.Get(EmbeddedFontName);
        }

        return fontFamily.CreateFont(FontMetricSize); // Size is irrelevant for metric
    }

    private void AddEmbeddedFont(FontCollection fontCollection)
    {
        var assembly = Assembly.GetExecutingAssembly();
        const string resourcePath = "XLibur.Graphics.Fonts.CarlitoBare-{0}.ttf";

        using var regular = assembly.GetManifestResourceStream(string.Format(resourcePath, "Regular"))!;
        fontCollection.Add(regular);

        using var bold = assembly.GetManifestResourceStream(string.Format(resourcePath, "Bold"))!;
        fontCollection.Add(bold);

        using var italic = assembly.GetManifestResourceStream(string.Format(resourcePath, "Italic"))!;
        fontCollection.Add(italic);

        using var boldItalic = assembly.GetManifestResourceStream(string.Format(resourcePath, "BoldItalic"))!;
        fontCollection.Add(boldItalic);
    }

    private double CalculateMaxDigitWidth(MetricId metricId)
    {
        var font = GetFont(metricId);
        var metrics = font.FontMetrics;
        var maxWidth = int.MinValue;
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
        private readonly FontStyle _style;

        public MetricId(IXLFontBase fontBase)
        {
            Name = fontBase.FontName;
            _style = GetFontStyle(fontBase);
        }

        public string Name { get; }

        public bool Equals(MetricId other) => Name == other.Name && _style == other._style;

        public override bool Equals(object? obj) => obj is MetricId other && Equals(other);

        public override int GetHashCode() => (Name.GetHashCode() * 397) ^ (int)_style;

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
