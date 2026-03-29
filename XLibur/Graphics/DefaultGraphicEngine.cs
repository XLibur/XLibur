using System;
using System.Collections.Generic;
using System.IO;
using XLibur.Excel;
using XLibur.Excel.Drawings;

namespace XLibur.Graphics;

public class DefaultGraphicEngine : IXLGraphicEngine, IXLFontEngine
{
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

    private readonly Dictionary<XLPictureFormat, ImageInfoReader> _readersByFormat;

    private readonly IXLFontEngine _fontEngine;

    /// <summary>
    /// Get a singleton instance of the engine that uses <c>Microsoft Sans Serif</c> as a fallback font.
    /// </summary>
    public static Lazy<DefaultGraphicEngine> Instance { get; } = new(() => new DefaultGraphicEngine("Microsoft Sans Serif"));

    /// <summary>
    /// Initialize a new instance of the engine.
    /// </summary>
    /// <param name="fallbackFont">A name of a font that is used when a font in a workbook is not available.</param>
    public DefaultGraphicEngine(string fallbackFont)
    {
        _fontEngine = DefaultFontEngineFactory.CreateFromFallbackFont(fallbackFont);
        _readersByFormat = BuildReadersByFormat(_imageReaders);
    }

    /// <summary>
    /// Initialize a new instance of the engine with an explicit font engine.
    /// </summary>
    /// <param name="fontEngine">The font engine to use for text measurement and font metrics.</param>
    public DefaultGraphicEngine(IXLFontEngine fontEngine)
    {
        ArgumentNullException.ThrowIfNull(fontEngine);
        _fontEngine = fontEngine;
        _readersByFormat = BuildReadersByFormat(_imageReaders);
    }

    private DefaultGraphicEngine(Stream fallbackFontStream, bool useSystemFonts, Stream[] fontStreams)
    {
        _fontEngine = DefaultFontEngineFactory.CreateFromStreams(fallbackFontStream, useSystemFonts, fontStreams);
        _readersByFormat = BuildReadersByFormat(_imageReaders);
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

    XLPictureInfo IXLGraphicEngine.GetPictureInfo(Stream imageStream, XLPictureFormat expectedFormat)
        => GetPictureInfo(imageStream, expectedFormat);

    public XLPictureInfo GetPictureInfo(Stream stream, XLPictureFormat expectedFormat)
    {
        if (expectedFormat != XLPictureFormat.Unknown
            && _readersByFormat.TryGetValue(expectedFormat, out var preferredReader)
            && preferredReader.TryGetInfo(stream, out var info))
        {
            return info;
        }

        foreach (var imageReader in _imageReaders)
        {
            if (imageReader.TryGetInfo(stream, out var dimensions))
                return dimensions;
        }

        throw new ArgumentException("Unable to determine the format of the image.");
    }

    public double GetDescent(IXLFontBase font, double dpiY) => _fontEngine.GetDescent(font, dpiY);

    double IXLGraphicEngine.GetMaxDigitWidth(IXLFontBase font, double dpiX)
        => GetMaxDigitWidth(font, dpiX);

    public double GetMaxDigitWidth(IXLFontBase fontBase, double dpiX) => _fontEngine.GetMaxDigitWidth(fontBase, dpiX);

    public double GetTextHeight(IXLFontBase font, double dpiY) => _fontEngine.GetTextHeight(font, dpiY);

    double IXLGraphicEngine.GetTextWidth(string text, IXLFontBase font, double dpiX)
        => GetTextWidth(text, font, dpiX);

    public double GetTextWidth(string text, IXLFontBase fontBase, double dpiX) => _fontEngine.GetTextWidth(text, fontBase, dpiX);

    /// <inheritdoc />
    public GlyphBox GetGlyphBox(ReadOnlySpan<int> graphemeCluster, IXLFontBase font, Dpi dpi)
        => _fontEngine.GetGlyphBox(graphemeCluster, font, dpi);

    private static Dictionary<XLPictureFormat, ImageInfoReader> BuildReadersByFormat(ImageInfoReader[] readers)
    {
        var map = new Dictionary<XLPictureFormat, ImageInfoReader>();
        foreach (var reader in readers)
        {
            var format = reader switch
            {
                PngInfoReader => XLPictureFormat.Png,
                JpegInfoReader => XLPictureFormat.Jpeg,
                GifInfoReader => XLPictureFormat.Gif,
                TiffInfoReader => XLPictureFormat.Tiff,
                BmpInfoReader => XLPictureFormat.Bmp,
                EmfInfoReader => XLPictureFormat.Emf,
                WmfInfoReader => XLPictureFormat.Wmf,
                WebpInfoReader => XLPictureFormat.Webp,
                SvgInfoReader => XLPictureFormat.Svg,
                PcxInfoReader => XLPictureFormat.Pcx,
                _ => XLPictureFormat.Unknown
            };

            if (format != XLPictureFormat.Unknown)
                map[format] = reader;
        }

        return map;
    }
}
