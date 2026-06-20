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
    /// Initialize a new instance of the engine with an explicit font engine.
    /// </summary>
    /// <param name="fontEngine">The font engine to use for text measurement and font metrics.</param>
    public DefaultGraphicEngine(IXLFontEngine fontEngine)
    {
        ArgumentNullException.ThrowIfNull(fontEngine);
        _fontEngine = fontEngine;
        _readersByFormat = BuildReadersByFormat(_imageReaders);
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

    public double GetMaxDigitWidth(IXLFontBase font, double dpiX) => _fontEngine.GetMaxDigitWidth(font, dpiX);

    public double GetTextHeight(IXLFontBase font, double dpiY) => _fontEngine.GetTextHeight(font, dpiY);

    double IXLGraphicEngine.GetTextWidth(string text, IXLFontBase font, double dpiX)
        => GetTextWidth(text, font, dpiX);

    public double GetTextWidth(string text, IXLFontBase font, double dpiX) => _fontEngine.GetTextWidth(text, font, dpiX);

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
