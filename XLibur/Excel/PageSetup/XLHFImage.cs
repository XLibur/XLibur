using System;
using System.Drawing;
using System.IO;
using XLibur.Excel.Drawings;
using XLibur.Graphics;

namespace XLibur.Excel;

/// <summary>
/// Represents an image in a header or footer section.
/// </summary>
internal sealed class XLHFImage
{
    private const double MaxWidthInches = 2.5;
    private const double MaxHeightInches = 0.6;
    private const double DefaultDpi = 96.0;

    /// <summary>
    /// The binary image data.
    /// </summary>
    public MemoryStream ImageStream { get; }

    /// <summary>
    /// Detected image format.
    /// </summary>
    public XLPictureFormat Format { get; }

    /// <summary>
    /// Scaled width in points (1/72 inch).
    /// </summary>
    public double WidthPt { get; }

    /// <summary>
    /// Scaled height in points (1/72 inch).
    /// </summary>
    public double HeightPt { get; }

    /// <summary>
    /// The header/footer position code for VML shape id (e.g. "LH", "CF", "RH").
    /// Set during save when position is known.
    /// </summary>
    public string? PositionCode { get; set; }

    internal XLHFImage(Stream stream, XLPictureInfo info)
    {
        Format = info.Format;

        var dpiX = info.DpiX > 0 ? info.DpiX : DefaultDpi;
        var dpiY = info.DpiY > 0 ? info.DpiY : DefaultDpi;
        var sizePx = info.GetSizePx(dpiX, dpiY);

        // Convert pixels to inches
        var widthInches = sizePx.Width / dpiX;
        var heightInches = sizePx.Height / dpiY;

        // Scale down to fit within max dimensions, preserving aspect ratio.
        // Do not upscale.
        var scale = 1.0;
        if (widthInches > MaxWidthInches)
            scale = Math.Min(scale, MaxWidthInches / widthInches);
        if (heightInches > MaxHeightInches)
            scale = Math.Min(scale, MaxHeightInches / heightInches);

        WidthPt = widthInches * scale * 72.0;
        HeightPt = heightInches * scale * 72.0;

        ImageStream = new MemoryStream();
        stream.Position = 0;
        stream.CopyTo(ImageStream);
        ImageStream.Position = 0;
    }

    /// <summary>
    /// Creates an <see cref="XLHFImage"/> from a file path, using the workbook's graphic engine for format detection.
    /// </summary>
    internal static XLHFImage FromFile(string imagePath, XLWorkbook workbook)
    {
        if (string.IsNullOrWhiteSpace(imagePath))
            throw new ArgumentException("Image path cannot be null or empty.", nameof(imagePath));

        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"Header/footer image file not found: '{imagePath}'", imagePath);

        using var fs = File.OpenRead(imagePath);
        var info = workbook.GraphicEngine.GetPictureInfo(fs, XLPictureFormat.Unknown);

        ValidateFormat(info.Format, imagePath);

        fs.Position = 0;
        return new XLHFImage(fs, info);
    }

    private static void ValidateFormat(XLPictureFormat format, string imagePath)
    {
        switch (format)
        {
            case XLPictureFormat.Png:
            case XLPictureFormat.Jpeg:
            case XLPictureFormat.Bmp:
            case XLPictureFormat.Gif:
            case XLPictureFormat.Tiff:
            case XLPictureFormat.Emf:
            case XLPictureFormat.Wmf:
                return;
            case XLPictureFormat.Unknown:
                throw new ArgumentException(
                    $"Unable to detect image format for header/footer image: '{imagePath}'. Supported formats: PNG, JPEG, BMP, GIF, TIFF, EMF, WMF.",
                    nameof(imagePath));
            default:
                throw new ArgumentException(
                    $"Unsupported image format '{format}' for header/footer image: '{imagePath}'. Supported formats: PNG, JPEG, BMP, GIF, TIFF, EMF, WMF.",
                    nameof(imagePath));
        }
    }
}
