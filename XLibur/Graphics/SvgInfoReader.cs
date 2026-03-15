using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using XLibur.Excel.Drawings;

namespace XLibur.Graphics;

/// <summary>
/// Reader of dimensions for SVG image format.
/// </summary>
internal sealed partial class SvgInfoReader : ImageInfoReader
{
    /// <summary>
    /// Default DPI for SVG images when converting physical units to pixels.
    /// </summary>
    private const double DefaultDpi = 96;

    /// <summary>
    /// Maximum number of bytes to read when looking for the &lt;svg&gt; element.
    /// </summary>
    private const int MaxHeaderBytes = 4096;

    protected override bool CheckHeader(Stream stream)
    {
        // SVG files are XML. They may start with an XML declaration, BOM, or whitespace.
        // Look for '<svg' or '<?xml' within the first chunk of the file.
        Span<byte> buffer = stackalloc byte[Math.Min(MaxHeaderBytes, (int)Math.Min(stream.Length, MaxHeaderBytes))];
        var bytesRead = stream.Read(buffer);
        if (bytesRead == 0)
            return false;

        var text = Encoding.UTF8.GetString(buffer[..bytesRead]);

        // Skip BOM and whitespace, look for XML declaration or svg element
        var trimmed = text.TrimStart('\uFEFF').TrimStart();
        return trimmed.StartsWith("<?xml", StringComparison.OrdinalIgnoreCase) &&
               text.Contains("<svg", StringComparison.OrdinalIgnoreCase)
               || trimmed.StartsWith("<svg", StringComparison.OrdinalIgnoreCase);
    }

    protected override XLPictureInfo ReadInfo(Stream stream)
    {
        // Read enough of the SVG to find the <svg> element and its attributes.
        var buffer = new byte[Math.Min(MaxHeaderBytes, (int)Math.Min(stream.Length, MaxHeaderBytes))];
        var bytesRead = stream.Read(buffer, 0, buffer.Length);
        var text = Encoding.UTF8.GetString(buffer, 0, bytesRead);

        var svgMatch = SvgElementRegex().Match(text);
        if (!svgMatch.Success)
            throw new ArgumentException("Unable to find <svg> element in SVG file.");

        var svgTag = svgMatch.Value;

        // Try to get width/height attributes first
        var widthAttr = GetAttribute(svgTag, "width");
        var heightAttr = GetAttribute(svgTag, "height");

        if (widthAttr is not null && heightAttr is not null &&
            TryParseSvgLength(widthAttr, out var widthPx) &&
            TryParseSvgLength(heightAttr, out var heightPx))
        {
            return new XLPictureInfo(XLPictureFormat.Svg, new Size((int)Math.Ceiling(widthPx), (int)Math.Ceiling(heightPx)), Size.Empty, DefaultDpi, DefaultDpi);
        }

        // Fall back to viewBox attribute
        var viewBox = GetAttribute(svgTag, "viewBox");
        if (viewBox is not null)
        {
            var parts = viewBox.Split([' ', ','], StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 4 &&
                double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out var vbWidth) &&
                double.TryParse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture, out var vbHeight))
            {
                return new XLPictureInfo(XLPictureFormat.Svg, new Size((int)Math.Ceiling(vbWidth), (int)Math.Ceiling(vbHeight)), Size.Empty, DefaultDpi, DefaultDpi);
            }
        }

        // Default fallback if no dimensions found
        return new XLPictureInfo(XLPictureFormat.Svg, new Size(300, 150), Size.Empty, DefaultDpi, DefaultDpi);
    }

    private static string? GetAttribute(string element, string name)
    {
        // Match attribute="value" or attribute='value'
        var pattern = $@"\b{name}\s*=\s*[""']([^""']*)[""']";
        var match = Regex.Match(element, pattern, RegexOptions.IgnoreCase);
        return match.Success ? match.Groups[1].Value : null;
    }

    private static bool TryParseSvgLength(string value, out double pixels)
    {
        pixels = 0;
        value = value.Trim();

        // Try plain number (interpreted as pixels by convention in user-agent context)
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out pixels))
            return pixels > 0;

        // Try with unit suffix
        if (value.Length < 2)
            return false;

        ReadOnlySpan<char> span = value;
        string? unit = null;
        ReadOnlySpan<char> numberPart = default;

        if (span.EndsWith("px", StringComparison.OrdinalIgnoreCase))
        {
            unit = "px";
            numberPart = span[..^2];
        }
        else if (span.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
        {
            unit = "pt";
            numberPart = span[..^2];
        }
        else if (span.EndsWith("in", StringComparison.OrdinalIgnoreCase))
        {
            unit = "in";
            numberPart = span[..^2];
        }
        else if (span.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
        {
            unit = "cm";
            numberPart = span[..^2];
        }
        else if (span.EndsWith("mm", StringComparison.OrdinalIgnoreCase))
        {
            unit = "mm";
            numberPart = span[..^2];
        }
        else if (span.EndsWith("em", StringComparison.OrdinalIgnoreCase))
        {
            unit = "em";
            numberPart = span[..^2];
        }

        if (unit is null)
            return false;

        if (!double.TryParse(numberPart.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var numValue) || numValue <= 0)
            return false;

        pixels = unit switch
        {
            "px" => numValue,
            "pt" => numValue * DefaultDpi / 72.0,
            "in" => numValue * DefaultDpi,
            "cm" => numValue * DefaultDpi / 2.54,
            "mm" => numValue * DefaultDpi / 25.4,
            "em" => numValue * 16.0, // 1em = 16px default
            _ => numValue
        };

        return true;
    }

    [GeneratedRegex(@"<svg\b[^>]*>", RegexOptions.IgnoreCase | RegexOptions.Singleline)]
    private static partial Regex SvgElementRegex();
}
