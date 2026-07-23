using System;
using System.Drawing;
using System.IO;
using XLibur.Excel.Drawings;
using XLibur.Utils;

namespace XLibur.Graphics;

/// <summary>
/// Read info about PCX picture.
/// https://moddingwiki.shikadi.net/wiki/PCX_Format
/// </summary>
internal sealed class PcxInfoReader : ImageInfoReader
{
    protected override bool CheckHeader(Stream stream)
    {
        // Read through the window fields so malformed dimensions are rejected here
        // rather than producing a bogus XLPictureInfo in ReadInfo.
        Span<byte> header = stackalloc byte[12];
        try
        {
            stream.ReadExactly(header);
        }
        catch (EndOfStreamException)
        {
            return false;
        }

        if (header[0] != 0xA ||
            header[1] > 5 || // version must be 0..5
            header[2] > 1) // encoding, nearly always should be 1
            return false;

        // Window bounds (XMin/YMin/XMax/YMax) as little-endian words at offsets 4..11.
        var winXMin = header[4] | header[5] << 8;
        var winYMin = header[6] | header[7] << 8;
        var winXMax = header[8] | header[9] << 8;
        var winYMax = header[10] | header[11] << 8;

        // Max must not be below Min, otherwise width/height would be zero or negative.
        return winXMax >= winXMin && winYMax >= winYMin;
    }

    protected override XLPictureInfo ReadInfo(Stream stream)
    {
        stream.Position += 4; // skip Manufacturer, Version, Encoding, BitsPerPixel
        var winXMin = stream.ReadU16LE();
        var winYMin = stream.ReadU16LE();
        var winXMax = stream.ReadU16LE();
        var winYMax = stream.ReadU16LE();
        // HDpi/VDpi are unreliable in practice: many tools wrote screen resolution or 0
        // here instead of a true DPI. XLPictureInfo tolerates 0, so pass them through as-is.
        var dpiX = stream.ReadU16LE();
        var dpiY = stream.ReadU16LE();

        var widthPx = winXMax - winXMin + 1;
        var heightPx = winYMax - winYMin + 1;
        return new XLPictureInfo(XLPictureFormat.Pcx, new Size(widthPx, heightPx), Size.Empty, dpiX, dpiY);
    }
}
