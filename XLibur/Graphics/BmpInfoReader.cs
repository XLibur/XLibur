using System;
using System.Drawing;
using System.IO;
using XLibur.Excel.Drawings;
using XLibur.Utils;

namespace XLibur.Graphics;

/// <summary>
/// A reader for BMP for Windows and OS/2.
/// Specification:
/// https://www.fileformat.info/format/bmp/corion.htm
/// https://www.fileformat.info/format/bmp/egff.htm
/// https://www.fileformat.info/format/os2bmp/egff.htm
/// </summary>
internal sealed class BmpInfoReader : ImageInfoReader
{
    private const int BitmapFileHeaderSize = 14;
    private const int BitmapArrayHeaderSize = 14;

    protected override bool CheckHeader(Stream stream)
    {
        Span<byte> s = stackalloc byte[16];
        if (stream.Read(s) != s.Length)
            return false;

        if (s[0] != 'B')
            return false;

        // Plain Windows/OS_2 bitmap: BITMAPFILEHEADER starts at offset 0.
        if (s[1] == 'M')
            return true;

        // OS/2 "BA" Bitmap Array: a 14-byte BITMAPARRAYHEADER wraps one or more bitmaps.
        // Only decode arrays whose first entry is a plain "BM" bitmap; the icon/pointer
        // variants (CI, CP, IC, PT) can't be read by Excel, so leave them undetected.
        return s[1] == 'A' && s[BitmapArrayHeaderSize] == 'B' && s[BitmapArrayHeaderSize + 1] == 'M';
    }

    protected override XLPictureInfo ReadInfo(Stream stream)
    {
        // OS/2 "BA" files prepend a BITMAPARRAYHEADER before the first BITMAPFILEHEADER;
        // plain "BM" files start with the BITMAPFILEHEADER directly.
        Span<byte> signature = stackalloc byte[2];
        stream.ReadExactly(signature);
        var fileHeaderStart = signature[1] == 'A' ? BitmapArrayHeaderSize : 0;

        stream.Position = fileHeaderStart + BitmapFileHeaderSize;
        var infoHeaderSize = stream.ReadS32LE();
        // BMP Version 1.x, used by IBM OS/2 1.x and Win 2.0 and later
        return infoHeaderSize == 12 ? ReadBmpV1X(stream) :
            // BMP Version 2.x used by IBM OS/2 has a different overall structure, but width/height and resolution have same offsets as V3.x
            // BMP Version 3.x has dimension and resolution at same offsets and V4.x+ only add fields
            ReadBmpV2X(stream);
    }

    private static XLPictureInfo ReadBmpV1X(Stream stream)
    {
        var widthPx = stream.ReadU16LE();
        var heightPx = stream.ReadU16LE();
        return new XLPictureInfo(XLPictureFormat.Bmp, new Size(widthPx, heightPx), Size.Empty);
    }

    private static XLPictureInfo ReadBmpV2X(Stream stream)
    {
        var widthPx = stream.ReadU32LE();
        var heightPx = stream.ReadU32LE();
        stream.Position += 12;
        var dpiX = PixelsPerMeterToDpi(stream.ReadU32LE());
        var dpiY = PixelsPerMeterToDpi(stream.ReadU32LE());
        return new XLPictureInfo(XLPictureFormat.Bmp, widthPx, heightPx, dpiX, dpiY);
    }

    private static double PixelsPerMeterToDpi(uint pixelsPerMeter)
        => pixelsPerMeter * 2.54d / 100d;
}
