using System;
using System.Drawing;
using System.IO;
using System.Linq;
using ClosedXML.Excel.Drawings;
using ClosedXML.Utils;

namespace XLibur.Graphics;

/// <summary>
/// Read <a href="https://www.w3.org/Graphics/JPEG/jfif3.pdf">JFIF</a> or EXIF.
/// </summary>
internal sealed class JpegInfoReader : ImageInfoReader
{
    private static readonly byte[] APP0Identifer = "JFIF\0"u8.ToArray();
    private static readonly byte[] APP1Identifer = "Exif\0\0"u8.ToArray();
    private static readonly byte[] APP14Identifer = "Adobe\0"u8.ToArray();

    protected override bool CheckHeader(Stream stream)
    {
        if (!stream.TryReadU16BE(out var marker) || marker != Marker.SOI)
            return false;

        // Per spec, APP0 should be the first marker, but there are many sloopy encoders
        while (TryGetMarker(stream, out marker) && TryGetLength(stream, out var length))
        {
            switch (marker)
            {
                case Marker.APP0:
                    return IsIdentifier(stream, APP0Identifer);
                case Marker.APP1:
                    return IsIdentifier(stream, APP1Identifer);
                case Marker.APP14:
                    return IsIdentifier(stream, APP14Identifer);
                default:
                    stream.Position += length;
                    break;
            }
        }

        return false;

        static bool IsIdentifier(Stream stream, byte[] identifer)
        {
            return !(from t in identifer let b = stream.ReadByte() where b == -1 || (byte)b != t select t).Any();
        }
    }

    protected override XLPictureInfo ReadInfo(Stream stream)
    {
        stream.Position += 2;
        double xDpi = 0, yDpi = 0;
        while (TryGetMarker(stream, out var marker) && TryGetLength(stream, out var length))
        {
            var segmentStart = stream.Position;
            if (marker == Marker.APP0)
            {
                const int versionLength = 2;
                stream.Position += APP0Identifer.Length + versionLength;

                var units = stream.ReadU8();
                var xDensity = stream.ReadU16BE();
                var yDensity = stream.ReadU16BE();

                xDpi = ConvertToDpi(xDensity, units);
                yDpi = ConvertToDpi(yDensity, units);
            }
            else if (Marker.SOFx.Contains(marker))
            {
                const int samplePrecisionLength = 1;
                stream.Position += samplePrecisionLength;
                var height = stream.ReadU16BE();
                var width = stream.ReadU16BE();

                // End here, before we get to SOS segment that doesn't contain explicit segment length
                return new XLPictureInfo(XLPictureFormat.Jpeg, new Size(width, height), Size.Empty, xDpi, yDpi);
            }

            stream.Position = segmentStart + length;
        }

        throw new ArgumentException("SOF not found in the JFIF.");
    }

    private bool TryGetMarker(Stream stream, out ushort marker)
    {
        if (!stream.TryReadU16BE(out marker))
            return false;

        return marker >> 8 == 0xFF;
    }

    private bool TryGetLength(Stream stream, out ushort length)
    {
        if (!stream.TryReadU16BE(out length))
            return false;

        length -= 2;
        return true;
    }
    private double ConvertToDpi(int density, byte units)
    {
        return units switch
        {
            DensityUnits.DotsPerInch => density,
            DensityUnits.DotsPerCm => density * 2.54d,
            _ => 0d
        };
    }

    private static class Marker
    {
        public const ushort SOI = 0xFFD8;
        public const ushort APP0 = 0xFFE0;
        public const ushort APP1 = 0xFFE1;
        public const ushort APP14 = 0xFFEE;
        public static readonly ushort[] SOFx = [0xFFC0, 0xFFC1, 0xFFC2, 0xFFC3, 0xFFC5, 0xFFC6, 0xFFC7, 0xFFC9, 0xFFCA, 0xFFCB, 0xFFCD, 0xFFCE, 0xFFCF];
    }

    private static class DensityUnits
    {
        public const byte DotsPerInch = 1;
        public const byte DotsPerCm = 2;
    }
}
