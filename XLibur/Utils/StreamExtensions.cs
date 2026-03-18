using System;
using System.IO;

namespace XLibur.Utils;

internal static class StreamExtensions
{
    // ReSharper disable once InconsistentNaming
    extension(Stream stream)
    {
        public int ReadS32LE()
        {
            var b1 = stream.ReadU8();
            var b2 = stream.ReadU8();
            var b3 = stream.ReadU8();
            var b4 = stream.ReadU8();
            return b4 << 24 | b3 << 16 | b2 << 8 | b1;
        }

        // ReSharper disable once InconsistentNaming
        public short ReadS16LE()
        {
            var b1 = stream.ReadU8();
            var b2 = stream.ReadU8();
            return (short)((b2 << 8) | b1);
        }

        // ReSharper disable once InconsistentNaming
        public ushort ReadU16BE()
        {
            return !stream.TryReadU16BE(out var number) ? throw EndOfStreamException() : number;
        }

        // ReSharper disable once InconsistentNaming
        public uint ReadU32BE()
        {
            return !stream.TryReadU32BE(out var number) ? throw EndOfStreamException() : number;
        }

        // ReSharper disable once InconsistentNaming
        public uint ReadU32LE()
        {
            return !stream.TryReadU32LE(out var number) ? throw EndOfStreamException() : number;
        }

        // ReSharper disable once InconsistentNaming
        public bool TryReadU32LE(out uint number)
        {
            if (!TryReadLE(stream, 4, out var result))
            {
                number = 0;
                return false;
            }

            number = (uint)result;
            return true;
        }

        // ReSharper disable once InconsistentNaming
        public ushort ReadU16LE()
        {
            return !stream.TryReadU16LE(out var number) ? throw EndOfStreamException() : number;
        }

        // ReSharper disable once InconsistentNaming
        public bool TryReadU16LE(out ushort number)
        {
            if (!TryReadLE(stream, 2, out var result))
            {
                number = 0;
                return false;
            }

            number = (ushort)result;
            return true;
        }

        // ReSharper disable once InconsistentNaming
        public int ReadU24BE()
        {
            return !TryReadBE(stream, 3, out var result) ? throw EndOfStreamException() : result;
        }

        // ReSharper disable once InconsistentNaming
        public int ReadU24LE()
        {
            return !TryReadLE(stream, 3, out var result) ? throw EndOfStreamException() : result;
        }

        public byte ReadU8()
        {
            var b = stream.ReadByte();
            if (b == -1)
                throw EndOfStreamException();
            return (byte)b;
        }

        // ReSharper disable once InconsistentNaming
        public bool TryReadU32BE(out uint number)
        {
            if (!TryReadBE(stream, 4, out var readNumber))
            {
                number = 0;
                return false;
            }

            number = (uint)readNumber;
            return true;
        }

        // ReSharper disable once InconsistentNaming
        public bool TryReadU16BE(out ushort number)
        {
            if (TryReadBE(stream, 2, out var readNumber))
            {
                number = (ushort)readNumber;
                return true;
            }

            number = 0;
            return false;
        }
    }

    // ReSharper disable once InconsistentNaming
    private static bool TryReadLE(Stream stream, int size, out int number)
    {
        number = 0;
        for (var i = 0; i < size; ++i)
        {
            var readByte = stream.ReadByte();
            if (readByte == -1)
                return false;

            number |= readByte << i * 8;
        }

        return true;
    }

    // ReSharper disable once InconsistentNaming
    private static bool TryReadBE(Stream stream, int size, out int number)
    {
        number = 0;
        for (var i = 1; i <= size; ++i)
        {
            var readByte = stream.ReadByte();
            if (readByte == -1)
                return false;

            number |= readByte << (size - i) * 8;
        }

        return true;
    }

    private static ArgumentException EndOfStreamException() => new("Unexpected end of stream.");
}
