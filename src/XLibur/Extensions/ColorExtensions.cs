
using System;
using System.Drawing;

namespace XLibur.Extensions;

internal static class ColorExtensions
{
    private static readonly char[] HexDigits = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F'];

    public static string ToHex(this Color color)
    {
        Span<char> chars = stackalloc char[8];
        WriteHexByte(chars, 0, color.A);
        WriteHexByte(chars, 2, color.R);
        WriteHexByte(chars, 4, color.G);
        WriteHexByte(chars, 6, color.B);
        return new string(chars);

        static void WriteHexByte(Span<char> dest, int offset, byte b)
        {
            dest[offset] = HexDigits[b >> 4];
            dest[offset + 1] = HexDigits[b & 0xF];
        }
    }
}
