using System;
using System.Buffers;
using System.IO;

namespace XLibur.Excel.IO;

/// <summary>
/// Copies a worksheet part with its <c>&lt;sheetData&gt;</c> reduced to an empty element, reading the
/// part as a stream, for a part too large for <see cref="WorksheetPartBuffer"/> (#723). Only the
/// markup around the cells is held, never the cells.
/// </summary>
/// <remarks>
/// The copy is the one <see cref="WorksheetPartBuffer.CopyWithoutSheetData"/> makes of a part it
/// buffers, byte for byte, and <c>&lt;sheetData&gt;</c> is placed by the same rules
/// (<see cref="SheetDataLocator"/>): any part the locator would refuse is refused here too.
/// <para>
/// The cells are scanned in windows, by byte search, not tokenised. Each window starts with the last
/// bytes of the one before, enough that a pattern split between two reads, such as the end tag, is
/// found whole in the next window. A pattern found twice that way is one that refuses the part, so
/// finding it twice does no harm.
/// </para>
/// </remarks>
internal static class PartWithoutSheetDataReader
{
    /// <summary>
    /// The longest copy kept. A part whose markup outside the cells is longer is refused, so what a
    /// sheet keeps stays below what <see cref="WorksheetPartBuffer"/> would hold for a smaller part.
    /// </summary>
    internal const int MaxCopyLength = WorksheetPartBuffer.MaxBufferedLength;

    private const int DefaultWindowLength = 64 * 1024;

    /// <summary>
    /// The most whitespace accepted between the name of the end tag and its <c>&gt;</c>, so the bytes
    /// carried from one window to the next stay bounded.
    /// </summary>
    private const int MaxEndTagWhitespace = 256;

    /// <summary>
    /// Reads <paramref name="part"/> to its end and returns the copy, or null when
    /// <see cref="SheetDataLocator"/> would not place <c>&lt;sheetData&gt;</c> exactly, or when the
    /// copy would be longer than <see cref="MaxCopyLength"/>.
    /// </summary>
    /// <param name="part">The inflated part, read from its start.</param>
    /// <param name="prefix">The prefix <c>&lt;sheetData&gt;</c> is written with, empty for none.</param>
    internal static byte[]? TryCopy(Stream part, out string prefix) => TryCopy(part, DefaultWindowLength, out prefix);

    /// <inheritdoc cref="TryCopy(Stream, out string)"/>
    /// <param name="part">The inflated part, read from its start.</param>
    /// <param name="windowLength">How many bytes to read at a time; small in tests, to split patterns.</param>
    /// <param name="prefix">The prefix <c>&lt;sheetData&gt;</c> is written with, empty for none.</param>
    internal static byte[]? TryCopy(Stream part, int windowLength, out string prefix)
    {
        prefix = string.Empty;
        var head = ArrayPool<byte>.Shared.Rent(windowLength);
        byte[]? window = null;
        try
        {
            if (!TryReadStartTag(part, ref head, out var headLength, out var atEnd, out var tag))
                return null;

            var xml = head.AsSpan(0, headLength);
            using var copy = new MemoryStream();

            // Already empty: the copy is the whole part.
            if (tag.IsEmptyElement(xml))
            {
                copy.Write(xml);
            }
            else
            {
                // Everything before the start tag's '>', then '/>' in its place.
                copy.Write(xml[..(tag.End - 1)]);
                copy.Write("/>"u8);

                ReadOnlySpan<byte> name = xml.Slice(tag.NameStart, tag.NameLength).ToArray();
                window = ArrayPool<byte>.Shared.Rent(windowLength + name.Length + 2 + MaxEndTagWhitespace);
                if (!TrySkipContent(part, xml[tag.End..], atEnd, name, window, copy))
                    return null;
            }

            if (!TryCopyRest(part, head, copy))
                return null;

            prefix = tag.Prefix;
            return copy.ToArray();
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(head);
            if (window is not null)
                ArrayPool<byte>.Shared.Return(window);
        }
    }

    /// <summary>
    /// Reads the part into <paramref name="head"/> until it holds the start tag of
    /// <c>&lt;sheetData&gt;</c>, doubling the buffer as it fills, so the start is searched again only
    /// a few times.
    /// </summary>
    private static bool TryReadStartTag(Stream part, ref byte[] head, out int headLength, out bool atEnd,
        out SheetDataLocator.StartTag tag)
    {
        headLength = 0;
        while (true)
        {
            if (headLength == head.Length)
            {
                if (head.Length >= MaxCopyLength)
                {
                    atEnd = false;
                    tag = default;
                    return false;
                }

                var larger = ArrayPool<byte>.Shared.Rent(head.Length * 2);
                head.AsSpan(0, headLength).CopyTo(larger);
                ArrayPool<byte>.Shared.Return(head);
                head = larger;
            }

            var wanted = head.Length - headLength;
            var read = part.ReadAtLeast(head.AsSpan(headLength), wanted, throwOnEndOfStream: false);
            headLength += read;
            atEnd = read < wanted;

            var result = SheetDataLocator.LocateStartTag(head.AsSpan(0, headLength), out tag);
            if (result == SheetDataLocator.StartTagResult.Found)
                return true;

            if (result == SheetDataLocator.StartTagResult.Refused || atEnd)
                return false;
        }
    }

    /// <summary>
    /// Reads past the content and the end tag of <c>&lt;sheetData&gt;</c>, and copies what follows
    /// the end tag in the last window read.
    /// </summary>
    /// <param name="part">The part, read up to the end of <paramref name="first"/>.</param>
    /// <param name="first">The content the head already holds, past the start tag.</param>
    /// <param name="atEnd">Has the part been read to its end?</param>
    /// <param name="name">The qualified name of the element.</param>
    /// <param name="window">The buffer to read into; longer than a window by the most carried.</param>
    /// <param name="copy">The copy so far.</param>
    private static bool TrySkipContent(Stream part, ReadOnlySpan<byte> first, bool atEnd, ReadOnlySpan<byte> name,
        byte[] window, MemoryStream copy)
    {
        var span = first;
        while (true)
        {
            switch (FindEndTag(span, name, atEnd, out var at))
            {
                case Scan.Found:
                    copy.Write(span[at..]);
                    return copy.Length <= MaxCopyLength;
                case Scan.Refused:
                    return false;
            }

            // Carry the bytes from 'at' into the next window, which may be the window itself.
            var carried = span.Length - at;
            span[at..].CopyTo(window);
            var wanted = window.Length - carried;
            var read = part.ReadAtLeast(window.AsSpan(carried), wanted, throwOnEndOfStream: false);
            atEnd = read < wanted;
            span = window.AsSpan(0, carried + read);
        }
    }

    private enum Scan
    {
        Found,
        NeedMore,
        Refused,
    }

    /// <summary>
    /// Looks for the end tag of <c>&lt;sheetData&gt;</c> in <paramref name="span"/>, a stretch of its
    /// content, refusing content <see cref="SheetDataLocator.IsPlainContent"/> refuses.
    /// </summary>
    /// <param name="span">Content of the element, possibly followed by its end tag and more.</param>
    /// <param name="name">The qualified name of the element.</param>
    /// <param name="isLast">Is <paramref name="span"/> the end of the part?</param>
    /// <param name="at">
    /// When found, the index just past the end tag. When more is needed, the index the next window
    /// has to start from.
    /// </param>
    private static Scan FindEndTag(ReadOnlySpan<byte> span, ReadOnlySpan<byte> name, bool isLast, out int at)
    {
        at = 0;
        Span<byte> endTag = stackalloc byte[name.Length + 2];
        endTag[0] = (byte)'<';
        endTag[1] = (byte)'/';
        name.CopyTo(endTag[2..]);

        var searchFrom = 0;
        while (true)
        {
            var found = span[searchFrom..].IndexOf(endTag);
            if (found < 0)
            {
                if (isLast || !SheetDataLocator.IsPlainContent(span, name))
                    return Scan.Refused;

                // An end tag or a refused pattern split by the window's end starts in its last
                // bytes; every pattern is at most as long as the end tag.
                at = Math.Max(0, span.Length - (endTag.Length - 1));
                return Scan.NeedMore;
            }

            var close = searchFrom + found;
            var gt = close + endTag.Length;
            while (gt < span.Length && IsWhitespace(span[gt]))
                gt++;

            if (gt == span.Length)
            {
                if (isLast || gt - close > endTag.Length + MaxEndTagWhitespace
                           || !SheetDataLocator.IsPlainContent(span[..close], name))
                {
                    return Scan.Refused;
                }

                at = close;
                return Scan.NeedMore;
            }

            if (span[gt] == (byte)'>')
            {
                if (!SheetDataLocator.IsPlainContent(span[..close], name))
                    return Scan.Refused;

                at = gt + 1;
                return Scan.Found;
            }

            // A longer name that starts with this one; the content check refuses its start tag.
            searchFrom = close + 1;
        }
    }

    /// <summary>Copies the rest of the part, refusing a copy longer than <see cref="MaxCopyLength"/>.</summary>
    private static bool TryCopyRest(Stream part, byte[] buffer, MemoryStream copy)
    {
        while (true)
        {
            if (copy.Length > MaxCopyLength)
                return false;

            var read = part.Read(buffer, 0, buffer.Length);
            if (read == 0)
                return true;

            copy.Write(buffer, 0, read);
        }
    }

    private static bool IsWhitespace(byte b) => b is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n';
}
