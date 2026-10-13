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
/// buffers, byte for byte. <c>&lt;sheetData&gt;</c> is placed by <see cref="SheetDataLocator"/>'s own
/// searches, fed a stretch of the part at a time, so any part the locator would refuse is refused
/// here too.
/// <para>
/// The cells are scanned in windows, by byte search, not tokenised. Each window starts with the bytes
/// of the one before that <see cref="SheetDataLocator.FindEndTag"/> asks to see again, enough that a
/// pattern split between two reads, such as the end tag, is found whole.
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
    /// Reads <paramref name="part"/> to its end and returns the copy, or null when
    /// <see cref="SheetDataLocator"/> would not place <c>&lt;sheetData&gt;</c> exactly, or when the
    /// copy would be longer than <see cref="MaxCopyLength"/>. Stops reading as soon as it refuses.
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

                Span<byte> endTag = stackalloc byte[tag.NameLength + 2];
                SheetDataLocator.WriteEndTag(xml.Slice(tag.NameStart, tag.NameLength), endTag);

                // Room for a window and the bytes carried into it, never longer than the end tag.
                window = ArrayPool<byte>.Shared.Rent(windowLength + endTag.Length);
                if (!TrySkipContent(part, xml[tag.End..], atEnd, endTag, window, copy))
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
            if (result == SheetDataLocator.ScanResult.Found)
                return true;

            if (result == SheetDataLocator.ScanResult.Refused || atEnd)
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
    /// <param name="endTag"><c>&lt;/</c> and the qualified name of the element.</param>
    /// <param name="window">The buffer to read into, a window plus <paramref name="endTag"/> long.</param>
    /// <param name="copy">The copy so far.</param>
    private static bool TrySkipContent(Stream part, ReadOnlySpan<byte> first, bool atEnd, ReadOnlySpan<byte> endTag,
        byte[] window, MemoryStream copy)
    {
        var span = first;
        while (true)
        {
            switch (SheetDataLocator.FindEndTag(span, endTag, atEnd, out var at, out var carried))
            {
                case SheetDataLocator.ScanResult.Found:
                    copy.Write(span[at..]);
                    return copy.Length <= MaxCopyLength;
                case SheetDataLocator.ScanResult.Refused:
                    return false;
            }

            // Carry what the search asks to see again into the next window. The span can be the
            // window itself, which a span copy allows.
            span.Slice(at, carried).CopyTo(window);
            var wanted = window.Length - carried;
            var read = part.ReadAtLeast(window.AsSpan(carried), wanted, throwOnEndOfStream: false);
            atEnd = read < wanted;
            span = window.AsSpan(0, carried + read);
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
}
