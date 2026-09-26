using System;
using System.Buffers;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace XLibur.Excel.IO;

/// <summary>
/// A worksheet part inflated into memory once and split around its <c>&lt;sheetData&gt;</c>, so
/// the loader reads the cells and the rest of the markup separately without inflating or
/// tokenising the part more than once.
/// </summary>
/// <remarks>
/// Streaming the part costs three full passes over the cells on an open-and-save round trip: the
/// structural pass tokenises them to skip them, the cell pass tokenises them to read them, and the
/// save tokenises them a third time to copy the rest of the part without them. On a small
/// workbook those skips are a large share of the whole round trip. Here the cells are tokenised
/// once, and the markup around them is cut out by <see cref="SheetDataLocator"/> and kept, which
/// serves both the structural pass and the first save.
/// <para>
/// Only parts up to <see cref="MaxBufferedLength"/> are buffered. A larger part keeps the
/// streaming path, which never holds the whole part in memory.
/// </para>
/// </remarks>
internal sealed class WorksheetPartBuffer : IDisposable
{
    /// <summary>
    /// The largest inflated part that is buffered. Big enough for any sheet where the fixed cost of
    /// the extra passes shows, small enough that the transient buffer stays well below the memory
    /// the loaded cells take anyway.
    /// </summary>
    internal const int MaxBufferedLength = 16 * 1024 * 1024;

    private byte[]? _buffer;
    private readonly int _length;

    private WorksheetPartBuffer(byte[] buffer, int length, SheetDataLocator.Location sheetData)
    {
        _buffer = buffer;
        _length = length;
        SheetData = sheetData;
    }

    internal SheetDataLocator.Location SheetData { get; }

    /// <summary>
    /// Reads and splits the part, or returns null when it is too large, when its stream does not
    /// match the length the package declares, or when <see cref="SheetDataLocator"/> cannot place
    /// the cells exactly.
    /// </summary>
    internal static WorksheetPartBuffer? TryRead(WorksheetPart worksheetPart)
    {
        using var stream = worksheetPart.GetStream(FileMode.Open, FileAccess.Read);

        long declaredLength;
        try
        {
            declaredLength = stream.Length;
        }
        catch (NotSupportedException)
        {
            return null;
        }

        if (declaredLength <= 0 || declaredLength > MaxBufferedLength)
            return null;

        // One byte of headroom, so a stream that runs past its declared length is caught rather
        // than silently truncated. The declared length comes from the zip directory and is not
        // trusted beyond sizing the buffer.
        var buffer = ArrayPool<byte>.Shared.Rent((int)declaredLength + 1);
        var length = stream.ReadAtLeast(buffer, buffer.Length, throwOnEndOfStream: false);

        if (length != declaredLength
            || !SheetDataLocator.TryLocate(buffer.AsSpan(0, length), out var sheetData))
        {
            ArrayPool<byte>.Shared.Return(buffer);
            return null;
        }

        return new WorksheetPartBuffer(buffer, length, sheetData);
    }

    /// <summary>A read-only stream over the whole part.</summary>
    internal MemoryStream OpenRead() => new(Buffer, 0, _length, writable: false);

    /// <summary>
    /// The part with <c>&lt;sheetData&gt;</c> reduced to an empty element and every other byte as
    /// it was.
    /// </summary>
    internal byte[] CopyWithoutSheetData()
    {
        var xml = Buffer.AsSpan(0, _length);
        var name = xml.Slice(SheetData.NameStart, SheetData.NameLength);
        var suffixLength = _length - SheetData.End;

        var copy = new byte[SheetData.Start + name.Length + 3 + suffixLength];
        var span = copy.AsSpan();

        xml[..SheetData.Start].CopyTo(span);
        span = span[SheetData.Start..];
        span[0] = (byte)'<';
        name.CopyTo(span[1..]);
        span[name.Length + 1] = (byte)'/';
        span[name.Length + 2] = (byte)'>';
        xml[SheetData.End..].CopyTo(span[(name.Length + 3)..]);

        return copy;
    }

    public void Dispose()
    {
        if (_buffer is null)
            return;

        ArrayPool<byte>.Shared.Return(_buffer);
        _buffer = null;
    }

    private byte[] Buffer => _buffer ?? throw new ObjectDisposedException(nameof(WorksheetPartBuffer));
}
