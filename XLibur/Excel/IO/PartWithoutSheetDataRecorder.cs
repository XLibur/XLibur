using System;
using System.IO;
using System.Xml;

namespace XLibur.Excel.IO;

/// <summary>
/// Passes a worksheet part on to the package as it is written, and keeps a copy of it with
/// <c>&lt;sheetData&gt;</c> reduced to an empty element, in the form the load keeps
/// (<see cref="WorksheetPartBuffer.CopyWithoutSheetData"/>). The next save of the workbook compares
/// against that copy instead of against the part the load read (#719).
/// </summary>
/// <remarks>
/// The copy is taken from the bytes on their way to the part, so it is what the part holds, byte for
/// byte, apart from the cells. Tier 2 compares it with the part that way. The content of
/// <c>&lt;sheetData&gt;</c> is not copied: <see cref="SheetDataStarts"/> keeps only the start tag,
/// and <see cref="SheetDataEnded"/> resumes after the end tag.
/// </remarks>
internal sealed class PartWithoutSheetDataRecorder : Stream
{
    private readonly Stream _part;
    private readonly MemoryStream _copy = new();

    private Mode _mode = Mode.Copy;

    /// <summary>The quote of the attribute value the start tag is in, or 0 outside one.</summary>
    private byte _quote;

    /// <summary>The length of the copy up to the <c>&gt;</c> that closes the start tag, or -1.</summary>
    private long _startTagEnd = -1;

    private bool _failed;

    internal PartWithoutSheetDataRecorder(Stream part) => _part = part;

    private enum Mode
    {
        /// <summary>Bytes are copied.</summary>
        Copy,

        /// <summary>Bytes are copied up to the <c>&gt;</c> that closes the start tag.</summary>
        StartTag,

        /// <summary>Bytes are not copied.</summary>
        Skip,
    }

    /// <summary>
    /// Called before <c>&lt;sheetData&gt;</c> is written. Closes the root's start tag, which the
    /// writer holds open while attributes could follow, so the next byte is the one that opens
    /// <c>&lt;sheetData&gt;</c>.
    /// </summary>
    internal void SheetDataStarts(XmlWriter xml)
    {
        xml.WriteRaw(string.Empty);
        xml.Flush();

        _failed |= _mode != Mode.Copy || _startTagEnd >= 0;
        _mode = Mode.StartTag;
    }

    /// <summary>Called once <c>&lt;sheetData&gt;</c> has been written, end tag and all.</summary>
    internal void SheetDataEnded(XmlWriter xml)
    {
        xml.Flush();

        _failed |= _mode != Mode.Skip;
        _mode = Mode.Copy;
    }

    /// <summary>
    /// The part with <c>&lt;sheetData&gt;</c> reduced to an empty element, or null when the part was
    /// not written as <see cref="SheetDataStarts"/> and <see cref="SheetDataEnded"/> expect.
    /// </summary>
    internal byte[]? ToPartWithoutSheetData()
    {
        if (_failed || _mode != Mode.Copy || _startTagEnd < 2)
            return null;

        var copy = _copy.GetBuffer().AsSpan(0, (int)_copy.Length);
        var startTagEnd = (int)_startTagEnd;

        // Already an empty element: the copy is the whole part.
        if (copy[startTagEnd - 2] == (byte)'/')
            return copy.ToArray();

        // Everything before the start tag's '>', then '/>' in its place.
        var markup = new byte[copy.Length + 1];
        copy[..(startTagEnd - 1)].CopyTo(markup);
        markup[startTagEnd - 1] = (byte)'/';
        markup[startTagEnd] = (byte)'>';
        copy[startTagEnd..].CopyTo(markup.AsSpan(startTagEnd + 1));
        return markup;
    }

    public override void Write(ReadOnlySpan<byte> buffer)
    {
        _part.Write(buffer);

        switch (_mode)
        {
            case Mode.Copy:
                _copy.Write(buffer);
                break;
            case Mode.StartTag:
                CopyStartTag(buffer);
                break;
        }
    }

    public override void Write(byte[] buffer, int offset, int count) => Write(buffer.AsSpan(offset, count));

    public override void WriteByte(byte value) => Write([value]);

    /// <summary>
    /// Copies up to the <c>&gt;</c> that closes the start tag, skipping any inside a quoted attribute
    /// value, where the character is legal. The tag can arrive across several writes.
    /// </summary>
    private void CopyStartTag(ReadOnlySpan<byte> buffer)
    {
        for (var i = 0; i < buffer.Length; i++)
        {
            var b = buffer[i];
            if (_quote != 0)
            {
                if (b == _quote)
                    _quote = 0;
            }
            else if (b is (byte)'"' or (byte)'\'')
            {
                _quote = b;
            }
            else if (b == (byte)'>')
            {
                _copy.Write(buffer[..(i + 1)]);
                _startTagEnd = _copy.Length;
                _mode = Mode.Skip;
                return;
            }
        }

        _copy.Write(buffer);
    }

    /// <summary>
    /// Does not flush the part. The writer flushes at the edges of <c>&lt;sheetData&gt;</c> to hand
    /// on what it buffered, and a flush of the part's compressor would end a deflate block there.
    /// The part is flushed when it is closed.
    /// </summary>
    public override void Flush()
    {
    }

    public override bool CanRead => false;

    public override bool CanSeek => false;

    public override bool CanWrite => true;

    public override long Length => throw new NotSupportedException();

    public override long Position
    {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

    public override void SetLength(long value) => throw new NotSupportedException();

    protected override void Dispose(bool disposing)
    {
        if (disposing)
            _copy.Dispose();

        base.Dispose(disposing);
    }
}
