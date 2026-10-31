using System;
using System.IO;

namespace XLibur.Excel.Drawings;

/// <summary>
/// The image of a picture, which notices when it is written to, so a save feeds the image part only
/// when the image changed since the picture was loaded or last saved.
/// </summary>
/// <remarks>
/// A caller changes a picture's image by writing into <see cref="IXLPicture.ImageStream"/>, the way
/// <c>SetLength(0)</c> then <c>Write(bytes)</c> does. Every write and every change of length marks
/// the stream. Bytes written straight into the array <see cref="MemoryStream.GetBuffer"/> returns
/// are not seen, because that method cannot be overridden.
/// </remarks>
internal sealed class XLPictureImageStream : MemoryStream
{
    /// <summary>Whether the image was written to since <see cref="MarkUnmodified"/>.</summary>
    internal bool IsModified { get; private set; }

    /// <summary>The image as it is now is the one the file holds.</summary>
    internal void MarkUnmodified() => IsModified = false;

    public override void Write(byte[] buffer, int offset, int count)
    {
        IsModified = true;
        base.Write(buffer, offset, count);
    }

    public override void Write(ReadOnlySpan<byte> buffer)
    {
        IsModified = true;
        base.Write(buffer);
    }

    public override void WriteByte(byte value)
    {
        IsModified = true;
        base.WriteByte(value);
    }

    public override void SetLength(long value)
    {
        IsModified = true;
        base.SetLength(value);
    }
}
