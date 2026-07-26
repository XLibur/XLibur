using System;
using System.IO;
using OpenMcdf;
using XLibur.Excel.Exceptions;

namespace XLibur.Excel.IO.Encryption;

/// <summary>
/// The OLE compound file that an encrypted workbook is delivered in. The container holds two
/// streams: <c>EncryptionInfo</c>, describing how the content was encrypted, and
/// <c>EncryptedPackage</c>, the .xlsx itself.
/// </summary>
/// <remarks>
/// [MS-CFB]. The compound file format is handled by OpenMcdf rather than in house; this type is the
/// seam that keeps that dependency to one file.
/// </remarks>
internal static class EncryptedPackageContainer
{
    internal const string EncryptionInfoStreamName = "EncryptionInfo";

    internal const string EncryptedPackageStreamName = "EncryptedPackage";

    /// <summary>The compound file signature, [MS-CFB] 2.2.</summary>
    private static ReadOnlySpan<byte> CompoundFileSignature => [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

    /// <summary>
    /// Whether the stream starts with the compound file signature. An .xlsx is a zip and begins
    /// "PK", so the two are told apart by their first bytes and never by their extension.
    /// </summary>
    /// <remarks>
    /// Leaves the position where it found it, because the caller goes on to read the same stream
    /// whichever way the answer comes out.
    /// </remarks>
    public static bool IsCompoundFile(Stream stream)
    {
        if (!stream.CanSeek)
        {
            throw new ArgumentException(
                "Detecting an encrypted workbook requires a seekable stream.", nameof(stream));
        }

        if (stream.Length - stream.Position < CompoundFileSignature.Length)
            return false;

        var origin = stream.Position;
        try
        {
            Span<byte> header = stackalloc byte[8];
            var read = 0;
            while (read < header.Length)
            {
                var count = stream.Read(header[read..]);
                if (count == 0)
                    return false;

                read += count;
            }

            return header.SequenceEqual(CompoundFileSignature);
        }
        finally
        {
            stream.Position = origin;
        }
    }

    /// <summary>
    /// Reads the two streams that make up an encrypted workbook.
    /// </summary>
    public static (byte[] EncryptionInfo, byte[] EncryptedPackage) ReadStreams(Stream stream)
    {
        RootStorage storage;
        try
        {
            storage = RootStorage.Open(stream, StorageModeFlags.LeaveOpen);
        }
        catch (Exception e) when (e is OpenMcdf.FileFormatException or IOException)
        {
            throw new XLEncryptionException("The file is not a readable compound file.", e);
        }

        using (storage)
        {
            if (!storage.ContainsEntry(EncryptionInfoStreamName) || !storage.ContainsEntry(EncryptedPackageStreamName))
            {
                // A compound file without these two streams is something else entirely, most often a
                // pre-2007 .xls, which is a different format rather than an encrypted one.
                throw new XLEncryptionException(
                    "The file is a compound file but not an encrypted workbook: it has no EncryptionInfo and EncryptedPackage streams. " +
                    "Legacy .xls workbooks are not supported.");
            }

            return (ReadAll(storage, EncryptionInfoStreamName), ReadAll(storage, EncryptedPackageStreamName));
        }
    }

    /// <summary>
    /// Writes the two streams into a new compound file.
    /// </summary>
    public static void WriteStreams(Stream destination, byte[] encryptionInfo, byte[] encryptedPackage)
    {
        // V4 uses 4096 byte sectors, which is what Excel writes for encrypted workbooks and what
        // keeps a package of any real size out of the mini stream.
        using var storage = RootStorage.Create(destination, OpenMcdf.Version.V4, StorageModeFlags.LeaveOpen);

        using (var infoStream = storage.CreateStream(EncryptionInfoStreamName))
            infoStream.Write(encryptionInfo, 0, encryptionInfo.Length);

        using (var packageStream = storage.CreateStream(EncryptedPackageStreamName))
            packageStream.Write(encryptedPackage, 0, encryptedPackage.Length);

        // No Commit here: that is for transacted storage only and throws otherwise. A
        // non-transacted storage writes through, and disposing it flushes the header and FAT.
    }

    private static byte[] ReadAll(Storage storage, string name)
    {
        using var stream = storage.OpenStream(name);
        var buffer = new byte[stream.Length];

        var read = 0;
        while (read < buffer.Length)
        {
            var count = stream.Read(buffer, read, buffer.Length - read);
            if (count == 0)
                throw new XLEncryptionException($"The '{name}' stream ended before its declared length.");

            read += count;
        }

        return buffer;
    }
}
