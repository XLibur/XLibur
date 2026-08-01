using System;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace XLibur.Excel.Streaming;

/// <summary>
/// Writes the OPC package of a streaming workbook straight into a zip, one part at a time.
/// </summary>
/// <remarks>
/// <para>
/// The OpenXML SDK's packaging layer cannot be used here. <c>System.IO.Packaging</c> opens a
/// package for read/write, which maps to <see cref="ZipArchiveMode.Update"/>, and that mode
/// holds every entry's <em>uncompressed</em> bytes in memory until the archive is closed - about
/// half a gigabyte for a million-row sheet, which is precisely what this API exists to avoid.
/// Opening the package write-only does give <see cref="ZipArchiveMode.Create"/>, but the SDK
/// enumerates parts as it works and fails on a write-only container.
/// </para>
/// <para>
/// So the package is assembled by hand over <see cref="ZipArchiveMode.Create"/>, which writes
/// each entry through to the output as it is produced and retains only the central directory.
/// The parts involved are few and simple - content types, two relationship parts, the workbook,
/// the worksheets, the styles and the shared strings - and the shape of an .xlsx package is
/// fixed. A side benefit is that the output stream need not be seekable, so a workbook can be
/// written directly to a network stream.
/// </para>
/// <para>
/// One consequence of never seeking: <c>[Content_Types].xml</c> is written last, once the set of
/// worksheets is known, rather than as the first entry. Readers resolve parts through the zip
/// central directory, so its position in the archive does not matter.
/// </para>
/// </remarks>
internal sealed class StreamingPackageWriter : IDisposable
{
    internal const string ContentTypesNs = "http://schemas.openxmlformats.org/package/2006/content-types";
    internal const string PackageRelationshipsNs = "http://schemas.openxmlformats.org/package/2006/relationships";

    private readonly ZipArchive _archive;
    private readonly CompressionLevel _compressionLevel;
    private bool _disposed;

    internal StreamingPackageWriter(Stream output, bool leaveOpen, CompressionLevel compressionLevel)
    {
        _archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen);
        _compressionLevel = compressionLevel;
    }

    /// <summary>
    /// Start a part. The returned writer owns the entry: disposing it closes both the writer and
    /// the entry, and no other part may be started until it has been.
    /// </summary>
    internal XmlWriter CreatePart(string entryName)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        var entry = _archive.CreateEntry(entryName, _compressionLevel);
        var settings = new XmlWriterSettings
        {
            CloseOutput = true,
            Encoding = XLHelper.NoBomUTF8
        };

        return XmlWriter.Create(entry.Open(), settings);
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;
        _archive.Dispose();
    }
}
