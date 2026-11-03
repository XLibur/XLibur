using System.IO.Compression;

namespace XLibur.Operations.Workbooks;

/// <summary>
/// Refuses an .xlsx that would inflate past the limits before XLibur parses any of it. XLibur and the
/// Open XML SDK read whatever a package holds, so a small file can otherwise expand to gigabytes.
/// </summary>
/// <remarks>
/// The check reads only the sizes in the zip directory, so it costs nothing like a load. A file can
/// understate them, but System.IO.Compression, which XLibur and the Open XML SDK read the package
/// through, stops each entry at its declared size, so the declared sizes bound what any parser is
/// given. <c>ZipGuardTests</c> pins that behaviour.
/// </remarks>
internal static class ZipGuard
{
    private const long RatioFloorBytes = 1024 * 1024;

    internal static void Check(byte[] bytes, string displayPath, Limits limits)
    {
        if (bytes.Length >= 8 && BitConverter.ToUInt64(bytes, 0) == 0xE11AB1A1E011CFD0)
            throw OperationException.Unsupported("encrypted_or_legacy",
                $"'{displayPath}' is an OLE compound file: an encrypted workbook or a legacy .xls, neither of which this tool opens.");

        ZipArchive archive;
        try
        {
            archive = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);
        }
        catch (InvalidDataException e)
        {
            throw OperationException.BadInput("not_a_workbook", $"'{displayPath}' is not an .xlsx package: {e.Message}", e);
        }

        using (archive)
        {
            if (archive.Entries.Count > limits.MaxZipEntries)
                throw Bomb(displayPath, $"it holds {archive.Entries.Count:N0} entries; the limit is {limits.MaxZipEntries:N0}");

            long total = 0;
            foreach (var entry in archive.Entries)
            {
                total += entry.Length;
                if (total > limits.MaxUncompressedBytes)
                    throw Bomb(displayPath, $"its entries inflate past {limits.MaxUncompressedBytes:N0} bytes");

                var ratio = entry.Length / Math.Max(1, entry.CompressedLength);
                if (entry.Length > RatioFloorBytes && ratio > limits.MaxCompressionRatio)
                    throw Bomb(displayPath, $"entry '{entry.FullName}' inflates {ratio:N0} times; the limit is {limits.MaxCompressionRatio:N0}");
            }
        }
    }

    private static OperationException Bomb(string path, string why)
        => OperationException.BadInput("package_too_large", $"Refusing '{path}': {why}.");
}
