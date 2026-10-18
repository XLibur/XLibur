using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Tests.Utils;
using SaveOptions = XLibur.Excel.SaveOptions;

namespace XLibur.Tests.Excel.Saving;

/// <summary>
/// <see cref="SaveOptions.CompressionLevel"/> applies to every part a save of a loaded workbook
/// writes, on every save path, and leaves the parts the save carries over as the file had them (#688).
/// </summary>
public class SaveCompressionLevelTests
{
    private const string Data = "xl/worksheets/sheet1.xml";

    private const string Other = "xl/worksheets/sheet2.xml";

    private const string Styles = "xl/styles.xml";

    private const string Workbook = "xl/workbook.xml";

    private const string Theme = "xl/theme/theme1.xml";

    private static readonly CompressionLevel[] Levels =
    [
        CompressionLevel.NoCompression,
        CompressionLevel.Fastest,
        CompressionLevel.Optimal,
        CompressionLevel.SmallestSize
    ];

    private static readonly Lazy<byte[]> Source = new(BuildSource);

    public enum SavePath
    {
        SaveAsStream,
        SaveAsFile,
        SaveInPlaceFromFile,
        SaveInPlaceFromStream,
        SecondSave
    }

    [Test]
    [Arguments(SavePath.SaveAsStream)]
    [Arguments(SavePath.SaveAsFile)]
    [Arguments(SavePath.SaveInPlaceFromFile)]
    [Arguments(SavePath.SaveInPlaceFromStream)]
    [Arguments(SavePath.SecondSave)]
    public async Task Parts_the_save_writes_follow_the_level(SavePath path)
    {
        // Each part is compared with its own content deflated at the level, not with the other
        // levels: deflate does not always order them. On .NET 10, SmallestSize makes this sheet
        // 31,158 bytes and Optimal 30,579.
        foreach (var level in Levels)
        {
            var saved = Save(path, level, wb => wb.Worksheet("Data").Cell(1, 1).Value = 0.5);
            using var zip = new ZipArchive(new MemoryStream(saved), ZipArchiveMode.Read);

            foreach (var part in new[] { Data, Styles, Workbook })
            {
                var entry = zip.GetEntry(part)!;
                var expected = level == CompressionLevel.NoCompression
                    ? entry.Length
                    : DeflatedLength(entry, level);

                await Assert.That(entry.CompressedLength).IsEqualTo(expected).Because($"{part} at {level}");
            }
        }
    }

    [Test]
    [Arguments(CompressionLevel.Optimal)]
    [Arguments(CompressionLevel.SmallestSize)]
    public async Task Package_relationships_and_core_properties_are_stored_as_a_new_package_stores_them(
        CompressionLevel level)
    {
        // System.IO.Packaging creates both parts uncompressed in every package. Deflating them at
        // Optimal cost a round trip 8% of its time, for parts of a few hundred bytes.
        var saved = Save(SavePath.SaveAsStream, level, wb => wb.Properties.Title = "Edited");
        using var zip = new ZipArchive(new MemoryStream(saved), ZipArchiveMode.Read);

        var relationships = zip.GetEntry("_rels/.rels")!;
        var coreProperties = zip.Entries.Single(e => e.FullName.EndsWith(".psmdcp", StringComparison.Ordinal));

        await Assert.That(relationships.CompressedLength).IsEqualTo(relationships.Length);
        await Assert.That(coreProperties.CompressedLength).IsEqualTo(coreProperties.Length);
    }

    [Test]
    [Arguments(SavePath.SaveAsStream)]
    [Arguments(SavePath.SaveAsFile)]
    [Arguments(SavePath.SaveInPlaceFromFile)]
    [Arguments(SavePath.SaveInPlaceFromStream)]
    [Arguments(SavePath.SecondSave)]
    public async Task Parts_the_save_carries_over_keep_their_compression(SavePath path)
    {
        // The Data sheet is unchanged, so the save leaves its part as the file had it (#702), and
        // nothing in a save of a loaded workbook writes the theme.
        var source = Entries(Source.Value);

        foreach (var level in new[] { CompressionLevel.NoCompression, CompressionLevel.SmallestSize })
        {
            var saved = Entries(Save(path, level, wb => wb.Worksheet("Other").Cell("A2").Value = "Edited"));

            foreach (var part in new[] { Data, Theme })
                await Assert.That(saved[part]).IsEqualTo(source[part]).Because($"{part} at {level}");

            await Assert.That(saved[Other].Compressed == saved[Other].Length)
                .IsEqualTo(level == CompressionLevel.NoCompression).Because($"{Other} at {level}");
        }
    }

    private static byte[] Save(SavePath path, CompressionLevel level, Action<XLWorkbook> edit)
    {
        var options = new SaveOptions { CompressionLevel = level };
        switch (path)
        {
            case SavePath.SaveAsStream:
                {
                    using var source = new MemoryStream(Source.Value);
                    using var wb = new XLWorkbook(source);
                    edit(wb);
                    using var output = new MemoryStream();
                    wb.SaveAs(output, options);
                    return output.ToArray();
                }
            case SavePath.SaveAsFile:
                {
                    using var file = new TemporaryFile();
                    using (var source = new MemoryStream(Source.Value))
                    using (var wb = new XLWorkbook(source))
                    {
                        edit(wb);
                        wb.SaveAs(file.Path, options);
                    }

                    return File.ReadAllBytes(file.Path);
                }
            case SavePath.SaveInPlaceFromFile:
                {
                    using var file = new TemporaryFile();
                    File.WriteAllBytes(file.Path, Source.Value);
                    using (var wb = new XLWorkbook(file.Path))
                    {
                        edit(wb);
                        wb.Save(options);
                    }

                    return File.ReadAllBytes(file.Path);
                }
            case SavePath.SaveInPlaceFromStream:
                {
                    using var stream = new MemoryStream();
                    stream.Write(Source.Value);
                    stream.Position = 0;
                    using (var wb = new XLWorkbook(stream))
                    {
                        edit(wb);
                        wb.Save(options);
                    }

                    return stream.ToArray();
                }
            case SavePath.SecondSave:
                {
                    using var source = new MemoryStream(Source.Value);
                    using var wb = new XLWorkbook(source);
                    using var first = new MemoryStream();
                    wb.SaveAs(first);
                    edit(wb);
                    using var second = new MemoryStream();
                    wb.SaveAs(second, options);
                    return second.ToArray();
                }
            default:
                throw new ArgumentOutOfRangeException(nameof(path), path, null);
        }
    }

    private static long DeflatedLength(ZipArchiveEntry entry, CompressionLevel level)
    {
        using var content = new MemoryStream();
        using (var stream = entry.Open())
            stream.CopyTo(content);

        using var deflated = new MemoryStream();
        using (var deflate = new DeflateStream(deflated, level, leaveOpen: true))
            deflate.Write(content.GetBuffer(), 0, (int)content.Length);

        return deflated.Length;
    }

    private static Dictionary<string, (long Length, long Compressed, uint Crc32)> Entries(byte[] package)
    {
        using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        return zip.Entries.ToDictionary(e => e.FullName, e => (e.Length, e.CompressedLength, e.Crc32));
    }

    /// <summary>A sheet of 1,000 rows of eight numbers, and a small second sheet.</summary>
    private static byte[] BuildSource()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet("Data");
        for (var row = 1; row <= 1_000; row++)
        {
            for (var column = 1; column <= 8; column++)
                data.Cell(row, column).Value = row * 8 + column;
        }

        wb.AddWorksheet("Other").Cell("A1").Value = "Other";

        using var output = new MemoryStream();
        wb.SaveAs(output);
        return output.ToArray();
    }
}
