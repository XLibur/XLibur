using System.Buffers.Binary;
using System.IO.Compression;
using XLibur.Operations.Workbooks;

namespace XLibur.Operations.Tests;

public class ZipGuardTests
{
    private static readonly Limits Small = Limits.Default with { MaxUncompressedBytes = 4 * 1024 * 1024, MaxZipEntries = 50 };

    [Test]
    public async Task A_package_that_inflates_past_the_limit_is_refused()
    {
        var bomb = Zip(("xl/worksheets/sheet1.xml", 8 * 1024 * 1024));

        await Expect.Refused(() => ZipGuard.Check(bomb, "bomb.xlsx", Small), ErrorKind.BadInput, "package_too_large");
    }

    [Test]
    public async Task An_entry_that_understates_its_size_inflates_no_further_than_it_declares()
    {
        var bomb = Zip(("xl/worksheets/sheet1.xml", 8 * 1024 * 1024));
        Understate(bomb);

        long inflated = 0;
        Exception? failure = null;
        try
        {
            using var archive = new ZipArchive(new MemoryStream(bomb), ZipArchiveMode.Read);
            using var stream = archive.Entries[0].Open();
            var buffer = new byte[81920];
            int n;
            while ((n = stream.Read(buffer)) > 0)
                inflated += n;
        }
        catch (InvalidDataException e)
        {
            failure = e;
        }

        await Assert.That(failure is not null || inflated <= 1).IsTrue();
        ZipGuard.Check(bomb, "liar.xlsx", Small);
    }

    [Test]
    public async Task An_entry_that_compresses_too_well_is_refused()
    {
        var bomb = Zip(("xl/sharedStrings.xml", 3 * 1024 * 1024));

        await Expect.Refused(() => ZipGuard.Check(bomb, "ratio.xlsx", Small with { MaxCompressionRatio = 100 }),
            ErrorKind.BadInput, "package_too_large");
    }

    [Test]
    public async Task Too_many_entries_are_refused()
    {
        var many = Zip(Enumerable.Range(0, 60).Select(i => ($"part{i}.xml", 10)).ToArray());

        await Expect.Refused(() => ZipGuard.Check(many, "many.xlsx", Small), ErrorKind.BadInput, "package_too_large");
    }

    [Test]
    public async Task An_encrypted_or_legacy_file_is_unsupported_and_a_non_zip_is_bad_input()
    {
        var compoundFile = new byte[512];
        BinaryPrimitives.WriteUInt64LittleEndian(compoundFile, 0xE11AB1A1E011CFD0);

        await Expect.Refused(() => ZipGuard.Check(compoundFile, "secret.xlsx", Small), ErrorKind.Unsupported, "encrypted_or_legacy");
        await Expect.Refused(() => ZipGuard.Check("plain text"u8.ToArray(), "x.xlsx", Small), ErrorKind.BadInput, "not_a_workbook");
    }

    [Test]
    public async Task An_ordinary_package_passes()
    {
        var ordinary = Zip(("xl/workbook.xml", 4096), ("xl/worksheets/sheet1.xml", 100_000));

        ZipGuard.Check(ordinary, "ok.xlsx", Small);

        await Assert.That(ordinary.Length).IsGreaterThan(0);
    }

    [Test]
    public async Task A_zip_the_guard_passes_but_that_holds_no_workbook_is_unreadable_input()
    {
        var notAWorkbook = Zip(("readme.txt", 10));

        await Expect.Refused(() => WorkbookLoader.Load(notAWorkbook, "readme.xlsx", Small, recalculate: false),
            ErrorKind.BadInput, "unreadable_workbook");
    }

    private static byte[] Zip(params (string Name, int Bytes)[] entries)
    {
        using var buffer = new MemoryStream();
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (name, size) in entries)
            {
                using var stream = archive.CreateEntry(name, CompressionLevel.SmallestSize).Open();
                stream.Write(new byte[size]);
            }
        }

        return buffer.ToArray();
    }

    /// <summary>Rewrites every uncompressed size in the local headers and the central directory to 1 byte.</summary>
    private static void Understate(byte[] zip)
    {
        for (var i = 0; i + 30 <= zip.Length; i++)
        {
            var signature = BinaryPrimitives.ReadUInt32LittleEndian(zip.AsSpan(i));
            if (signature == 0x04034b50)
                BinaryPrimitives.WriteUInt32LittleEndian(zip.AsSpan(i + 22), 1);
            else if (signature == 0x02014b50)
                BinaryPrimitives.WriteUInt32LittleEndian(zip.AsSpan(i + 24), 1);
        }
    }
}
