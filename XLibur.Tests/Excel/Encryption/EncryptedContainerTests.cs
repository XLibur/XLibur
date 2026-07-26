using System;
using System.IO;
using System.Threading.Tasks;
using OpenMcdf;
using XLibur.Excel.Exceptions;
using XLibur.Excel.IO.Encryption;

namespace XLibur.Tests.Excel.Encryption;

/// <summary>
/// The compound file wrapper and the choice of scheme made from the EncryptionInfo version.
/// </summary>
internal class EncryptedContainerTests
{
    /// <summary>A compound file carrying the two streams an encrypted workbook needs.</summary>
    private static MemoryStream Container(byte[] encryptionInfo, byte[] encryptedPackage)
    {
        var ms = new MemoryStream();
        EncryptedPackageContainer.WriteStreams(ms, encryptionInfo, encryptedPackage);
        ms.Position = 0;
        return ms;
    }

    private static byte[] VersionedEncryptionInfo(ushort major, ushort minor)
    {
        var info = new byte[64];
        BitConverter.TryWriteBytes(info.AsSpan(0), major);
        BitConverter.TryWriteBytes(info.AsSpan(2), minor);
        return info;
    }

    [Test]
    public async Task An_ordinary_package_is_not_mistaken_for_a_compound_file()
    {
        // An .xlsx is a zip and opens "PK". The two formats are told apart by their first bytes,
        // never by the file extension, which is the same for both.
        using var zipLike = new MemoryStream([0x50, 0x4B, 0x03, 0x04, 0, 0, 0, 0, 0, 0]);

        await Assert.That(EncryptedPackageContainer.IsCompoundFile(zipLike)).IsFalse();
    }

    [Test]
    public async Task Detection_leaves_the_stream_where_it_found_it()
    {
        // The caller reads the same stream afterwards whichever answer comes back, so a sniff that
        // consumed bytes would corrupt the ordinary path rather than the encrypted one.
        using var stream = new MemoryStream([0x50, 0x4B, 0x03, 0x04, 0, 0, 0, 0, 0, 0]);
        stream.Position = 2;

        _ = EncryptedPackageContainer.IsCompoundFile(stream);

        await Assert.That(stream.Position).IsEqualTo(2);
    }

    [Test]
    public async Task A_stream_too_short_to_hold_a_signature_is_not_a_compound_file()
    {
        using var tiny = new MemoryStream([0xD0, 0xCF]);

        await Assert.That(EncryptedPackageContainer.IsCompoundFile(tiny)).IsFalse();
    }

    [Test]
    public async Task Detection_requires_a_seekable_stream()
    {
        // Sniffing means reading and rewinding. A forward-only stream cannot be rewound, so this is
        // refused with an explanation rather than silently consuming the first eight bytes.
        using var forwardOnly = new ForwardOnlyStream([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);

        await Assert.That(() => EncryptedPackageContainer.IsCompoundFile(forwardOnly))
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task Bytes_that_start_like_a_compound_file_but_are_not_one_are_reported_clearly()
    {
        var garbage = new byte[512];
        garbage[0] = 0xD0;
        garbage[1] = 0xCF;
        garbage[2] = 0x11;
        garbage[3] = 0xE0;
        garbage[4] = 0xA1;
        garbage[5] = 0xB1;
        garbage[6] = 0x1A;
        garbage[7] = 0xE1;

        using var stream = new MemoryStream(garbage);

        await Assert.That(() => EncryptedPackageContainer.ReadStreams(stream))
            .Throws<XLEncryptionException>();
    }

    [Test]
    public async Task A_compound_file_missing_the_encrypted_package_is_reported_clearly()
    {
        using var ms = new MemoryStream();
        using (var storage = RootStorage.Create(ms, OpenMcdf.Version.V4, StorageModeFlags.LeaveOpen))
        {
            using var info = storage.CreateStream("EncryptionInfo");
            info.Write([1, 2, 3, 4], 0, 4);
        }

        ms.Position = 0;

        var exception = await Assert.That(() => EncryptedPackageContainer.ReadStreams(ms))
            .Throws<XLEncryptionException>();

        await Assert.That(exception!.Message).Contains("EncryptedPackage");
    }

    [Test]
    public async Task The_two_streams_survive_a_write_and_read_of_the_container()
    {
        var encryptionInfo = new byte[100];
        var encryptedPackage = new byte[10_000];
        Random.Shared.NextBytes(encryptionInfo);
        Random.Shared.NextBytes(encryptedPackage);

        // 10,000 bytes is deliberately over the 4096 byte cutoff below which a compound file keeps
        // content in the mini stream, so this covers the ordinary sector path as well.
        using var container = Container(encryptionInfo, encryptedPackage);
        var (readInfo, readPackage) = EncryptedPackageContainer.ReadStreams(container);

        await Assert.That(readInfo).IsEquivalentTo(encryptionInfo);
        await Assert.That(readPackage).IsEquivalentTo(encryptedPackage);
    }

    [Test]
    public async Task A_tiny_package_below_the_mini_stream_cutoff_also_survives()
    {
        var encryptionInfo = new byte[16];
        var encryptedPackage = new byte[64];
        Random.Shared.NextBytes(encryptionInfo);
        Random.Shared.NextBytes(encryptedPackage);

        using var container = Container(encryptionInfo, encryptedPackage);
        var (readInfo, readPackage) = EncryptedPackageContainer.ReadStreams(container);

        await Assert.That(readInfo).IsEquivalentTo(encryptionInfo);
        await Assert.That(readPackage).IsEquivalentTo(encryptedPackage);
    }

    [Test]
    public async Task An_rc4_encrypted_workbook_says_what_it_is_and_what_to_do()
    {
        using var container = Container(VersionedEncryptionInfo(2, 2), new byte[64]);

        var exception = await Assert.That(() => WorkbookEncryption.Decrypt(container, "password"))
            .Throws<XLEncryptionException>();

        await Assert.That(exception!.Message).Contains("RC4");
    }

    [Test]
    [Arguments((ushort)5, (ushort)5)]
    [Arguments((ushort)1, (ushort)1)]
    [Arguments((ushort)4, (ushort)3)]
    public async Task An_unknown_encryption_version_is_named_in_the_error(ushort major, ushort minor)
    {
        using var container = Container(VersionedEncryptionInfo(major, minor), new byte[64]);

        var exception = await Assert.That(() => WorkbookEncryption.Decrypt(container, "password"))
            .Throws<XLEncryptionException>();

        await Assert.That(exception!.Message).Contains($"{major}.{minor}");
    }

    [Test]
    public async Task An_encryption_info_too_short_to_hold_a_version_is_rejected()
    {
        using var container = Container(new byte[2], new byte[64]);

        await Assert.That(() => WorkbookEncryption.Decrypt(container, "password"))
            .Throws<XLEncryptionException>();
    }

    [Test]
    public async Task Decrypting_without_a_password_says_where_to_put_one()
    {
        using var container = Container(VersionedEncryptionInfo(4, 4), new byte[64]);

        var exception = await Assert.That(() => WorkbookEncryption.Decrypt(container, null))
            .Throws<XLInvalidPasswordException>();

        await Assert.That(exception!.Message).Contains("LoadOptions.Password");
    }

    [Test]
    public async Task Encrypting_without_a_password_is_a_programming_error_not_a_file_error()
    {
        // Nothing about a file is wrong here: the caller asked to encrypt and supplied nothing to
        // encrypt with, which is an argument problem rather than an encryption one.
        using var destination = new MemoryStream();

        await Assert.That(() => WorkbookEncryption.Encrypt(destination, [1, 2, 3], string.Empty))
            .Throws<ArgumentException>();
    }

    /// <summary>A stream that cannot seek, for exercising the forward-only path.</summary>
    private sealed class ForwardOnlyStream(byte[] data) : Stream
    {
        private readonly MemoryStream _inner = new(data);

        public override bool CanRead => true;

        public override bool CanSeek => false;

        public override bool CanWrite => false;

        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush() => _inner.Flush();

        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();

            base.Dispose(disposing);
        }
    }
}
