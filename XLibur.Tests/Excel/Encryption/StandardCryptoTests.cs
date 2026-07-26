using System;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using XLibur.Excel.Exceptions;
using XLibur.Excel.IO.Encryption;
using XLibur.Excel.IO.Encryption.Standard;

namespace XLibur.Tests.Excel.Encryption;

/// <summary>
/// Standard encryption, the Office 2007 scheme. XLibur reads it and never writes it, so no round
/// trip through the workbook API reaches this code and no file in the corpus covers it either.
/// </summary>
/// <remarks>
/// <para>
/// The descriptors here are built by hand from [MS-OFFCRYPTO] 2.3.4.5 through 2.3.4.9, which is
/// worth being clear about: the malformed-descriptor tests are worth exactly what they look like,
/// but the round trip is a second implementation by the same author and so shows internal
/// consistency and freedom from regression rather than conformance to the specification. Only a
/// workbook encrypted by Office 2007 can show that, and the corpus README asks for one.
/// </para>
/// </remarks>
internal class StandardCryptoTests
{
    private const string Password = "a password";

    private const uint AlgIdAes128 = 0x660E;
    private const uint AlgIdAes256 = 0x6610;
    private const uint AlgIdHashSha1 = 0x8004;

    /// <summary>
    /// Derives the key the way [MS-OFFCRYPTO] 2.3.4.7 describes: a spin of SHA-1 over the password,
    /// then an ipad/opad pair that resembles HMAC without being it.
    /// </summary>
    private static byte[] DeriveKey(string password, byte[] salt, int keyBytes)
    {
        var passwordBytes = Encoding.Unicode.GetBytes(password);
        var seed = new byte[salt.Length + passwordBytes.Length];
        salt.CopyTo(seed, 0);
        passwordBytes.CopyTo(seed, salt.Length);

        var current = SHA1.HashData(seed);
        for (var i = 0; i < 50000; i++)
        {
            var buffer = new byte[4 + current.Length];
            BitConverter.TryWriteBytes(buffer.AsSpan(0, 4), i);
            current.CopyTo(buffer, 4);
            current = SHA1.HashData(buffer);
        }

        var block = new byte[current.Length + 4];
        current.CopyTo(block, 0);
        var hFinal = SHA1.HashData(block);

        var x1 = SHA1.HashData(XorWithPad(hFinal, 0x36));
        var x2 = SHA1.HashData(XorWithPad(hFinal, 0x5C));

        var x3 = new byte[x1.Length + x2.Length];
        x1.CopyTo(x3, 0);
        x2.CopyTo(x3, x1.Length);

        return x3.AsSpan(0, keyBytes).ToArray();
    }

    private static byte[] XorWithPad(byte[] hash, byte pad)
    {
        var buffer = new byte[64];
        buffer.AsSpan().Fill(pad);
        for (var i = 0; i < hash.Length; i++)
            buffer[i] ^= hash[i];

        return buffer;
    }

    private static byte[] Ecb(byte[] key, byte[] data, bool encrypting)
    {
        using var aes = Aes.Create();
        aes.KeySize = key.Length * 8;
        aes.Key = key;
        aes.Mode = CipherMode.ECB;
        aes.Padding = PaddingMode.None;

        using var transform = encrypting ? aes.CreateEncryptor() : aes.CreateDecryptor();
        return transform.TransformFinalBlock(data, 0, data.Length);
    }

    private static byte[] Pad(byte[] value, int blockSize)
    {
        if (value.Length % blockSize == 0)
            return value;

        var padded = new byte[value.Length + (blockSize - value.Length % blockSize)];
        value.CopyTo(padded, 0);
        return padded;
    }

    /// <summary>
    /// Builds an EncryptionInfo stream. The pieces are separately overridable so a test can make one
    /// of them invalid without hand-assembling the whole thing.
    /// </summary>
    private static byte[] EncryptionInfo(
        byte[] salt,
        byte[] encryptedVerifier,
        byte[] encryptedVerifierHash,
        uint algId = AlgIdAes256,
        uint algIdHash = AlgIdHashSha1,
        int keyBits = 256,
        int? saltSizeOverride = null,
        int? headerSizeOverride = null,
        ushort majorVersion = 4)
    {
        var header = new byte[32];
        BitConverter.TryWriteBytes(header.AsSpan(8), algId);
        BitConverter.TryWriteBytes(header.AsSpan(12), algIdHash);
        BitConverter.TryWriteBytes(header.AsSpan(16), keyBits);

        var verifier = new byte[4 + salt.Length + encryptedVerifier.Length + 4 + encryptedVerifierHash.Length];
        var offset = 0;
        BitConverter.TryWriteBytes(verifier.AsSpan(offset), saltSizeOverride ?? salt.Length);
        offset += 4;
        salt.CopyTo(verifier.AsSpan(offset));
        offset += salt.Length;
        encryptedVerifier.CopyTo(verifier.AsSpan(offset));
        offset += encryptedVerifier.Length;
        BitConverter.TryWriteBytes(verifier.AsSpan(offset), 20);
        offset += 4;
        encryptedVerifierHash.CopyTo(verifier.AsSpan(offset));

        var stream = new byte[12 + header.Length + verifier.Length];
        BitConverter.TryWriteBytes(stream.AsSpan(0), majorVersion);
        BitConverter.TryWriteBytes(stream.AsSpan(2), (ushort)2);
        BitConverter.TryWriteBytes(stream.AsSpan(4), 0x24u);
        BitConverter.TryWriteBytes(stream.AsSpan(8), headerSizeOverride ?? header.Length);
        header.CopyTo(stream.AsSpan(12));
        verifier.CopyTo(stream.AsSpan(12 + header.Length));
        return stream;
    }

    /// <summary>Builds a valid descriptor and the matching encrypted package.</summary>
    private static (byte[] EncryptionInfo, byte[] EncryptedPackage) Encrypt(
        byte[] payload, string password, int keyBits = 256, ushort majorVersion = 4)
    {
        var salt = new byte[16];
        RandomNumberGenerator.Fill(salt);

        var key = DeriveKey(password, salt, keyBits / 8);

        var verifier = new byte[16];
        RandomNumberGenerator.Fill(verifier);

        var encryptedVerifier = Ecb(key, verifier, encrypting: true);
        var encryptedVerifierHash = Ecb(key, Pad(SHA1.HashData(verifier), 16), encrypting: true);

        var package = new byte[8 + Pad(payload, 16).Length];
        BitConverter.TryWriteBytes(package.AsSpan(0, 8), (long)payload.Length);
        Ecb(key, Pad(payload, 16), encrypting: true).CopyTo(package.AsSpan(8));

        return (
            EncryptionInfo(salt, encryptedVerifier, encryptedVerifierHash, keyBits: keyBits, majorVersion: majorVersion),
            package);
    }

    [Test]
    [Arguments(128)]
    [Arguments(256)]
    public async Task A_standard_encrypted_package_round_trips(int keyBits)
    {
        var payload = new byte[3000];
        RandomNumberGenerator.Fill(payload);

        var (encryptionInfo, encryptedPackage) = Encrypt(payload, Password, keyBits);

        var descriptor = StandardEncryptionDescriptor.Parse(encryptionInfo);
        var key = descriptor.DeriveAndVerifyKey(Password);
        var decrypted = StandardEncryptionDescriptor.DecryptPackage(key, encryptedPackage);

        await Assert.That(descriptor.KeyBytes).IsEqualTo(keyBits / 8);
        await Assert.That(decrypted).IsEquivalentTo(payload);
    }

    [Test]
    [Arguments((ushort)3)]
    [Arguments((ushort)4)]
    public async Task A_standard_container_decrypts_through_the_workbook_entry_point(ushort majorVersion)
    {
        // Both 3.2 and 4.2 mean standard encryption, and the entry point has to route either of them
        // away from the agile reader. Driving the descriptor directly, as the tests above do, would
        // never exercise that choice.
        var payload = new byte[1000];
        RandomNumberGenerator.Fill(payload);

        var (encryptionInfo, encryptedPackage) = Encrypt(payload, Password, majorVersion: majorVersion);

        using var container = new MemoryStream();
        EncryptedPackageContainer.WriteStreams(container, encryptionInfo, encryptedPackage);
        container.Position = 0;

        using var decrypted = WorkbookEncryption.Decrypt(container, Password);

        await Assert.That(decrypted.ToArray()).IsEquivalentTo(payload);
    }

    [Test]
    public async Task A_wrong_password_on_a_standard_container_reaches_the_caller_as_a_password_error()
    {
        var (encryptionInfo, encryptedPackage) = Encrypt([1, 2, 3, 4], Password);

        using var container = new MemoryStream();
        EncryptedPackageContainer.WriteStreams(container, encryptionInfo, encryptedPackage);
        container.Position = 0;

        await Assert.That(() => WorkbookEncryption.Decrypt(container, "the wrong password"))
            .Throws<XLInvalidPasswordException>();
    }

    [Test]
    public async Task A_wrong_password_is_refused_by_the_verifier()
    {
        var (encryptionInfo, _) = Encrypt([1, 2, 3, 4], Password);
        var descriptor = StandardEncryptionDescriptor.Parse(encryptionInfo);

        await Assert.That(() => descriptor.DeriveAndVerifyKey("the wrong password"))
            .Throws<XLInvalidPasswordException>();
    }

    [Test]
    public async Task The_length_prefix_trims_the_block_padding()
    {
        // 100 bytes pad up to 112. Without the prefix being honoured the caller would receive 12
        // trailing zero bytes as though they were content.
        var payload = new byte[100];
        RandomNumberGenerator.Fill(payload);

        var (encryptionInfo, encryptedPackage) = Encrypt(payload, Password);
        var descriptor = StandardEncryptionDescriptor.Parse(encryptionInfo);
        var key = descriptor.DeriveAndVerifyKey(Password);

        await Assert.That(StandardEncryptionDescriptor.DecryptPackage(key, encryptedPackage).Length).IsEqualTo(100);
    }

    [Test]
    public async Task An_unsupported_cipher_names_the_identifier()
    {
        // 0x6801 is RC4, which the standard scheme allows and XLibur does not read.
        var info = EncryptionInfo(new byte[16], new byte[16], new byte[32], algId: 0x6801);

        var exception = await Assert.That(() => StandardEncryptionDescriptor.Parse(info))
            .Throws<XLEncryptionException>();

        await Assert.That(exception!.Message).Contains("6801");
    }

    [Test]
    public async Task An_unsupported_hash_names_the_identifier()
    {
        var info = EncryptionInfo(new byte[16], new byte[16], new byte[32], algIdHash: 0x800C);

        var exception = await Assert.That(() => StandardEncryptionDescriptor.Parse(info))
            .Throws<XLEncryptionException>();

        await Assert.That(exception!.Message).Contains("800C");
    }

    [Test]
    public async Task A_zero_hash_identifier_is_read_as_the_provider_default()
    {
        // Zero means "whatever this provider uses", which for standard encryption is SHA-1. Treating
        // it as an unknown algorithm would reject files that are perfectly ordinary.
        var info = EncryptionInfo(new byte[16], new byte[16], new byte[32], algIdHash: 0);

        await Assert.That(() => StandardEncryptionDescriptor.Parse(info)).ThrowsNothing();
    }

    [Test]
    public async Task The_key_length_in_the_header_wins_over_the_one_implied_by_the_cipher()
    {
        // AlgID is allowed to be the generic AES identifier, so the header's own key size decides.
        var info = EncryptionInfo(new byte[16], new byte[16], new byte[32], algId: AlgIdAes128, keyBits: 256);
        var descriptor = StandardEncryptionDescriptor.Parse(info);

        await Assert.That(descriptor.KeyBytes).IsEqualTo(32);
    }

    [Test]
    public async Task An_impossible_key_length_is_rejected()
    {
        var info = EncryptionInfo(new byte[16], new byte[16], new byte[32], keyBits: 64);

        await Assert.That(() => StandardEncryptionDescriptor.Parse(info)).Throws<XLEncryptionException>();
    }

    [Test]
    public async Task An_encryption_info_too_short_to_hold_a_header_is_rejected()
    {
        await Assert.That(() => StandardEncryptionDescriptor.Parse(new byte[8]))
            .Throws<XLEncryptionException>();
    }

    [Test]
    [Arguments(8)]
    [Arguments(100_000)]
    public async Task A_header_size_that_does_not_fit_the_stream_is_rejected(int headerSize)
    {
        // Too small to hold the fields the header must have, or larger than the stream itself.
        var info = EncryptionInfo(new byte[16], new byte[16], new byte[32], headerSizeOverride: headerSize);

        await Assert.That(() => StandardEncryptionDescriptor.Parse(info)).Throws<XLEncryptionException>();
    }

    [Test]
    public async Task A_salt_of_the_wrong_size_is_rejected()
    {
        // The scheme fixes the salt at 16 bytes, so a descriptor saying otherwise is malformed
        // rather than merely unusual.
        var info = EncryptionInfo(new byte[16], new byte[16], new byte[32], saltSizeOverride: 8);

        await Assert.That(() => StandardEncryptionDescriptor.Parse(info)).Throws<XLEncryptionException>();
    }

    [Test]
    public async Task A_package_shorter_than_its_length_prefix_is_rejected()
    {
        var (encryptionInfo, _) = Encrypt([1, 2, 3, 4], Password);
        var descriptor = StandardEncryptionDescriptor.Parse(encryptionInfo);
        var key = descriptor.DeriveAndVerifyKey(Password);

        await Assert.That(() => StandardEncryptionDescriptor.DecryptPackage(key, new byte[4]))
            .Throws<XLEncryptionException>();
    }

    [Test]
    public async Task A_package_claiming_more_content_than_it_carries_is_rejected()
    {
        var (encryptionInfo, encryptedPackage) = Encrypt([1, 2, 3, 4], Password);
        var descriptor = StandardEncryptionDescriptor.Parse(encryptionInfo);
        var key = descriptor.DeriveAndVerifyKey(Password);

        BitConverter.TryWriteBytes(encryptedPackage.AsSpan(0, 8), (long)100_000);

        await Assert.That(() => StandardEncryptionDescriptor.DecryptPackage(key, encryptedPackage))
            .Throws<XLEncryptionException>();
    }
}
