using System;
using System.Security.Cryptography;
using System.Threading.Tasks;
using XLibur.Excel.Exceptions;
using XLibur.Excel.IO.Encryption;
using XLibur.Excel.IO.Encryption.Agile;

namespace XLibur.Tests.Excel.Encryption;

/// <summary>
/// The agile cipher below the workbook API: algorithm agility, segment handling and the integrity
/// check.
/// </summary>
/// <remarks>
/// XLibur writes one profile, AES-256 with SHA-512, so saving and reloading a workbook only ever
/// exercises that one. It has to <em>read</em> whatever the specification allows, and other
/// producers do pick differently, so the combinations are driven directly here rather than through
/// a saved file that could never contain them.
/// </remarks>
// Internal because the test cases are parameterised by OfficeHashAlgorithm, which is internal.
internal class AgileCryptoTests
{
    private const string Password = "a password";

    // A low spin count: these tests exercise algorithm selection, not the cost of the spin, and
    // 100,000 iterations per combination would dominate the suite's runtime for no added signal.
    private const int SpinCount = 64;

    private static AgileCipherParameters Parameters(OfficeHashAlgorithm hash, int keyBits)
    {
        var salt = new byte[16];
        RandomNumberGenerator.Fill(salt);

        return new AgileCipherParameters
        {
            SaltSize = 16,
            BlockSize = 16,
            KeyBits = keyBits,
            HashSize = hash.GetHashSize(),
            ChainingMode = CipherMode.CBC,
            HashAlgorithm = hash,
            SaltValue = salt,
        };
    }

    /// <summary>
    /// Encrypts a payload under the given profile and returns the descriptor describing it.
    /// </summary>
    private static (AgileEncryptionDescriptor Descriptor, byte[] EncryptedPackage) Encrypt(
        byte[] payload, OfficeHashAlgorithm hash, int keyBits)
    {
        var keyData = Parameters(hash, keyBits);
        var keyEncryptor = Parameters(hash, keyBits);

        var packageKey = new byte[keyBits / 8];
        RandomNumberGenerator.Fill(packageKey);

        var (verifierInput, verifierValue, encryptedKeyValue) =
            AgileCrypto.CreateVerifier(keyEncryptor, SpinCount, Password, packageKey);

        var withoutIntegrity = new AgileEncryptionDescriptor
        {
            KeyData = keyData,
            PasswordKeyEncryptor = new AgilePasswordKeyEncryptor
            {
                Parameters = keyEncryptor,
                SpinCount = SpinCount,
                EncryptedVerifierHashInput = verifierInput,
                EncryptedVerifierHashValue = verifierValue,
                EncryptedKeyValue = encryptedKeyValue,
            },
        };

        var encryptedPackage = AgileCrypto.EncryptPackage(withoutIntegrity, packageKey, payload);
        var (hmacKey, hmacValue) = AgileCrypto.CreateIntegrity(keyData, packageKey, encryptedPackage);

        var descriptor = new AgileEncryptionDescriptor
        {
            KeyData = keyData,
            PasswordKeyEncryptor = withoutIntegrity.PasswordKeyEncryptor,
            EncryptedHmacKey = hmacKey,
            EncryptedHmacValue = hmacValue,
        };

        return (descriptor, encryptedPackage);
    }

    [Test]
    [Arguments(OfficeHashAlgorithm.Sha1, 128)]
    [Arguments(OfficeHashAlgorithm.Sha1, 256)]
    [Arguments(OfficeHashAlgorithm.Sha256, 128)]
    [Arguments(OfficeHashAlgorithm.Sha256, 192)]
    [Arguments(OfficeHashAlgorithm.Sha256, 256)]
    [Arguments(OfficeHashAlgorithm.Sha384, 256)]
    [Arguments(OfficeHashAlgorithm.Sha512, 128)]
    [Arguments(OfficeHashAlgorithm.Sha512, 256)]
    public async Task Every_supported_hash_and_key_length_round_trips(OfficeHashAlgorithm hash, int keyBits)
    {
        // SHA-1 with a 256 bit key is the interesting corner: the hash is 20 bytes and the key needs
        // 32, so the derived material has to be padded rather than truncated.
        var payload = new byte[5000];
        RandomNumberGenerator.Fill(payload);

        var (descriptor, encryptedPackage) = Encrypt(payload, hash, keyBits);

        var packageKey = AgileCrypto.DecryptPackageKey(descriptor, Password);
        AgileCrypto.VerifyIntegrity(descriptor, packageKey, encryptedPackage);
        var decrypted = AgileCrypto.DecryptPackage(descriptor, packageKey, encryptedPackage);

        await Assert.That(decrypted).IsEquivalentTo(payload);
    }

    [Test]
    [Arguments(0)]
    [Arguments(1)]
    [Arguments(15)]
    [Arguments(16)]
    [Arguments(4095)]
    [Arguments(4096)]
    [Arguments(4097)]
    [Arguments(8192)]
    [Arguments(12289)]
    public async Task Payloads_around_the_segment_and_block_boundaries_round_trip(int length)
    {
        // The package is encrypted in 4096 byte segments, each with its own IV, and a trailing
        // segment is padded to a whole cipher block. Lengths either side of both boundaries are
        // where an off-by-one in that loop shows up; an ordinary workbook is far too big to land on
        // one by accident.
        var payload = new byte[length];
        RandomNumberGenerator.Fill(payload);

        var (descriptor, encryptedPackage) = Encrypt(payload, OfficeHashAlgorithm.Sha512, 256);

        var packageKey = AgileCrypto.DecryptPackageKey(descriptor, Password);
        var decrypted = AgileCrypto.DecryptPackage(descriptor, packageKey, encryptedPackage);

        await Assert.That(decrypted.Length).IsEqualTo(length);
        await Assert.That(decrypted).IsEquivalentTo(payload);
    }

    [Test]
    public async Task The_length_prefix_is_what_decides_the_plaintext_length()
    {
        // Ciphertext is padded to whole blocks, so the prefix is the only thing that says where the
        // content stops. A payload that is not a multiple of the block size proves the padding is
        // not being handed back as content.
        var payload = new byte[100];
        RandomNumberGenerator.Fill(payload);

        var (descriptor, encryptedPackage) = Encrypt(payload, OfficeHashAlgorithm.Sha512, 256);

        // 8 byte prefix plus 112 bytes of ciphertext, being 100 padded up to seven 16 byte blocks.
        await Assert.That(encryptedPackage.Length).IsEqualTo(8 + 112);

        var packageKey = AgileCrypto.DecryptPackageKey(descriptor, Password);
        await Assert.That(AgileCrypto.DecryptPackage(descriptor, packageKey, encryptedPackage).Length).IsEqualTo(100);
    }

    [Test]
    public async Task A_wrong_password_is_refused_before_anything_is_decrypted()
    {
        var (descriptor, _) = Encrypt([1, 2, 3], OfficeHashAlgorithm.Sha512, 256);

        await Assert.That(() => AgileCrypto.DecryptPackageKey(descriptor, "the wrong password"))
            .Throws<XLInvalidPasswordException>();
    }

    [Test]
    public async Task A_package_shorter_than_its_length_prefix_is_rejected()
    {
        var (descriptor, _) = Encrypt([1, 2, 3], OfficeHashAlgorithm.Sha512, 256);
        var packageKey = AgileCrypto.DecryptPackageKey(descriptor, Password);

        await Assert.That(() => AgileCrypto.DecryptPackage(descriptor, packageKey, new byte[4]))
            .Throws<XLEncryptionException>();
    }

    [Test]
    public async Task A_package_claiming_more_content_than_it_carries_is_rejected()
    {
        var (descriptor, encryptedPackage) = Encrypt([1, 2, 3], OfficeHashAlgorithm.Sha512, 256);
        var packageKey = AgileCrypto.DecryptPackageKey(descriptor, Password);

        // Truncation is the shape a partial download or a cut-off copy takes. The declared length
        // then exceeds what is there, and inventing the difference would hand back invented data.
        var truncated = new byte[encryptedPackage.Length];
        encryptedPackage.CopyTo(truncated, 0);
        BitConverter.TryWriteBytes(truncated.AsSpan(0, 8), (long)100_000);

        await Assert.That(() => AgileCrypto.DecryptPackage(descriptor, packageKey, truncated))
            .Throws<XLEncryptionException>();
    }

    [Test]
    public async Task A_descriptor_without_data_integrity_skips_the_check_rather_than_failing_it()
    {
        var payload = new byte[64];
        RandomNumberGenerator.Fill(payload);

        var (descriptor, encryptedPackage) = Encrypt(payload, OfficeHashAlgorithm.Sha512, 256);
        var withoutIntegrity = new AgileEncryptionDescriptor
        {
            KeyData = descriptor.KeyData,
            PasswordKeyEncryptor = descriptor.PasswordKeyEncryptor,
        };

        var packageKey = AgileCrypto.DecryptPackageKey(withoutIntegrity, Password);

        await Assert.That(() => AgileCrypto.VerifyIntegrity(withoutIntegrity, packageKey, encryptedPackage))
            .ThrowsNothing();
    }

    [Test]
    public async Task The_integrity_check_notices_a_single_flipped_bit()
    {
        var payload = new byte[4096];
        RandomNumberGenerator.Fill(payload);

        var (descriptor, encryptedPackage) = Encrypt(payload, OfficeHashAlgorithm.Sha512, 256);
        var packageKey = AgileCrypto.DecryptPackageKey(descriptor, Password);

        encryptedPackage[encryptedPackage.Length / 2] ^= 0x01;

        await Assert.That(() => AgileCrypto.VerifyIntegrity(descriptor, packageKey, encryptedPackage))
            .Throws<XLEncryptionException>();
    }

    [Test]
    public async Task Each_save_uses_fresh_random_material()
    {
        // Salts and the package key are generated per save. If any of them were fixed, two saves of
        // the same content under the same password would produce identical ciphertext, which leaks
        // that the content is unchanged.
        var payload = new byte[256];
        RandomNumberGenerator.Fill(payload);

        var (first, firstPackage) = Encrypt(payload, OfficeHashAlgorithm.Sha512, 256);
        var (second, secondPackage) = Encrypt(payload, OfficeHashAlgorithm.Sha512, 256);

        await Assert.That(first.KeyData.SaltValue).IsNotEquivalentTo(second.KeyData.SaltValue);
        await Assert.That(firstPackage).IsNotEquivalentTo(secondPackage);
    }
}
