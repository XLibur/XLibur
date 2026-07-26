using System;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using XLibur.Excel.Exceptions;
using XLibur.Excel.IO.Encryption;
using XLibur.Excel.IO.Encryption.Agile;

namespace XLibur.Tests.Excel.Encryption;

/// <summary>
/// The agile encryption descriptor: what XLibur emits, what it accepts, and what it refuses.
/// </summary>
/// <remarks>
/// The specification allows far more combinations than Excel writes, and the design deliberately
/// rejects anything outside that profile by name instead of approximating it. Approximating would
/// turn an unusual file into plausible-looking garbage rather than an error, so the refusals are
/// part of the contract and are asserted here rather than assumed.
/// </remarks>
// Internal because some cases are parameterised by OfficeHashAlgorithm, which is internal.
internal class AgileDescriptorTests
{
    /// <summary>
    /// Wraps descriptor XML in the 8 byte EncryptionInfo header so it can go through the real parser.
    /// </summary>
    private static byte[] EncryptionInfo(string xml)
    {
        var body = Encoding.UTF8.GetBytes(xml);
        var stream = new byte[8 + body.Length];
        BitConverter.TryWriteBytes(stream.AsSpan(0), (ushort)4);
        BitConverter.TryWriteBytes(stream.AsSpan(2), (ushort)4);
        BitConverter.TryWriteBytes(stream.AsSpan(4), 0x40u);
        body.CopyTo(stream.AsSpan(8));
        return stream;
    }

    private static string Descriptor(
        string keyDataAttributes =
            "saltSize=\"16\" blockSize=\"16\" keyBits=\"256\" hashSize=\"64\" cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" saltValue=\"AAAAAAAAAAAAAAAAAAAAAA==\"",
        string encryptedKeyAttributes =
            "spinCount=\"100000\" saltSize=\"16\" blockSize=\"16\" keyBits=\"256\" hashSize=\"64\" cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" saltValue=\"AAAAAAAAAAAAAAAAAAAAAA==\" encryptedVerifierHashInput=\"AAAAAAAAAAAAAAAAAAAAAA==\" encryptedVerifierHashValue=\"AAAAAAAAAAAAAAAAAAAAAA==\" encryptedKeyValue=\"AAAAAAAAAAAAAAAAAAAAAA==\"",
        string keyEncryptorUri = "http://schemas.microsoft.com/office/2006/keyEncryptor/password")
    {
        return $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                            xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                  <keyData {keyDataAttributes}/>
                  <keyEncryptors>
                    <keyEncryptor uri="{keyEncryptorUri}">
                      <p:encryptedKey {encryptedKeyAttributes}/>
                    </keyEncryptor>
                  </keyEncryptors>
                </encryption>
                """;
    }

    [Test]
    [Arguments(OfficeHashAlgorithm.Sha1, 128)]
    [Arguments(OfficeHashAlgorithm.Sha256, 192)]
    [Arguments(OfficeHashAlgorithm.Sha384, 256)]
    [Arguments(OfficeHashAlgorithm.Sha512, 256)]
    public async Task What_XLibur_emits_is_what_XLibur_accepts(OfficeHashAlgorithm hash, int keyBits)
    {
        // The writer and the reader are separate bodies of code over the same XML. This is the test
        // that fails if one drifts from the other, which no round trip through a saved workbook
        // would localise as clearly.
        //
        // Driven across every supported algorithm even though XLibur only ever writes SHA-512:
        // the name each algorithm is written under has to be the name the parser reads back, and a
        // mismatch in the rarely used ones would otherwise sit undetected until a file needed them.
        var salt = new byte[16];
        RandomNumberGenerator.Fill(salt);

        var parameters = new AgileCipherParameters
        {
            SaltSize = 16,
            BlockSize = 16,
            KeyBits = keyBits,
            HashSize = hash.GetHashSize(),
            ChainingMode = CipherMode.CBC,
            HashAlgorithm = hash,
            SaltValue = salt,
        };

        var original = new AgileEncryptionDescriptor
        {
            KeyData = parameters,
            PasswordKeyEncryptor = new AgilePasswordKeyEncryptor
            {
                Parameters = parameters,
                SpinCount = 100_000,
                EncryptedVerifierHashInput = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16],
                EncryptedVerifierHashValue = [16, 15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1],
                EncryptedKeyValue = [2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32],
            },
            EncryptedHmacKey = [9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9],
            EncryptedHmacValue = [8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8],
        };

        var parsed = AgileEncryptionDescriptor.Parse(original.ToEncryptionInfo());

        await Assert.That(parsed.KeyData.SaltValue).IsEquivalentTo(salt);
        await Assert.That(parsed.KeyData.KeyBits).IsEqualTo(keyBits);
        await Assert.That(parsed.KeyData.HashAlgorithm).IsEqualTo(hash);
        await Assert.That(parsed.KeyData.HashSize).IsEqualTo(hash.GetHashSize());
        await Assert.That(parsed.KeyData.ChainingMode).IsEqualTo(CipherMode.CBC);
        await Assert.That(parsed.PasswordKeyEncryptor.SpinCount).IsEqualTo(100_000);
        await Assert.That(parsed.PasswordKeyEncryptor.EncryptedVerifierHashInput)
            .IsEquivalentTo(original.PasswordKeyEncryptor.EncryptedVerifierHashInput);
        await Assert.That(parsed.PasswordKeyEncryptor.EncryptedVerifierHashValue)
            .IsEquivalentTo(original.PasswordKeyEncryptor.EncryptedVerifierHashValue);
        await Assert.That(parsed.PasswordKeyEncryptor.EncryptedKeyValue)
            .IsEquivalentTo(original.PasswordKeyEncryptor.EncryptedKeyValue);
        await Assert.That(parsed.EncryptedHmacKey).IsEquivalentTo(original.EncryptedHmacKey);
        await Assert.That(parsed.EncryptedHmacValue).IsEquivalentTo(original.EncryptedHmacValue);
    }

    [Test]
    public async Task A_descriptor_without_data_integrity_is_accepted()
    {
        // dataIntegrity is optional. A file without it simply has no HMAC to check, which is not the
        // same as a file whose HMAC fails.
        var descriptor = AgileEncryptionDescriptor.Parse(EncryptionInfo(Descriptor()));

        await Assert.That(descriptor.EncryptedHmacKey).IsNull();
        await Assert.That(descriptor.EncryptedHmacValue).IsNull();
    }

    [Test]
    [Arguments("cipherAlgorithm=\"RC4\"", "RC4")]
    [Arguments("cipherAlgorithm=\"DES\"", "DES")]
    public async Task An_unsupported_cipher_is_named_in_the_error(string cipherAttribute, string expectedInMessage)
    {
        var keyData = "saltSize=\"16\" blockSize=\"16\" keyBits=\"256\" hashSize=\"64\" " + cipherAttribute +
                      " cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" saltValue=\"AAAAAAAAAAAAAAAAAAAAAA==\"";

        var exception = await Assert.That(() => AgileEncryptionDescriptor.Parse(EncryptionInfo(Descriptor(keyData))))
            .Throws<XLEncryptionException>();

        await Assert.That(exception!.Message).Contains(expectedInMessage);
    }

    [Test]
    public async Task The_unsupported_cfb_chaining_mode_is_called_out_by_name()
    {
        // CFB is the one alternative the specification defines, so it is worth a message of its own
        // rather than falling into the generic "unsupported" branch.
        var keyData = "saltSize=\"16\" blockSize=\"16\" keyBits=\"256\" hashSize=\"64\" cipherAlgorithm=\"AES\" " +
                      "cipherChaining=\"ChainingModeCFB\" hashAlgorithm=\"SHA512\" saltValue=\"AAAAAAAAAAAAAAAAAAAAAA==\"";

        var exception = await Assert.That(() => AgileEncryptionDescriptor.Parse(EncryptionInfo(Descriptor(keyData))))
            .Throws<XLEncryptionException>();

        await Assert.That(exception!.Message).Contains("ChainingModeCFB");
    }

    [Test]
    public async Task An_unsupported_hash_algorithm_is_named_in_the_error()
    {
        var keyData = "saltSize=\"16\" blockSize=\"16\" keyBits=\"256\" hashSize=\"64\" cipherAlgorithm=\"AES\" " +
                      "cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"MD5\" saltValue=\"AAAAAAAAAAAAAAAAAAAAAA==\"";

        var exception = await Assert.That(() => AgileEncryptionDescriptor.Parse(EncryptionInfo(Descriptor(keyData))))
            .Throws<XLEncryptionException>();

        await Assert.That(exception!.Message).Contains("MD5");
    }

    [Test]
    public async Task An_impossible_key_length_is_rejected()
    {
        var keyData = "saltSize=\"16\" blockSize=\"16\" keyBits=\"64\" hashSize=\"64\" cipherAlgorithm=\"AES\" " +
                      "cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" saltValue=\"AAAAAAAAAAAAAAAAAAAAAA==\"";

        var exception = await Assert.That(() => AgileEncryptionDescriptor.Parse(EncryptionInfo(Descriptor(keyData))))
            .Throws<XLEncryptionException>();

        await Assert.That(exception!.Message).Contains("64");
    }

    [Test]
    public async Task A_salt_whose_length_contradicts_its_declared_size_is_rejected()
    {
        // Declares 32 bytes of salt but carries 16. Trusting either number over the other would
        // silently derive the wrong key.
        var keyData = "saltSize=\"32\" blockSize=\"16\" keyBits=\"256\" hashSize=\"64\" cipherAlgorithm=\"AES\" " +
                      "cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" saltValue=\"AAAAAAAAAAAAAAAAAAAAAA==\"";

        await Assert.That(() => AgileEncryptionDescriptor.Parse(EncryptionInfo(Descriptor(keyData))))
            .Throws<XLEncryptionException>();
    }

    [Test]
    public async Task A_hash_size_that_contradicts_the_hash_algorithm_is_rejected()
    {
        // SHA-512 produces 64 bytes; a descriptor claiming 32 disagrees with itself.
        var keyData = "saltSize=\"16\" blockSize=\"16\" keyBits=\"256\" hashSize=\"32\" cipherAlgorithm=\"AES\" " +
                      "cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" saltValue=\"AAAAAAAAAAAAAAAAAAAAAA==\"";

        await Assert.That(() => AgileEncryptionDescriptor.Parse(EncryptionInfo(Descriptor(keyData))))
            .Throws<XLEncryptionException>();
    }

    [Test]
    public async Task A_missing_attribute_names_the_attribute()
    {
        var keyData = "blockSize=\"16\" keyBits=\"256\" hashSize=\"64\" cipherAlgorithm=\"AES\" " +
                      "cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" saltValue=\"AAAAAAAAAAAAAAAAAAAAAA==\"";

        var exception = await Assert.That(() => AgileEncryptionDescriptor.Parse(EncryptionInfo(Descriptor(keyData))))
            .Throws<XLEncryptionException>();

        await Assert.That(exception!.Message).Contains("saltSize");
    }

    [Test]
    public async Task A_non_numeric_attribute_names_the_attribute_and_the_value()
    {
        var keyData = "saltSize=\"sixteen\" blockSize=\"16\" keyBits=\"256\" hashSize=\"64\" cipherAlgorithm=\"AES\" " +
                      "cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" saltValue=\"AAAAAAAAAAAAAAAAAAAAAA==\"";

        var exception = await Assert.That(() => AgileEncryptionDescriptor.Parse(EncryptionInfo(Descriptor(keyData))))
            .Throws<XLEncryptionException>();

        await Assert.That(exception!.Message).Contains("saltSize");
        await Assert.That(exception.Message).Contains("sixteen");
    }

    [Test]
    public async Task Malformed_base64_names_the_attribute()
    {
        var keyData = "saltSize=\"16\" blockSize=\"16\" keyBits=\"256\" hashSize=\"64\" cipherAlgorithm=\"AES\" " +
                      "cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" saltValue=\"not base64 at all!\"";

        var exception = await Assert.That(() => AgileEncryptionDescriptor.Parse(EncryptionInfo(Descriptor(keyData))))
            .Throws<XLEncryptionException>();

        await Assert.That(exception!.Message).Contains("saltValue");
    }

    [Test]
    public async Task Malformed_xml_is_reported_as_such()
    {
        await Assert.That(() => AgileEncryptionDescriptor.Parse(EncryptionInfo("<encryption><unclosed>")))
            .Throws<XLEncryptionException>();
    }

    [Test]
    public async Task A_descriptor_with_no_key_data_is_rejected()
    {
        const string xml = """
                           <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                           <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"/>
                           """;

        var exception = await Assert.That(() => AgileEncryptionDescriptor.Parse(EncryptionInfo(xml)))
            .Throws<XLEncryptionException>();

        await Assert.That(exception!.Message).Contains("keyData");
    }

    [Test]
    public async Task A_descriptor_with_no_key_encryptors_is_rejected()
    {
        const string xml = """
                           <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                           <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption">
                             <keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES"
                                      cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
                           </encryption>
                           """;

        var exception = await Assert.That(() => AgileEncryptionDescriptor.Parse(EncryptionInfo(xml)))
            .Throws<XLEncryptionException>();

        await Assert.That(exception!.Message).Contains("keyEncryptors");
    }

    [Test]
    public async Task A_workbook_encrypted_to_a_certificate_says_so_rather_than_blaming_the_password()
    {
        // A certificate key encryptor is a well-formed file that no password can open. Reporting it
        // as a bad password would send the caller round a loop they can never get out of.
        var exception = await Assert.That(() => AgileEncryptionDescriptor.Parse(EncryptionInfo(Descriptor(
                keyEncryptorUri: "http://schemas.microsoft.com/office/2006/keyEncryptor/certificate"))))
            .Throws<XLEncryptionException>();

        await Assert.That(exception!.Message).Contains("password");
    }

    [Test]
    public async Task An_encryption_info_too_short_to_hold_xml_is_rejected()
    {
        await Assert.That(() => AgileEncryptionDescriptor.Parse(new byte[4]))
            .Throws<XLEncryptionException>();
    }
}
