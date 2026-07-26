using System;
using System.Security.Cryptography;
using System.Text;
using XLibur.Excel.Exceptions;

namespace XLibur.Excel.IO.Encryption.Standard;

/// <summary>
/// Standard encryption, the scheme Office 2007 used before agile encryption replaced it. Read only:
/// Excel has not written this format since 2007, so XLibur opens such files but never produces one.
/// </summary>
/// <remarks>
/// [MS-OFFCRYPTO] 2.3.4.5 through 2.3.4.9. Compared to agile encryption it is fixed rather than
/// described: SHA-1, a spin count of 50,000, ECB for the package, and no integrity check at all.
/// </remarks>
internal sealed class StandardEncryptionDescriptor
{
    private const int SpinCount = 50000;

    private const int VerifierLength = 16;

    // AlgID values from the EncryptionHeader. Anything else is a cipher XLibur does not read.
    private const uint AlgIdAes128 = 0x660E;
    private const uint AlgIdAes192 = 0x660F;
    private const uint AlgIdAes256 = 0x6610;

    private const uint AlgIdHashSha1 = 0x8004;

    public required int KeyBytes { get; init; }

    public required byte[] Salt { get; init; }

    public required byte[] EncryptedVerifier { get; init; }

    public required byte[] EncryptedVerifierHash { get; init; }

    public static StandardEncryptionDescriptor Parse(ReadOnlySpan<byte> encryptionInfo)
    {
        // 2 bytes major, 2 bytes minor, 4 bytes flags, 4 bytes header size, then the header.
        const int prologue = 12;
        if (encryptionInfo.Length < prologue)
            throw new XLEncryptionException("The EncryptionInfo stream is too short to be a standard descriptor.");

        var headerSize = BitConverter.ToInt32(encryptionInfo[8..]);
        if (headerSize < 32 || prologue + headerSize > encryptionInfo.Length)
            throw new XLEncryptionException($"The EncryptionInfo stream declares an unusable header size of {headerSize}.");

        var header = encryptionInfo.Slice(prologue, headerSize);
        var algId = BitConverter.ToUInt32(header[8..]);
        var algIdHash = BitConverter.ToUInt32(header[12..]);
        var keyBits = BitConverter.ToInt32(header[16..]);

        var keyBytes = algId switch
        {
            AlgIdAes128 => 16,
            AlgIdAes192 => 24,
            AlgIdAes256 => 32,
            _ => throw new XLEncryptionException(
                $"Unsupported cipher 0x{algId:X4} in a standard-encrypted workbook. XLibur reads AES only."),
        };

        // A zero AlgIDHash means "the default for this provider", which for standard encryption is SHA-1.
        if (algIdHash is not (0 or AlgIdHashSha1))
        {
            throw new XLEncryptionException(
                $"Unsupported hash 0x{algIdHash:X4} in a standard-encrypted workbook. XLibur reads SHA-1 only.");
        }

        // The header's own key size wins when it is set, since the AlgID is allowed to be generic.
        if (keyBits > 0)
        {
            OfficeCryptoAlgorithms.ValidateKeyBits(keyBits);
            keyBytes = keyBits / 8;
        }

        var verifier = encryptionInfo[(prologue + headerSize)..];
        if (verifier.Length < 8)
            throw new XLEncryptionException("The EncryptionInfo stream has no room for its verifier.");

        var saltSize = BitConverter.ToInt32(verifier);
        if (saltSize != VerifierLength || verifier.Length < 4 + saltSize + VerifierLength + 4)
            throw new XLEncryptionException("The standard encryption verifier is malformed.");

        var verifierHashSize = BitConverter.ToInt32(verifier[(4 + saltSize + VerifierLength)..]);
        var encryptedHashOffset = 4 + saltSize + VerifierLength + 4;

        // The stored hash is SHA-1, but it is written padded out to whole cipher blocks.
        var encryptedHashLength = verifier.Length - encryptedHashOffset;
        if (verifierHashSize <= 0 || encryptedHashLength < verifierHashSize)
            throw new XLEncryptionException("The standard encryption verifier hash is malformed.");

        return new StandardEncryptionDescriptor
        {
            KeyBytes = keyBytes,
            Salt = verifier.Slice(4, saltSize).ToArray(),
            EncryptedVerifier = verifier.Slice(4 + saltSize, VerifierLength).ToArray(),
            EncryptedVerifierHash = verifier.Slice(encryptedHashOffset, encryptedHashLength - encryptedHashLength % 16).ToArray(),
        };
    }

    /// <summary>
    /// Checks the password against the verifier and returns the key the package is encrypted with.
    /// Unlike agile encryption there is no wrapped key: this key is used directly.
    /// </summary>
    /// <exception cref="XLInvalidPasswordException">The password is wrong.</exception>
    public byte[] DeriveAndVerifyKey(string password)
    {
        var key = DeriveKey(password);
        try
        {
            var verifier = OfficeCryptoAlgorithms.AesTransform(
                key, [], EncryptedVerifier, CipherMode.ECB, encrypting: false);

            var storedHash = OfficeCryptoAlgorithms.AesTransform(
                key, [], EncryptedVerifierHash, CipherMode.ECB, encrypting: false);

            var actualHash = SHA1.HashData(verifier);

            if (!CryptographicOperations.FixedTimeEquals(actualHash, storedHash.AsSpan(0, actualHash.Length)))
                throw new XLInvalidPasswordException();

            return key;
        }
        catch
        {
            CryptographicOperations.ZeroMemory(key);
            throw;
        }
    }

    private byte[] DeriveKey(string password)
    {
        // H_0 = SHA1(salt + password), the password as UTF-16LE.
        var passwordBytes = Encoding.Unicode.GetBytes(password);
        var seed = new byte[Salt.Length + passwordBytes.Length];
        Salt.CopyTo(seed, 0);
        passwordBytes.CopyTo(seed, Salt.Length);

        var current = SHA1.HashData(seed);
        CryptographicOperations.ZeroMemory(passwordBytes);
        CryptographicOperations.ZeroMemory(seed);

        var spinBuffer = new byte[4 + current.Length];
        for (var i = 0; i < SpinCount; i++)
        {
            BitConverter.TryWriteBytes(spinBuffer.AsSpan(0, 4), i);
            current.CopyTo(spinBuffer, 4);
            current = SHA1.HashData(spinBuffer);
        }

        CryptographicOperations.ZeroMemory(spinBuffer);

        // H_final = SHA1(H_spinCount + LE32(0)). Standard encryption only ever uses block 0, since
        // the whole package is encrypted under a single key rather than per segment.
        var blockInput = new byte[current.Length + 4];
        current.CopyTo(blockInput, 0);
        var hFinal = SHA1.HashData(blockInput);
        CryptographicOperations.ZeroMemory(blockInput);
        CryptographicOperations.ZeroMemory(current);

        // Derive the key by the ipad/opad construction of [MS-OFFCRYPTO] 2.3.4.7. It resembles HMAC
        // but is not HMAC, so it is written out rather than delegated to HMACSHA1.
        var x1 = SHA1.HashData(XorWithPad(hFinal, 0x36));
        var x2 = SHA1.HashData(XorWithPad(hFinal, 0x5C));
        CryptographicOperations.ZeroMemory(hFinal);

        var x3 = new byte[x1.Length + x2.Length];
        x1.CopyTo(x3, 0);
        x2.CopyTo(x3, x1.Length);

        if (x3.Length < KeyBytes)
            throw new XLEncryptionException($"A {KeyBytes * 8} bit key cannot be derived from SHA-1.");

        var key = x3.AsSpan(0, KeyBytes).ToArray();
        CryptographicOperations.ZeroMemory(x1);
        CryptographicOperations.ZeroMemory(x2);
        CryptographicOperations.ZeroMemory(x3);
        return key;
    }

    private static byte[] XorWithPad(byte[] hash, byte pad)
    {
        var buffer = new byte[64];
        buffer.AsSpan().Fill(pad);
        for (var i = 0; i < hash.Length && i < buffer.Length; i++)
            buffer[i] ^= hash[i];

        return buffer;
    }

    /// <summary>
    /// Decrypts an <c>EncryptedPackage</c> stream. As with agile encryption the stream opens with
    /// the plaintext length, but the body is a single ECB run rather than addressable segments.
    /// </summary>
    public static byte[] DecryptPackage(byte[] key, byte[] encryptedPackage)
    {
        const int prefixLength = 8;
        if (encryptedPackage.Length < prefixLength)
            throw new XLEncryptionException("The EncryptedPackage stream is too short to hold its length prefix.");

        var declaredLength = BitConverter.ToInt64(encryptedPackage, 0);
        var ciphertextLength = encryptedPackage.Length - prefixLength;
        ciphertextLength -= ciphertextLength % 16;

        if (declaredLength < 0 || declaredLength > ciphertextLength)
        {
            throw new XLEncryptionException(
                $"The EncryptedPackage stream declares {declaredLength} bytes of content but carries room for {ciphertextLength}.");
        }

        var decrypted = OfficeCryptoAlgorithms.AesTransform(
            key, [], encryptedPackage.AsSpan(prefixLength, ciphertextLength), CipherMode.ECB, encrypting: false);

        var plaintext = decrypted.AsSpan(0, (int)declaredLength).ToArray();
        CryptographicOperations.ZeroMemory(decrypted);
        return plaintext;
    }
}
