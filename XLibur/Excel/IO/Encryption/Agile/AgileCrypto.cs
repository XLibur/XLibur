using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using XLibur.Excel.Exceptions;

namespace XLibur.Excel.IO.Encryption.Agile;

/// <summary>
/// Agile encryption, the scheme Office 2010 and later use: an iterated hash turns the password into
/// an intermediate key, that key unwraps the key the package is really encrypted with, and the
/// package is encrypted in independently addressable segments.
/// </summary>
/// <remarks>
/// [MS-OFFCRYPTO] 2.3.4.11 through 2.3.4.15.
/// </remarks>
internal static class AgileCrypto
{
    /// <summary>
    /// Plaintext block size of the segmented package cipher. Fixed by the specification, not a
    /// tunable: it is what makes a segment's IV derivable from its index.
    /// </summary>
    internal const int SegmentLength = 4096;

    // Block keys from [MS-OFFCRYPTO] 2.3.4.12/2.3.4.14. Each one makes the same password produce a
    // different key, so the verifier, the wrapped key and the integrity check never share one.
    private static ReadOnlySpan<byte> BlockVerifierHashInput => [0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79];

    private static ReadOnlySpan<byte> BlockVerifierHashValue => [0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e];

    private static ReadOnlySpan<byte> BlockKeyValue => [0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6];

    private static ReadOnlySpan<byte> BlockHmacKey => [0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6];

    private static ReadOnlySpan<byte> BlockHmacValue => [0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33];

    /// <summary>
    /// Derives the intermediate key for one block key by spinning the password through the hash
    /// <c>spinCount</c> times. The spin is the whole cost of a wrong-password guess and is why the
    /// count is six figures.
    /// </summary>
    internal static byte[] DeriveKey(
        string password,
        AgileCipherParameters parameters,
        int spinCount,
        ReadOnlySpan<byte> blockKey)
    {
        var hash = parameters.HashAlgorithm;

        // H_0 = H(salt + password), the password as UTF-16LE with no terminator.
        var passwordBytes = Encoding.Unicode.GetBytes(password);
        var seed = new byte[parameters.SaltValue.Length + passwordBytes.Length];
        parameters.SaltValue.CopyTo(seed, 0);
        passwordBytes.CopyTo(seed, parameters.SaltValue.Length);

        var current = hash.Hash(seed);
        CryptographicOperations.ZeroMemory(passwordBytes);
        CryptographicOperations.ZeroMemory(seed);

        // H_n = H(LE32(n) + H_n-1). The buffer is reused across iterations; at a spin count of
        // 100,000 the allocations would otherwise dominate the loop.
        var spinBuffer = new byte[4 + current.Length];
        for (var i = 0; i < spinCount; i++)
        {
            BitConverter.TryWriteBytes(spinBuffer.AsSpan(0, 4), i);
            current.CopyTo(spinBuffer, 4);
            current = hash.Hash(spinBuffer);
        }

        CryptographicOperations.ZeroMemory(spinBuffer);

        // H_final = H(H_spinCount + blockKey), then cut or padded to the key length.
        var finalInput = new byte[current.Length + blockKey.Length];
        current.CopyTo(finalInput, 0);
        blockKey.CopyTo(finalInput.AsSpan(current.Length));

        var finalHash = hash.Hash(finalInput);
        CryptographicOperations.ZeroMemory(finalInput);
        CryptographicOperations.ZeroMemory(current);

        var key = OfficeCryptoAlgorithms.Fit(finalHash, parameters.KeyBytes, 0x36);
        CryptographicOperations.ZeroMemory(finalHash);
        return key;
    }

    /// <summary>
    /// Checks the password against the verifier and returns the key the package is encrypted with.
    /// </summary>
    /// <exception cref="XLInvalidPasswordException">The password is wrong.</exception>
    internal static byte[] DecryptPackageKey(AgileEncryptionDescriptor descriptor, string password)
    {
        var encryptor = descriptor.PasswordKeyEncryptor;
        var parameters = encryptor.Parameters;

        // The verifier is a random value stored twice: encrypted, and hashed then encrypted. A
        // password that reproduces the hash of what it just decrypted is the right password.
        var verifierInputKey = DeriveKey(password, parameters, encryptor.SpinCount, BlockVerifierHashInput);
        var verifierValueKey = DeriveKey(password, parameters, encryptor.SpinCount, BlockVerifierHashValue);

        try
        {
            var verifierInput = OfficeCryptoAlgorithms.AesTransform(
                verifierInputKey, parameters.SaltValue, encryptor.EncryptedVerifierHashInput,
                parameters.ChainingMode, encrypting: false);

            var expectedHash = OfficeCryptoAlgorithms.AesTransform(
                verifierValueKey, parameters.SaltValue, encryptor.EncryptedVerifierHashValue,
                parameters.ChainingMode, encrypting: false);

            var actualHash = parameters.HashAlgorithm.Hash(verifierInput.AsSpan(0, parameters.SaltSize));

            if (!CryptographicOperations.FixedTimeEquals(
                    actualHash.AsSpan(0, parameters.HashSize),
                    expectedHash.AsSpan(0, parameters.HashSize)))
            {
                throw new XLInvalidPasswordException();
            }

            var keyKey = DeriveKey(password, parameters, encryptor.SpinCount, BlockKeyValue);
            try
            {
                var keyValue = OfficeCryptoAlgorithms.AesTransform(
                    keyKey, parameters.SaltValue, encryptor.EncryptedKeyValue,
                    parameters.ChainingMode, encrypting: false);

                return OfficeCryptoAlgorithms.Fit(keyValue, descriptor.KeyData.KeyBytes, 0x00);
            }
            finally
            {
                CryptographicOperations.ZeroMemory(keyKey);
            }
        }
        finally
        {
            CryptographicOperations.ZeroMemory(verifierInputKey);
            CryptographicOperations.ZeroMemory(verifierValueKey);
        }
    }

    /// <summary>
    /// IV for one segment of the package, or for one of the integrity blobs when
    /// <paramref name="blockKey"/> is a block key rather than a segment index.
    /// </summary>
    private static byte[] SegmentIv(AgileCipherParameters keyData, ReadOnlySpan<byte> blockKey)
    {
        var input = new byte[keyData.SaltValue.Length + blockKey.Length];
        keyData.SaltValue.CopyTo(input, 0);
        blockKey.CopyTo(input.AsSpan(keyData.SaltValue.Length));

        var hash = keyData.HashAlgorithm.Hash(input);
        return OfficeCryptoAlgorithms.Fit(hash, keyData.BlockSize, 0x36);
    }

    private static byte[] SegmentIv(AgileCipherParameters keyData, int segmentIndex)
    {
        Span<byte> blockKey = stackalloc byte[4];
        BitConverter.TryWriteBytes(blockKey, segmentIndex);
        return SegmentIv(keyData, blockKey);
    }

    /// <summary>
    /// Decrypts an <c>EncryptedPackage</c> stream into the .xlsx bytes it carries. The stream opens
    /// with the plaintext length as a little-endian 64 bit integer, then the segmented ciphertext.
    /// </summary>
    internal static byte[] DecryptPackage(AgileEncryptionDescriptor descriptor, byte[] packageKey, byte[] encryptedPackage)
    {
        const int prefixLength = 8;
        if (encryptedPackage.Length < prefixLength)
            throw new XLEncryptionException("The EncryptedPackage stream is too short to hold its length prefix.");

        var declaredLength = BitConverter.ToInt64(encryptedPackage, 0);
        var ciphertextLength = encryptedPackage.Length - prefixLength;
        if (declaredLength < 0 || declaredLength > ciphertextLength)
        {
            throw new XLEncryptionException(
                $"The EncryptedPackage stream declares {declaredLength} bytes of content but carries room for {ciphertextLength}.");
        }

        var keyData = descriptor.KeyData;
        var plaintext = new byte[declaredLength];
        var written = 0;

        // The loop ends when the declared content has been recovered, not when the segments run
        // out — a trailing segment can carry padding past the declared length.
        var segmentIndex = 0;
        while (written < declaredLength)
        {
            var offset = prefixLength + segmentIndex * SegmentLength;
            var segmentLength = Math.Min(SegmentLength, encryptedPackage.Length - offset);
            if (segmentLength <= 0)
                break;

            // A trailing segment is padded up to the cipher's block size; decrypt the whole thing
            // and let the declared length decide how much of it counts.
            segmentLength -= segmentLength % keyData.BlockSize;
            if (segmentLength == 0)
                break;

            var iv = SegmentIv(keyData, segmentIndex);
            var decrypted = OfficeCryptoAlgorithms.AesTransform(
                packageKey, iv, encryptedPackage.AsSpan(offset, segmentLength),
                keyData.ChainingMode, encrypting: false);

            var copyLength = (int)Math.Min(decrypted.Length, declaredLength - written);
            decrypted.AsSpan(0, copyLength).CopyTo(plaintext.AsSpan(written));
            written += copyLength;

            CryptographicOperations.ZeroMemory(decrypted);
            segmentIndex++;
        }

        if (written != declaredLength)
            throw new XLEncryptionException("The EncryptedPackage stream ended before the declared content length.");

        return plaintext;
    }

    /// <summary>
    /// Encrypts .xlsx bytes into the body of an <c>EncryptedPackage</c> stream, length prefix included.
    /// </summary>
    internal static byte[] EncryptPackage(AgileEncryptionDescriptor descriptor, byte[] packageKey, byte[] package)
    {
        var keyData = descriptor.KeyData;
        var segmentCount = (package.Length + SegmentLength - 1) / SegmentLength;

        using var output = new MemoryStream(8 + segmentCount * SegmentLength + keyData.BlockSize);
        Span<byte> prefix = stackalloc byte[8];
        BitConverter.TryWriteBytes(prefix, (long)package.Length);
        output.Write(prefix);

        for (var segmentIndex = 0; segmentIndex < segmentCount; segmentIndex++)
        {
            var offset = segmentIndex * SegmentLength;
            var length = Math.Min(SegmentLength, package.Length - offset);

            // CBC needs whole blocks, so a short final segment is zero padded. The length prefix is
            // what tells a reader where the real content stops, so the padding is never ambiguous.
            var padded = length % keyData.BlockSize == 0
                ? length
                : length + (keyData.BlockSize - length % keyData.BlockSize);

            var segment = new byte[padded];
            package.AsSpan(offset, length).CopyTo(segment);

            var iv = SegmentIv(keyData, segmentIndex);
            var encrypted = OfficeCryptoAlgorithms.AesTransform(
                packageKey, iv, segment, keyData.ChainingMode, encrypting: true);

            output.Write(encrypted, 0, encrypted.Length);
            CryptographicOperations.ZeroMemory(segment);
        }

        return output.ToArray();
    }

    /// <summary>
    /// Verifies the HMAC over the encrypted package. Detects a file altered after it was written,
    /// which a correct password would otherwise turn into plausible-looking garbage.
    /// </summary>
    internal static void VerifyIntegrity(AgileEncryptionDescriptor descriptor, byte[] packageKey, byte[] encryptedPackage)
    {
        if (descriptor.EncryptedHmacKey is null || descriptor.EncryptedHmacValue is null)
            return;

        var keyData = descriptor.KeyData;

        var hmacKeyIv = SegmentIv(keyData, BlockHmacKey);
        var hmacKey = OfficeCryptoAlgorithms.AesTransform(
            packageKey, hmacKeyIv, descriptor.EncryptedHmacKey, keyData.ChainingMode, encrypting: false);

        var hmacValueIv = SegmentIv(keyData, BlockHmacValue);
        var expected = OfficeCryptoAlgorithms.AesTransform(
            packageKey, hmacValueIv, descriptor.EncryptedHmacValue, keyData.ChainingMode, encrypting: false);

        try
        {
            var truncatedKey = hmacKey.AsSpan(0, keyData.HashSize).ToArray();
            var actual = keyData.HashAlgorithm.Hmac(truncatedKey, encryptedPackage);
            CryptographicOperations.ZeroMemory(truncatedKey);

            if (!CryptographicOperations.FixedTimeEquals(actual, expected.AsSpan(0, keyData.HashSize)))
            {
                throw new XLEncryptionException(
                    "The integrity check over the encrypted package failed. The file has been altered or is corrupt.");
            }
        }
        finally
        {
            CryptographicOperations.ZeroMemory(hmacKey);
            CryptographicOperations.ZeroMemory(expected);
        }
    }

    /// <summary>
    /// Produces the encrypted HMAC key and value for a package that has just been encrypted.
    /// </summary>
    internal static (byte[] EncryptedKey, byte[] EncryptedValue) CreateIntegrity(
        AgileCipherParameters keyData,
        byte[] packageKey,
        byte[] encryptedPackage)
    {
        var hmacKey = new byte[keyData.HashSize];
        RandomNumberGenerator.Fill(hmacKey);

        try
        {
            var actual = keyData.HashAlgorithm.Hmac(hmacKey, encryptedPackage);

            var encryptedKey = OfficeCryptoAlgorithms.AesTransform(
                packageKey, SegmentIv(keyData, BlockHmacKey), Pad(hmacKey, keyData.BlockSize),
                keyData.ChainingMode, encrypting: true);

            var encryptedValue = OfficeCryptoAlgorithms.AesTransform(
                packageKey, SegmentIv(keyData, BlockHmacValue), Pad(actual, keyData.BlockSize),
                keyData.ChainingMode, encrypting: true);

            return (encryptedKey, encryptedValue);
        }
        finally
        {
            CryptographicOperations.ZeroMemory(hmacKey);
        }
    }

    /// <summary>
    /// Builds the verifier and the wrapped package key for a new descriptor.
    /// </summary>
    internal static (byte[] EncryptedVerifierHashInput, byte[] EncryptedVerifierHashValue, byte[] EncryptedKeyValue)
        CreateVerifier(AgileCipherParameters parameters, int spinCount, string password, byte[] packageKey)
    {
        var verifierInput = new byte[parameters.SaltSize];
        RandomNumberGenerator.Fill(verifierInput);

        var verifierInputKey = DeriveKey(password, parameters, spinCount, BlockVerifierHashInput);
        var verifierValueKey = DeriveKey(password, parameters, spinCount, BlockVerifierHashValue);
        var keyKey = DeriveKey(password, parameters, spinCount, BlockKeyValue);

        try
        {
            var encryptedVerifierHashInput = OfficeCryptoAlgorithms.AesTransform(
                verifierInputKey, parameters.SaltValue, Pad(verifierInput, parameters.BlockSize),
                parameters.ChainingMode, encrypting: true);

            var verifierHash = parameters.HashAlgorithm.Hash(verifierInput);
            var encryptedVerifierHashValue = OfficeCryptoAlgorithms.AesTransform(
                verifierValueKey, parameters.SaltValue, Pad(verifierHash, parameters.BlockSize),
                parameters.ChainingMode, encrypting: true);

            var encryptedKeyValue = OfficeCryptoAlgorithms.AesTransform(
                keyKey, parameters.SaltValue, Pad(packageKey, parameters.BlockSize),
                parameters.ChainingMode, encrypting: true);

            return (encryptedVerifierHashInput, encryptedVerifierHashValue, encryptedKeyValue);
        }
        finally
        {
            CryptographicOperations.ZeroMemory(verifierInput);
            CryptographicOperations.ZeroMemory(verifierInputKey);
            CryptographicOperations.ZeroMemory(verifierValueKey);
            CryptographicOperations.ZeroMemory(keyKey);
        }
    }

    /// <summary>Zero pads to a whole number of cipher blocks.</summary>
    private static byte[] Pad(byte[] value, int blockSize)
    {
        if (value.Length % blockSize == 0)
            return value;

        var padded = new byte[value.Length + (blockSize - value.Length % blockSize)];
        value.CopyTo(padded, 0);
        return padded;
    }
}
