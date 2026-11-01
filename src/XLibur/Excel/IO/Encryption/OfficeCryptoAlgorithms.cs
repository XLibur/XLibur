using System;
using System.Security.Cryptography;
using XLibur.Excel.Exceptions;

namespace XLibur.Excel.IO.Encryption;

/// <summary>
/// The hash algorithms MS-OFFCRYPTO names in an encryption descriptor.
/// </summary>
/// <remarks>
/// The specification allows more than Excel ever writes. Everything outside the set below is
/// rejected by name rather than approximated, so an unusual file fails with a message that says
/// which algorithm it wanted instead of silently producing wrong plaintext.
/// </remarks>
internal enum OfficeHashAlgorithm
{
    Sha1,
    Sha256,
    Sha384,
    Sha512,
}

internal static class OfficeCryptoAlgorithms
{
    /// <summary>Size of the hash in bytes.</summary>
    public static int GetHashSize(this OfficeHashAlgorithm algorithm) => algorithm switch
    {
        OfficeHashAlgorithm.Sha1 => 20,
        OfficeHashAlgorithm.Sha256 => 32,
        OfficeHashAlgorithm.Sha384 => 48,
        OfficeHashAlgorithm.Sha512 => 64,
        _ => throw new XLEncryptionException($"Unsupported hash algorithm '{algorithm}'."),
    };

    public static byte[] Hash(this OfficeHashAlgorithm algorithm, ReadOnlySpan<byte> data) => algorithm switch
    {
        OfficeHashAlgorithm.Sha1 => SHA1.HashData(data),
        OfficeHashAlgorithm.Sha256 => SHA256.HashData(data),
        OfficeHashAlgorithm.Sha384 => SHA384.HashData(data),
        OfficeHashAlgorithm.Sha512 => SHA512.HashData(data),
        _ => throw new XLEncryptionException($"Unsupported hash algorithm '{algorithm}'."),
    };

    public static byte[] Hmac(this OfficeHashAlgorithm algorithm, byte[] key, ReadOnlySpan<byte> data) => algorithm switch
    {
        OfficeHashAlgorithm.Sha1 => HMACSHA1.HashData(key, data),
        OfficeHashAlgorithm.Sha256 => HMACSHA256.HashData(key, data),
        OfficeHashAlgorithm.Sha384 => HMACSHA384.HashData(key, data),
        OfficeHashAlgorithm.Sha512 => HMACSHA512.HashData(key, data),
        _ => throw new XLEncryptionException($"Unsupported hash algorithm '{algorithm}'."),
    };

    public static OfficeHashAlgorithm ParseHashAlgorithm(string? name) => name switch
    {
        "SHA1" or "SHA-1" => OfficeHashAlgorithm.Sha1,
        "SHA256" or "SHA-256" => OfficeHashAlgorithm.Sha256,
        "SHA384" or "SHA-384" => OfficeHashAlgorithm.Sha384,
        "SHA512" or "SHA-512" => OfficeHashAlgorithm.Sha512,
        null => throw new XLEncryptionException("The encryption descriptor does not name a hash algorithm."),
        _ => throw new XLEncryptionException(
            $"Unsupported hash algorithm '{name}'. XLibur reads the algorithms Excel writes: SHA1, SHA256, SHA384 and SHA512."),
    };

    public static string ToXmlName(this OfficeHashAlgorithm algorithm) => algorithm switch
    {
        OfficeHashAlgorithm.Sha1 => "SHA1",
        OfficeHashAlgorithm.Sha256 => "SHA256",
        OfficeHashAlgorithm.Sha384 => "SHA384",
        OfficeHashAlgorithm.Sha512 => "SHA512",
        _ => throw new XLEncryptionException($"Unsupported hash algorithm '{algorithm}'."),
    };

    /// <summary>
    /// Validates the cipher named by a descriptor. Only AES is accepted, which is all Excel writes
    /// for the encryption versions XLibur supports.
    /// </summary>
    public static void ValidateCipherIsAes(string? cipherAlgorithm)
    {
        if (!string.Equals(cipherAlgorithm, "AES", StringComparison.Ordinal))
        {
            throw new XLEncryptionException(
                $"Unsupported cipher algorithm '{cipherAlgorithm}'. XLibur supports AES, which is what Excel writes.");
        }
    }

    public static CipherMode ParseChainingMode(string? cipherChaining) => cipherChaining switch
    {
        "ChainingModeCBC" => CipherMode.CBC,
        "ChainingModeCFB" => throw new XLEncryptionException(
            "Cipher chaining mode 'ChainingModeCFB' is not supported. XLibur supports CBC, which is what Excel writes."),
        null => throw new XLEncryptionException("The encryption descriptor does not name a cipher chaining mode."),
        _ => throw new XLEncryptionException($"Unsupported cipher chaining mode '{cipherChaining}'."),
    };

    public static void ValidateKeyBits(int keyBits)
    {
        if (keyBits is not (128 or 192 or 256))
        {
            throw new XLEncryptionException(
                $"Unsupported key length of {keyBits} bits. AES keys are 128, 192 or 256 bits.");
        }
    }

    /// <summary>
    /// Runs AES over one buffer with an explicit key and IV and no padding, which is how every
    /// MS-OFFCRYPTO operation uses the cipher. The caller has already sized the input to a whole
    /// number of blocks.
    /// </summary>
    public static byte[] AesTransform(
        byte[] key,
        byte[] iv,
        ReadOnlySpan<byte> input,
        CipherMode mode,
        bool encrypting)
    {
        using var aes = Aes.Create();
        aes.KeySize = key.Length * 8;
        aes.Key = key;
        aes.Mode = mode;
        aes.Padding = PaddingMode.None;

        if (mode != CipherMode.ECB)
            aes.IV = iv;

        using var transform = encrypting ? aes.CreateEncryptor() : aes.CreateDecryptor();

        var buffer = input.ToArray();
        try
        {
            return transform.TransformFinalBlock(buffer, 0, buffer.Length);
        }
        finally
        {
            CryptographicOperations.ZeroMemory(buffer);
        }
    }

    /// <summary>
    /// Grows or trims <paramref name="value"/> to <paramref name="size"/> bytes, padding with
    /// <paramref name="padding"/>. MS-OFFCRYPTO relies on this wherever a hash and a key or block
    /// size disagree, e.g. a SHA-512 hash feeding a 128-bit key.
    /// </summary>
    public static byte[] Fit(ReadOnlySpan<byte> value, int size, byte padding)
    {
        var result = new byte[size];
        if (value.Length >= size)
        {
            value[..size].CopyTo(result);
            return result;
        }

        value.CopyTo(result);
        result.AsSpan(value.Length).Fill(padding);
        return result;
    }
}
