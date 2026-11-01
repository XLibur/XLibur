using System;
using System.IO;
using System.Security.Cryptography;
using XLibur.Excel.Exceptions;
using XLibur.Excel.IO.Encryption.Agile;
using XLibur.Excel.IO.Encryption.Standard;

namespace XLibur.Excel.IO.Encryption;

/// <summary>
/// Turns a password-protected workbook into the .xlsx bytes inside it, and back.
/// </summary>
internal static class WorkbookEncryption
{
    /// <summary>
    /// Parameters for a newly encrypted workbook. These are the values Excel itself writes, chosen
    /// so a file XLibur produces is indistinguishable from one Excel produces.
    /// </summary>
    private const int WriteKeyBits = 256;

    private const int WriteSaltSize = 16;

    private const int WriteBlockSize = 16;

    private const int WriteSpinCount = 100_000;

    private const OfficeHashAlgorithm WriteHashAlgorithm = OfficeHashAlgorithm.Sha512;

    /// <summary>
    /// Decrypts an encrypted workbook into a stream positioned at the start of the .xlsx it carries.
    /// </summary>
    /// <exception cref="XLInvalidPasswordException">
    /// <paramref name="password"/> is null, empty or wrong.
    /// </exception>
    /// <exception cref="XLEncryptionException">
    /// The container or descriptor is malformed, uses something outside the profile Excel writes, or
    /// fails its integrity check.
    /// </exception>
    public static MemoryStream Decrypt(Stream source, string? password)
    {
        if (string.IsNullOrEmpty(password))
        {
            throw new XLInvalidPasswordException(
                "The workbook is encrypted. Supply the password through LoadOptions.Password to open it.");
        }

        var (encryptionInfo, encryptedPackage) = EncryptedPackageContainer.ReadStreams(source);
        var package = DecryptPackage(encryptionInfo, encryptedPackage, password);

        // Over the decrypted bytes rather than a copy of them: the workbook keeps this as the base
        // package a save copies out of, and only ever reads it. A save of an encrypted workbook
        // builds a fresh package to patch, because the compound file it produces has to be written
        // whole rather than patched in place.
        return new MemoryStream(package);
    }

    private static byte[] DecryptPackage(byte[] encryptionInfo, byte[] encryptedPackage, string password)
    {
        if (encryptionInfo.Length < 4)
            throw new XLEncryptionException("The EncryptionInfo stream is too short to carry a version.");

        // The version pair selects the scheme: 4.4 is agile, 3.2 and 4.2 are standard. 2.x is the
        // RC4 era, which XLibur does not read.
        var major = BitConverter.ToUInt16(encryptionInfo, 0);
        var minor = BitConverter.ToUInt16(encryptionInfo, 2);

        return (major, minor) switch
        {
            (4, 4) => DecryptAgile(encryptionInfo, encryptedPackage, password),
            (3, 2) or (4, 2) => DecryptStandard(encryptionInfo, encryptedPackage, password),
            (2, 2) => throw new XLEncryptionException(
                "The workbook uses RC4 encryption, which XLibur does not support. Re-save it from Excel to upgrade the encryption."),
            _ => throw new XLEncryptionException(
                $"Unsupported encryption version {major}.{minor}."),
        };
    }

    private static byte[] DecryptAgile(byte[] encryptionInfo, byte[] encryptedPackage, string password)
    {
        var descriptor = AgileEncryptionDescriptor.Parse(encryptionInfo);
        var packageKey = AgileCrypto.DecryptPackageKey(descriptor, password);

        try
        {
            // Integrity first: a tampered package should be reported as tampered rather than handed
            // to the zip reader to fail as a corrupt archive.
            AgileCrypto.VerifyIntegrity(descriptor, packageKey, encryptedPackage);
            return AgileCrypto.DecryptPackage(descriptor, packageKey, encryptedPackage);
        }
        finally
        {
            CryptographicOperations.ZeroMemory(packageKey);
        }
    }

    private static byte[] DecryptStandard(byte[] encryptionInfo, byte[] encryptedPackage, string password)
    {
        var descriptor = StandardEncryptionDescriptor.Parse(encryptionInfo);
        var key = descriptor.DeriveAndVerifyKey(password);

        try
        {
            // Standard encryption has no integrity check to run; the verifier is all it offers.
            return StandardEncryptionDescriptor.DecryptPackage(key, encryptedPackage);
        }
        finally
        {
            CryptographicOperations.ZeroMemory(key);
        }
    }

    /// <summary>
    /// Encrypts .xlsx bytes into a compound file written to <paramref name="destination"/>, using
    /// agile encryption with the parameters Excel writes.
    /// </summary>
    public static void Encrypt(Stream destination, byte[] package, string password)
    {
        if (string.IsNullOrEmpty(password))
            throw new ArgumentException("A password is required to encrypt a workbook.", nameof(password));

        var keyDataSalt = new byte[WriteSaltSize];
        var keyEncryptorSalt = new byte[WriteSaltSize];
        var packageKey = new byte[WriteKeyBits / 8];
        RandomNumberGenerator.Fill(keyDataSalt);
        RandomNumberGenerator.Fill(keyEncryptorSalt);
        RandomNumberGenerator.Fill(packageKey);

        try
        {
            var keyData = NewCipherParameters(keyDataSalt);
            var keyEncryptorParameters = NewCipherParameters(keyEncryptorSalt);

            var (verifierHashInput, verifierHashValue, encryptedKeyValue) =
                AgileCrypto.CreateVerifier(keyEncryptorParameters, WriteSpinCount, password, packageKey);

            var descriptorWithoutIntegrity = new AgileEncryptionDescriptor
            {
                KeyData = keyData,
                PasswordKeyEncryptor = new AgilePasswordKeyEncryptor
                {
                    Parameters = keyEncryptorParameters,
                    SpinCount = WriteSpinCount,
                    EncryptedVerifierHashInput = verifierHashInput,
                    EncryptedVerifierHashValue = verifierHashValue,
                    EncryptedKeyValue = encryptedKeyValue,
                },
            };

            var encryptedPackage = AgileCrypto.EncryptPackage(descriptorWithoutIntegrity, packageKey, package);

            // The HMAC covers the encrypted package, so it can only be computed once that exists.
            var (encryptedHmacKey, encryptedHmacValue) =
                AgileCrypto.CreateIntegrity(keyData, packageKey, encryptedPackage);

            var descriptor = new AgileEncryptionDescriptor
            {
                KeyData = keyData,
                PasswordKeyEncryptor = descriptorWithoutIntegrity.PasswordKeyEncryptor,
                EncryptedHmacKey = encryptedHmacKey,
                EncryptedHmacValue = encryptedHmacValue,
            };

            EncryptedPackageContainer.WriteStreams(destination, descriptor.ToEncryptionInfo(), encryptedPackage);
        }
        finally
        {
            CryptographicOperations.ZeroMemory(packageKey);
        }
    }

    private static AgileCipherParameters NewCipherParameters(byte[] salt) => new()
    {
        SaltSize = WriteSaltSize,
        BlockSize = WriteBlockSize,
        KeyBits = WriteKeyBits,
        HashSize = WriteHashAlgorithm.GetHashSize(),
        ChainingMode = CipherMode.CBC,
        HashAlgorithm = WriteHashAlgorithm,
        SaltValue = salt,
    };
}
