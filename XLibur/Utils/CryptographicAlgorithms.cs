using System;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using static XLibur.Excel.XLProtectionAlgorithm;

namespace XLibur.Utils;

internal static class CryptographicAlgorithms
{
    public static string Base64Decode(string base64EncodedData)
    {
        var base64EncodedBytes = Convert.FromBase64String(base64EncodedData);
        return Encoding.UTF8.GetString(base64EncodedBytes);
    }

    public static string Base64Encode(string plainText)
    {
        var plainTextBytes = Encoding.UTF8.GetBytes(plainText);
        return Convert.ToBase64String(plainTextBytes);
    }

    public static string GenerateNewSalt(Algorithm algorithm)
    {
        return RequiresSalt(algorithm) ? GetSalt() : string.Empty;
    }

    public static string GetPasswordHash(Algorithm algorithm, string password, string salt = "", uint spinCount = 0)
    {
        ArgumentNullException.ThrowIfNull(password);
        ArgumentNullException.ThrowIfNull(salt);

        if (password.Length == 0) return "";

        switch (algorithm)
        {
            case Algorithm.SimpleHash:
                return GetDefaultPasswordHash(password);

            case Algorithm.SHA512:
                return GetSha512PasswordHash(password, salt, spinCount);

            default:
                return string.Empty;
        }
    }

    public static string GetSalt(int length = 32)
    {
        var salt = new byte[length];
        RandomNumberGenerator.Fill(salt);
        // Ensure no zero bytes (matching previous RNGCryptoServiceProvider.GetNonZeroBytes behavior)
        Span<byte> singleByte = stackalloc byte[1];
        for (int i = 0; i < salt.Length; i++)
        {
            while (salt[i] == 0)
            {
                RandomNumberGenerator.Fill(singleByte);
                salt[i] = singleByte[0];
            }
        }
        return Convert.ToBase64String(salt);
    }

    public static bool RequiresSalt(Algorithm algorithm)
    {
        switch (algorithm)
        {
            case Algorithm.SimpleHash:
                return false;

            case Algorithm.SHA512:
                return true;

            default:
                return false;
        }
    }

    private static string GetDefaultPasswordHash(string password)
    {
        ArgumentNullException.ThrowIfNull(password);

        // http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/
        // http://sc.openoffice.org/excelfileformat.pdf - 4.18.4
        // http://web.archive.org/web/20080906232341/http://blogs.infosupport.com/wouterv/archive/2006/11/21/Hashing-password-for-use-in-SpreadsheetML.aspx
        byte[] passwordCharacters = Encoding.ASCII.GetBytes(password);
        int hash = 0;
        if (passwordCharacters.Length > 0)
        {
            int charIndex = passwordCharacters.Length;

            while (charIndex-- > 0)
            {
                hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
                hash ^= passwordCharacters[charIndex];
            }
            // Main difference from spec, also hash with charcount
            hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
            hash ^= passwordCharacters.Length;
            hash ^= (0x8000 | ('N' << 8) | 'K');
        }

        return Convert.ToString(hash, 16).ToUpperInvariant();
    }

    private static string GetSha512PasswordHash(string password, string salt, uint spinCount)
    {
        ArgumentNullException.ThrowIfNull(password);
        ArgumentNullException.ThrowIfNull(salt);

        var saltBytes = Convert.FromBase64String(salt);
        var passwordBytes = Encoding.Unicode.GetBytes(password);
        var bytes = saltBytes.Concat(passwordBytes).ToArray();

        var hashedBytes = SHA512.HashData(bytes);

        bytes = new byte[hashedBytes.Length + sizeof(uint)];
        for (uint i = 0; i < spinCount; i++)
        {
            var le = BitConverter.GetBytes(i);
            Array.Copy(hashedBytes, bytes, hashedBytes.Length);
            Array.Copy(le, 0, bytes, hashedBytes.Length, le.Length);
            hashedBytes = SHA512.HashData(bytes);
        }

        return Convert.ToBase64String(hashedBytes);
    }
}
