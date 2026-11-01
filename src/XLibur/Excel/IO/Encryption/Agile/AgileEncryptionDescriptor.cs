using System;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using XLibur.Excel.Exceptions;

namespace XLibur.Excel.IO.Encryption.Agile;

/// <summary>
/// The parameters shared by the package cipher and the password key encryptor. Agile encryption
/// repeats the same attribute set in both places and they are allowed to differ, so each is parsed
/// into its own instance rather than assumed equal.
/// </summary>
internal sealed class AgileCipherParameters
{
    public required int SaltSize { get; init; }

    public required int BlockSize { get; init; }

    public required int KeyBits { get; init; }

    public required int HashSize { get; init; }

    public required CipherMode ChainingMode { get; init; }

    public required OfficeHashAlgorithm HashAlgorithm { get; init; }

    public required byte[] SaltValue { get; init; }

    public int KeyBytes => KeyBits / 8;

    public static AgileCipherParameters Parse(XElement element)
    {
        OfficeCryptoAlgorithms.ValidateCipherIsAes(Attribute(element, "cipherAlgorithm"));
        var keyBits = ParseInt(element, "keyBits");
        OfficeCryptoAlgorithms.ValidateKeyBits(keyBits);

        var saltValue = ParseBase64(element, "saltValue");
        var saltSize = ParseInt(element, "saltSize");
        if (saltValue.Length != saltSize)
        {
            throw new XLEncryptionException(
                $"The encryption descriptor declares a salt of {saltSize} bytes but carries {saltValue.Length}.");
        }

        var hashAlgorithm = OfficeCryptoAlgorithms.ParseHashAlgorithm(Attribute(element, "hashAlgorithm"));
        var hashSize = ParseInt(element, "hashSize");
        if (hashSize != hashAlgorithm.GetHashSize())
        {
            throw new XLEncryptionException(
                $"The encryption descriptor declares a hash size of {hashSize} bytes, which does not match {hashAlgorithm.ToXmlName()}.");
        }

        return new AgileCipherParameters
        {
            SaltSize = saltSize,
            BlockSize = ParseInt(element, "blockSize"),
            KeyBits = keyBits,
            HashSize = hashSize,
            ChainingMode = OfficeCryptoAlgorithms.ParseChainingMode(Attribute(element, "cipherChaining")),
            HashAlgorithm = hashAlgorithm,
            SaltValue = saltValue,
        };
    }

    internal static string? Attribute(XElement element, string name) => element.Attribute(name)?.Value;

    internal static int ParseInt(XElement element, string name)
    {
        var raw = Attribute(element, name)
                  ?? throw new XLEncryptionException($"The encryption descriptor is missing the '{name}' attribute.");

        return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? value
            : throw new XLEncryptionException($"The encryption descriptor has a non-numeric '{name}' of '{raw}'.");
    }

    internal static byte[] ParseBase64(XElement element, string name)
    {
        var raw = Attribute(element, name)
                  ?? throw new XLEncryptionException($"The encryption descriptor is missing the '{name}' attribute.");

        try
        {
            return Convert.FromBase64String(raw);
        }
        catch (FormatException e)
        {
            throw new XLEncryptionException($"The encryption descriptor has a malformed base64 '{name}'.", e);
        }
    }
}

/// <summary>
/// The password key encryptor: the cipher parameters plus the material that turns a password into
/// the key the package is actually encrypted with.
/// </summary>
internal sealed class AgilePasswordKeyEncryptor
{
    public required AgileCipherParameters Parameters { get; init; }

    public required int SpinCount { get; init; }

    public required byte[] EncryptedVerifierHashInput { get; init; }

    public required byte[] EncryptedVerifierHashValue { get; init; }

    public required byte[] EncryptedKeyValue { get; init; }
}

/// <summary>
/// A parsed agile <c>EncryptionInfo</c> stream (version 4.4): an 8 byte header followed by the
/// <c>&lt;encryption&gt;</c> XML descriptor.
/// </summary>
/// <remarks>
/// [MS-OFFCRYPTO] 2.3.4.10.
/// </remarks>
internal sealed class AgileEncryptionDescriptor
{
    internal static readonly XNamespace EncryptionNs = "http://schemas.microsoft.com/office/2006/encryption";

    internal static readonly XNamespace PasswordNs = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

    private const string PasswordKeyEncryptorUri = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

    /// <summary>Cipher parameters for the encrypted package itself.</summary>
    public required AgileCipherParameters KeyData { get; init; }

    public required AgilePasswordKeyEncryptor PasswordKeyEncryptor { get; init; }

    /// <summary>
    /// Encrypted HMAC key for the integrity check, absent in files written without one.
    /// </summary>
    public byte[]? EncryptedHmacKey { get; init; }

    public byte[]? EncryptedHmacValue { get; init; }

    public static AgileEncryptionDescriptor Parse(ReadOnlySpan<byte> encryptionInfo)
    {
        // 2 bytes major version, 2 bytes minor version, 4 bytes reserved flags, then UTF-8 XML.
        const int headerSize = 8;
        if (encryptionInfo.Length <= headerSize)
            throw new XLEncryptionException("The EncryptionInfo stream is too short to be an agile descriptor.");

        var xml = Encoding.UTF8.GetString(encryptionInfo[headerSize..]);
        XDocument document;
        try
        {
            document = XDocument.Parse(xml);
        }
        catch (System.Xml.XmlException e)
        {
            throw new XLEncryptionException("The agile encryption descriptor is not well-formed XML.", e);
        }

        var root = document.Root
                   ?? throw new XLEncryptionException("The agile encryption descriptor is empty.");

        var keyDataElement = root.Element(EncryptionNs + "keyData")
                             ?? throw new XLEncryptionException("The agile encryption descriptor has no keyData element.");

        var keyEncryptor = FindPasswordKeyEncryptor(root);
        var dataIntegrity = root.Element(EncryptionNs + "dataIntegrity");

        return new AgileEncryptionDescriptor
        {
            KeyData = AgileCipherParameters.Parse(keyDataElement),
            PasswordKeyEncryptor = new AgilePasswordKeyEncryptor
            {
                Parameters = AgileCipherParameters.Parse(keyEncryptor),
                SpinCount = AgileCipherParameters.ParseInt(keyEncryptor, "spinCount"),
                EncryptedVerifierHashInput = AgileCipherParameters.ParseBase64(keyEncryptor, "encryptedVerifierHashInput"),
                EncryptedVerifierHashValue = AgileCipherParameters.ParseBase64(keyEncryptor, "encryptedVerifierHashValue"),
                EncryptedKeyValue = AgileCipherParameters.ParseBase64(keyEncryptor, "encryptedKeyValue"),
            },
            EncryptedHmacKey = dataIntegrity is null
                ? null
                : AgileCipherParameters.ParseBase64(dataIntegrity, "encryptedHmacKey"),
            EncryptedHmacValue = dataIntegrity is null
                ? null
                : AgileCipherParameters.ParseBase64(dataIntegrity, "encryptedHmacValue"),
        };
    }

    private static XElement FindPasswordKeyEncryptor(XElement root)
    {
        var keyEncryptors = root.Element(EncryptionNs + "keyEncryptors")
                            ?? throw new XLEncryptionException("The agile encryption descriptor has no keyEncryptors element.");

        foreach (var keyEncryptor in keyEncryptors.Elements(EncryptionNs + "keyEncryptor"))
        {
            if (keyEncryptor.Attribute("uri")?.Value != PasswordKeyEncryptorUri)
                continue;

            var encryptedKey = keyEncryptor.Element(PasswordNs + "encryptedKey");
            if (encryptedKey is not null)
                return encryptedKey;
        }

        // Certificate key encryptors are the other kind the specification defines. A file using one
        // is well-formed and simply not something a password can open.
        throw new XLEncryptionException(
            "The workbook is encrypted, but not with a password. XLibur supports password key encryptors only.");
    }

    /// <summary>
    /// Renders the descriptor back to a complete <c>EncryptionInfo</c> stream, header included.
    /// </summary>
    public byte[] ToEncryptionInfo()
    {
        var keyData = new XElement(EncryptionNs + "keyData");
        WriteCipherAttributes(keyData, KeyData);

        var encryptedKey = new XElement(PasswordNs + "encryptedKey",
            new XAttribute("spinCount", PasswordKeyEncryptor.SpinCount.ToString(CultureInfo.InvariantCulture)));
        WriteCipherAttributes(encryptedKey, PasswordKeyEncryptor.Parameters);
        encryptedKey.SetAttributeValue("encryptedVerifierHashInput", Convert.ToBase64String(PasswordKeyEncryptor.EncryptedVerifierHashInput));
        encryptedKey.SetAttributeValue("encryptedVerifierHashValue", Convert.ToBase64String(PasswordKeyEncryptor.EncryptedVerifierHashValue));
        encryptedKey.SetAttributeValue("encryptedKeyValue", Convert.ToBase64String(PasswordKeyEncryptor.EncryptedKeyValue));

        var root = new XElement(EncryptionNs + "encryption",
            new XAttribute(XNamespace.Xmlns + "p", PasswordNs.NamespaceName),
            keyData);

        if (EncryptedHmacKey is not null && EncryptedHmacValue is not null)
        {
            root.Add(new XElement(EncryptionNs + "dataIntegrity",
                new XAttribute("encryptedHmacKey", Convert.ToBase64String(EncryptedHmacKey)),
                new XAttribute("encryptedHmacValue", Convert.ToBase64String(EncryptedHmacValue))));
        }

        root.Add(new XElement(EncryptionNs + "keyEncryptors",
            new XElement(EncryptionNs + "keyEncryptor",
                new XAttribute("uri", PasswordKeyEncryptorUri),
                encryptedKey)));

        // Written by hand rather than through XDocument.ToString, which drops the declaration, and
        // fully qualified because the unqualified SaveOptions here would bind to XLibur's own.
        const string declaration = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
        var xml = Encoding.UTF8.GetBytes(
            declaration + root.ToString(System.Xml.Linq.SaveOptions.DisableFormatting));

        var stream = new byte[8 + xml.Length];
        // Version 4.4 with the reserved flag bit set, the combination that marks agile encryption.
        BitConverter.TryWriteBytes(stream.AsSpan(0), (ushort)4);
        BitConverter.TryWriteBytes(stream.AsSpan(2), (ushort)4);
        BitConverter.TryWriteBytes(stream.AsSpan(4), 0x40u);
        xml.CopyTo(stream.AsSpan(8));
        return stream;
    }

    private static void WriteCipherAttributes(XElement element, AgileCipherParameters parameters)
    {
        element.SetAttributeValue("saltSize", parameters.SaltSize.ToString(CultureInfo.InvariantCulture));
        element.SetAttributeValue("blockSize", parameters.BlockSize.ToString(CultureInfo.InvariantCulture));
        element.SetAttributeValue("keyBits", parameters.KeyBits.ToString(CultureInfo.InvariantCulture));
        element.SetAttributeValue("hashSize", parameters.HashSize.ToString(CultureInfo.InvariantCulture));
        element.SetAttributeValue("cipherAlgorithm", "AES");
        element.SetAttributeValue("cipherChaining", "ChainingModeCBC");
        element.SetAttributeValue("hashAlgorithm", parameters.HashAlgorithm.ToXmlName());
        element.SetAttributeValue("saltValue", Convert.ToBase64String(parameters.SaltValue));
    }
}
