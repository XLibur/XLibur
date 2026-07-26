using System;
using System.IO;
using System.Threading.Tasks;
using OpenMcdf;
using XLibur.Excel;
using XLibur.Excel.Exceptions;

namespace XLibur.Tests.Excel.Encryption;

/// <summary>
/// Tests against workbooks encrypted by other applications. These are the only tests that show
/// XLibur agrees with the rest of the world: a round trip through XLibur's own encrypt and decrypt
/// would still pass if both sides shared a mistake.
/// </summary>
/// <remarks>
/// See <c>XLibur.Tests/Resource/Encrypted/README.md</c> for the corpus and how to extend it.
/// </remarks>
public class RealWorldEncryptedFileTests
{
    private const string ResourcePath = @"Encrypted\Encrypted_XL365.xlsx";

    private const string Password = "Password";

    private static Stream OpenResource() =>
        TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(ResourcePath));

    [Test]
    public async Task Excel365_agile_encrypted_workbook_opens_with_its_password()
    {
        using var stream = OpenResource();
        using var wb = new XLWorkbook(stream, new LoadOptions { Password = Password });

        await Assert.That(wb.Worksheets.Count).IsEqualTo(1);
        await Assert.That(wb.Worksheet("Sheet1").Cell("A1").GetString()).IsEqualTo("PROTECTED FILE");
    }

    [Test]
    public async Task Excel365_encrypted_workbook_uses_agile_encryption()
    {
        // Pins which code path the file above actually exercises. Without this, a change that broke
        // agile parsing but left the file readable some other way would look like a passing test.
        using var stream = OpenResource();
        using var storage = RootStorage.Open(stream, StorageModeFlags.LeaveOpen);
        using var encryptionInfo = storage.OpenStream("EncryptionInfo");

        var header = new byte[4];
        _ = encryptionInfo.Read(header, 0, header.Length);

        await Assert.That(BitConverter.ToUInt16(header, 0)).IsEqualTo((ushort)4).Because("major version");
        await Assert.That(BitConverter.ToUInt16(header, 2)).IsEqualTo((ushort)4).Because("minor version");
    }

    [Test]
    public async Task Excel365_encrypted_workbook_rejects_a_wrong_password()
    {
        using var stream = OpenResource();
        await Assert.That(() => new XLWorkbook(stream, new LoadOptions { Password = "wrong" }))
            .Throws<XLInvalidPasswordException>();
    }

    [Test]
    public async Task Excel365_encrypted_workbook_rejects_a_missing_password()
    {
        using var stream = OpenResource();
        await Assert.That(() => new XLWorkbook(stream)).Throws<XLInvalidPasswordException>();
    }

    [Test]
    public async Task Excel365_encrypted_workbook_survives_a_decrypt_edit_re_encrypt_cycle()
    {
        using var reEncrypted = new MemoryStream();

        using (var stream = OpenResource())
        using (var wb = new XLWorkbook(stream, new LoadOptions { Password = Password }))
        {
            wb.Worksheet("Sheet1").Cell("A2").Value = "added by XLibur";
            wb.SaveAs(reEncrypted, new SaveOptions { Password = "a new password" });
        }

        reEncrypted.Position = 0;

        using var reopened = new XLWorkbook(reEncrypted, new LoadOptions { Password = "a new password" });
        await Assert.That(reopened.Worksheet("Sheet1").Cell("A1").GetString()).IsEqualTo("PROTECTED FILE");
        await Assert.That(reopened.Worksheet("Sheet1").Cell("A2").GetString()).IsEqualTo("added by XLibur");
    }

    [Test]
    public async Task Excel365_encrypted_workbook_can_be_saved_unencrypted()
    {
        using var plain = new MemoryStream();

        using (var stream = OpenResource())
        using (var wb = new XLWorkbook(stream, new LoadOptions { Password = Password }))
        {
            wb.SaveAs(plain);
        }

        plain.Position = 0;

        // The password does not follow the workbook to the save, so this is an ordinary package.
        using var reopened = new XLWorkbook(plain);
        await Assert.That(reopened.Worksheet("Sheet1").Cell("A1").GetString()).IsEqualTo("PROTECTED FILE");
    }
}
