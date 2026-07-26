using System.IO;
using System.Threading.Tasks;
using OpenMcdf;
using XLibur.Excel;
using XLibur.Excel.Exceptions;

namespace XLibur.Tests.Excel.Encryption;

public class WorkbookEncryptionTests
{
    private const string Password = "correct horse battery staple";

    /// <summary>
    /// Builds a small workbook and saves it encrypted, returning the compound file bytes.
    /// </summary>
    private static byte[] CreateEncryptedWorkbook(string password, string cellValue = "Hello encrypted world")
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = cellValue;
            ws.Cell("A2").Value = 42;
            ws.Cell("A3").FormulaA1 = "A2*2";

            wb.SaveAs(ms, new SaveOptions { Password = password });
        }

        return ms.ToArray();
    }

    [Test]
    public async Task Encrypted_workbook_round_trips_through_a_password()
    {
        var encrypted = CreateEncryptedWorkbook(Password);

        using var ms = new MemoryStream(encrypted);
        using var wb = new XLWorkbook(ms, new LoadOptions { Password = Password });
        var ws = wb.Worksheet("Data");

        await Assert.That(ws.Cell("A1").GetString()).IsEqualTo("Hello encrypted world");
        await Assert.That(ws.Cell("A2").GetDouble()).IsEqualTo(42);
        await Assert.That(ws.Cell("A3").FormulaA1).IsEqualTo("A2*2");
    }

    [Test]
    public async Task Encrypted_workbook_is_a_compound_file_with_the_two_expected_streams()
    {
        var encrypted = CreateEncryptedWorkbook(Password);

        // The container has to look the way Excel expects before anything inside it can matter.
        await Assert.That(encrypted[0]).IsEqualTo((byte)0xD0);
        await Assert.That(encrypted[1]).IsEqualTo((byte)0xCF);
        await Assert.That(encrypted[2]).IsEqualTo((byte)0x11);
        await Assert.That(encrypted[3]).IsEqualTo((byte)0xE0);

        using var ms = new MemoryStream(encrypted);
        using var storage = RootStorage.Open(ms, StorageModeFlags.LeaveOpen);

        await Assert.That(storage.ContainsEntry("EncryptionInfo")).IsTrue();
        await Assert.That(storage.ContainsEntry("EncryptedPackage")).IsTrue();
    }

    [Test]
    public async Task Encrypted_workbook_does_not_leak_its_content_in_the_clear()
    {
        var encrypted = CreateEncryptedWorkbook(Password, "TotallySecretValue");

        // A zip stored rather than encrypted would leave the string and the local file header
        // visible. Neither may appear anywhere in the container.
        var asText = System.Text.Encoding.ASCII.GetString(encrypted);
        await Assert.That(asText).DoesNotContain("TotallySecretValue");
        await Assert.That(asText).DoesNotContain("xl/workbook.xml");
    }

    [Test]
    public async Task Wrong_password_throws_invalid_password()
    {
        var encrypted = CreateEncryptedWorkbook(Password);

        using var ms = new MemoryStream(encrypted);
        await Assert.That(() => new XLWorkbook(ms, new LoadOptions { Password = "not the password" }))
            .Throws<XLInvalidPasswordException>();
    }

    [Test]
    public async Task Missing_password_throws_invalid_password()
    {
        var encrypted = CreateEncryptedWorkbook(Password);

        using var ms = new MemoryStream(encrypted);
        await Assert.That(() => new XLWorkbook(ms)).Throws<XLInvalidPasswordException>();
    }

    [Test]
    public async Task Tampered_package_fails_the_integrity_check_rather_than_returning_garbage()
    {
        var encrypted = CreateEncryptedWorkbook(Password);

        using var ms = new MemoryStream();
        ms.Write(encrypted, 0, encrypted.Length);
        ms.Position = 0;

        // Flip one bit in the middle of the ciphertext, the tampering the spec calls out.
        using (var storage = RootStorage.Open(ms, StorageModeFlags.LeaveOpen))
        {
            using var packageStream = storage.OpenStream("EncryptedPackage");
            var package = new byte[packageStream.Length];
            _ = packageStream.Read(package, 0, package.Length);

            package[package.Length / 2] ^= 0xFF;

            packageStream.Position = 0;
            packageStream.Write(package, 0, package.Length);
        }

        ms.Position = 0;
        await Assert.That(() => new XLWorkbook(ms, new LoadOptions { Password = Password }))
            .Throws<XLEncryptionException>();
    }

    [Test]
    [Arguments("a")]
    [Arguments("pässwörd with ünicode ☕")]
    [Arguments("0123456789012345678901234567890123456789012345678901234567890123456789")]
    [Arguments("   leading and trailing spaces   ")]
    public async Task Passwords_of_awkward_shapes_round_trip(string password)
    {
        var encrypted = CreateEncryptedWorkbook(password);

        using var ms = new MemoryStream(encrypted);
        using var wb = new XLWorkbook(ms, new LoadOptions { Password = password });

        await Assert.That(wb.Worksheet("Data").Cell("A1").GetString()).IsEqualTo("Hello encrypted world");
    }

    [Test]
    public async Task Password_on_load_is_ignored_for_an_unencrypted_workbook()
    {
        using var plain = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("Data").Cell("A1").Value = "not a secret";
            wb.SaveAs(plain);
        }

        plain.Position = 0;

        // Supplying a password for a file that turns out not to need one is not an error: the
        // caller often cannot know which kind of file they were handed.
        using var wb2 = new XLWorkbook(plain, new LoadOptions { Password = Password });
        await Assert.That(wb2.Worksheet("Data").Cell("A1").GetString()).IsEqualTo("not a secret");
    }

    [Test]
    public async Task Saving_without_a_password_after_loading_with_one_produces_a_plain_workbook()
    {
        var encrypted = CreateEncryptedWorkbook(Password);

        using var source = new MemoryStream(encrypted);
        using var wb = new XLWorkbook(source, new LoadOptions { Password = Password });
        wb.Worksheet("Data").Cell("A4").Value = "added after decryption";

        using var resaved = new MemoryStream();
        wb.SaveAs(resaved);
        resaved.Position = 0;

        // A password is never carried over implicitly, so this is an ordinary package that opens
        // with no password at all.
        await Assert.That(resaved.ToArray()[0]).IsEqualTo((byte)'P');

        using var reopened = new XLWorkbook(resaved);
        await Assert.That(reopened.Worksheet("Data").Cell("A4").GetString()).IsEqualTo("added after decryption");
    }

    [Test]
    public async Task A_decrypted_workbook_can_be_re_encrypted_with_a_different_password()
    {
        var encrypted = CreateEncryptedWorkbook(Password);

        using var source = new MemoryStream(encrypted);
        using var wb = new XLWorkbook(source, new LoadOptions { Password = Password });
        wb.Worksheet("Data").Cell("A5").Value = "round two";

        using var reEncrypted = new MemoryStream();
        wb.SaveAs(reEncrypted, new SaveOptions { Password = "a different password" });
        reEncrypted.Position = 0;

        using var reopened = new XLWorkbook(reEncrypted, new LoadOptions { Password = "a different password" });
        await Assert.That(reopened.Worksheet("Data").Cell("A5").GetString()).IsEqualTo("round two");
        await Assert.That(reopened.Worksheet("Data").Cell("A1").GetString()).IsEqualTo("Hello encrypted world");
    }

    [Test]
    public async Task Encrypted_save_and_load_work_through_a_file_on_disk()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlibur-encrypted-{System.Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = new XLWorkbook())
            {
                wb.AddWorksheet("Data").Cell("A1").Value = "from disk";
                wb.SaveAs(path, new SaveOptions { Password = Password });
            }

            using var reopened = new XLWorkbook(path, new LoadOptions { Password = Password });
            await Assert.That(reopened.Worksheet("Data").Cell("A1").GetString()).IsEqualTo("from disk");
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    [Test]
    public async Task A_plain_compound_file_that_is_not_an_encrypted_workbook_is_reported_as_such()
    {
        using var ms = new MemoryStream();
        using (var storage = RootStorage.Create(ms, OpenMcdf.Version.V4, StorageModeFlags.LeaveOpen))
        {
            using var stream = storage.CreateStream("Workbook");
            stream.Write([1, 2, 3], 0, 3);
        }

        ms.Position = 0;

        // A legacy .xls is a compound file too. It must not be mistaken for an encrypted workbook
        // with a bad password, which would send the caller off to re-prompt for one.
        await Assert.That(() => new XLWorkbook(ms, new LoadOptions { Password = Password }))
            .Throws<XLEncryptionException>();
    }
}
