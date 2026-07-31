using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OpenMcdf;
using XLibur.Excel;
using XLibur.Excel.Exceptions;
using XLibur.Excel.IO.Encryption;
using XLibur.Tests.Utils;

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

    /// <summary>
    /// Writes a small encrypted workbook to <paramref name="path"/>, as the file the round-trip
    /// tests below open and save back.
    /// </summary>
    private static void WriteEncryptedFile(string path, string password, string cellValue = "before")
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Data").Cell("A1").Value = cellValue;
        wb.SaveAs(path, new SaveOptions { Password = password });
    }

    /// <summary>
    /// Whether the bytes are an encrypted container rather than an ordinary package. The two are
    /// told apart by their first byte: a compound file starts 0xD0, a zip starts 'P'.
    /// </summary>
    private static bool IsEncryptedContainer(byte[] bytes) => bytes.Length > 0 && bytes[0] == 0xD0;

    [Test]
    public async Task Save_writes_an_encrypted_file_back_encrypted_with_the_password_it_was_opened_with()
    {
        using var file = new TemporaryFile();
        WriteEncryptedFile(file.Path, Password);

        using (var wb = new XLWorkbook(file.Path, new LoadOptions { Password = Password }))
        {
            wb.Worksheet("Data").Cell("A1").Value = "after";
            wb.Save();
        }

        // The edit reached the file, and the file is still encrypted under the same password.
        await Assert.That(IsEncryptedContainer(File.ReadAllBytes(file.Path))).IsTrue();

        using var reopened = new XLWorkbook(file.Path, new LoadOptions { Password = Password });
        await Assert.That(reopened.Worksheet("Data").Cell("A1").GetString()).IsEqualTo("after");
    }

    [Test]
    public async Task Save_repeated_on_an_encrypted_file_keeps_working()
    {
        using var file = new TemporaryFile();
        WriteEncryptedFile(file.Path, Password);

        using (var wb = new XLWorkbook(file.Path, new LoadOptions { Password = Password }))
        {
            wb.Worksheet("Data").Cell("A1").Value = "first";
            wb.Save();

            wb.Worksheet("Data").Cell("A1").Value = "second";
            wb.Save();
        }

        using var reopened = new XLWorkbook(file.Path, new LoadOptions { Password = Password });
        await Assert.That(reopened.Worksheet("Data").Cell("A1").GetString()).IsEqualTo("second");
    }

    [Test]
    public async Task Save_with_a_password_rotates_the_password_in_place()
    {
        using var file = new TemporaryFile();
        WriteEncryptedFile(file.Path, Password);

        using (var wb = new XLWorkbook(file.Path, new LoadOptions { Password = Password }))
            wb.Save(new SaveOptions { Password = "n3wsecret" });

        using var reopened = new XLWorkbook(file.Path, new LoadOptions { Password = "n3wsecret" });
        await Assert.That(reopened.Worksheet("Data").Cell("A1").GetString()).IsEqualTo("before");

        // The password it was opened with no longer opens it.
        await Assert.That(() => new XLWorkbook(file.Path, new LoadOptions { Password = Password }))
            .Throws<XLInvalidPasswordException>();
    }

    [Test]
    public async Task Save_with_a_password_encrypts_a_workbook_that_was_loaded_from_a_plain_file()
    {
        using var file = new TemporaryFile();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("Data").Cell("A1").Value = "was plain";
            wb.SaveAs(file.Path);
        }

        using (var wb = new XLWorkbook(file.Path))
        {
            wb.Worksheet("Data").Cell("A2").Value = "now secret";
            wb.Save(new SaveOptions { Password = Password });
        }

        await Assert.That(IsEncryptedContainer(File.ReadAllBytes(file.Path))).IsTrue();

        using var reopened = new XLWorkbook(file.Path, new LoadOptions { Password = Password });
        await Assert.That(reopened.Worksheet("Data").Cell("A1").GetString()).IsEqualTo("was plain");
        await Assert.That(reopened.Worksheet("Data").Cell("A2").GetString()).IsEqualTo("now secret");
    }

    [Test]
    public async Task Save_on_a_workbook_loaded_from_a_plain_file_leaves_it_plain()
    {
        using var file = new TemporaryFile();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("Data").Cell("A1").Value = "before";
            wb.SaveAs(file.Path);
        }

        using (var wb = new XLWorkbook(file.Path))
        {
            wb.Worksheet("Data").Cell("A1").Value = "after";
            wb.Save();
        }

        await Assert.That(IsEncryptedContainer(File.ReadAllBytes(file.Path))).IsFalse();

        using var reopened = new XLWorkbook(file.Path);
        await Assert.That(reopened.Worksheet("Data").Cell("A1").GetString()).IsEqualTo("after");
    }

    [Test]
    public async Task Save_writes_the_encrypted_container_back_to_the_stream_it_was_loaded_from()
    {
        using var origin = new MemoryStream();
        var encrypted = CreateEncryptedWorkbook(Password);
        origin.Write(encrypted, 0, encrypted.Length);
        origin.Position = 0;

        using (var wb = new XLWorkbook(origin, new LoadOptions { Password = Password }))
        {
            wb.Worksheet("Data").Cell("A4").Value = "written back";
            wb.Save();
        }

        var saved = origin.ToArray();
        await Assert.That(IsEncryptedContainer(saved)).IsTrue();

        using var reopened = new XLWorkbook(new MemoryStream(saved), new LoadOptions { Password = Password });
        await Assert.That(reopened.Worksheet("Data").Cell("A4").GetString()).IsEqualTo("written back");
        await Assert.That(reopened.Worksheet("Data").Cell("A1").GetString()).IsEqualTo("Hello encrypted world");
    }

    [Test]
    public async Task Save_back_to_a_stream_leaves_no_tail_of_the_larger_workbook_it_replaced()
    {
        const string RowText = "row text long enough to make the package worth measuring";

        using var origin = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            for (var row = 1; row <= 2000; row++)
                ws.Cell(row, 1).Value = RowText;

            wb.SaveAs(origin, new SaveOptions { Password = Password });
        }

        var lengthBefore = origin.Length;
        origin.Position = 0;

        using (var wb = new XLWorkbook(origin, new LoadOptions { Password = Password }))
        {
            wb.Worksheet("Data").Rows(2, 2000).Delete();
            wb.Save();
        }

        // The stream has to end where the new container ends. Whatever the larger one left beyond
        // that is trailing rubbish on the file, which is why the truncation happens after the
        // write rather than before the encryption.
        await Assert.That(origin.Length).IsLessThan(lengthBefore);

        using var reopened = new XLWorkbook(new MemoryStream(origin.ToArray()), new LoadOptions { Password = Password });
        await Assert.That(reopened.Worksheet("Data").Cell(1, 1).GetString()).IsEqualTo(RowText);
        await Assert.That(reopened.Worksheet("Data").Cell(2, 1).GetString()).IsEqualTo(string.Empty);
    }

    [Test]
    public async Task Save_says_so_when_the_stream_it_was_loaded_from_cannot_be_written_back_to()
    {
        // A read-only stream can still be decrypted and read; it is only the write back that is
        // impossible, and that has to be said rather than skipped.
        using var origin = new MemoryStream(CreateEncryptedWorkbook(Password), writable: false);
        using var wb = new XLWorkbook(origin, new LoadOptions { Password = Password });

        await Assert.That(() => wb.Save()).Throws<InvalidOperationException>();
    }

    [Test]
    public async Task Save_after_an_encrypted_SaveAs_updates_the_file_that_SaveAs_wrote()
    {
        using var file = new TemporaryFile();
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Data").Cell("A1").Value = "first";
        wb.SaveAs(file.Path, new SaveOptions { Password = Password });

        // SaveAs used to leave the load source pointing at whatever the workbook was built from,
        // which made this Save a no-op.
        wb.Worksheet("Data").Cell("A1").Value = "second";
        wb.Save();

        using var reopened = new XLWorkbook(file.Path, new LoadOptions { Password = Password });
        await Assert.That(reopened.Worksheet("Data").Cell("A1").GetString()).IsEqualTo("second");
    }

    [Test]
    public async Task Save_after_an_encrypted_SaveAs_to_a_stream_updates_that_stream()
    {
        using var destination = new MemoryStream();
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Data").Cell("A1").Value = "first";
        wb.SaveAs(destination, new SaveOptions { Password = Password });

        wb.Worksheet("Data").Cell("A1").Value = "second";
        wb.Save();

        using var reopened = new XLWorkbook(
            new MemoryStream(destination.ToArray()), new LoadOptions { Password = Password });
        await Assert.That(reopened.Worksheet("Data").Cell("A1").GetString()).IsEqualTo("second");
    }

    [Test]
    public async Task SaveAs_without_a_password_over_the_same_path_removes_the_encryption()
    {
        using var file = new TemporaryFile();
        WriteEncryptedFile(file.Path, Password);

        using (var wb = new XLWorkbook(file.Path, new LoadOptions { Password = Password }))
        {
            wb.Worksheet("Data").Cell("A1").Value = "no longer secret";

            // The documented way to decrypt in place: SaveAs states the encryption of the file it
            // writes, and no password means none.
            wb.SaveAs(file.Path);
        }

        await Assert.That(IsEncryptedContainer(File.ReadAllBytes(file.Path))).IsFalse();

        using var reopened = new XLWorkbook(file.Path);
        await Assert.That(reopened.Worksheet("Data").Cell("A1").GetString()).IsEqualTo("no longer secret");
    }

    [Test]
    public async Task Save_after_a_plain_SaveAs_does_not_bring_the_encryption_back()
    {
        using var encryptedFile = new TemporaryFile();
        using var plainFile = new TemporaryFile();
        WriteEncryptedFile(encryptedFile.Path, Password);

        using (var wb = new XLWorkbook(encryptedFile.Path, new LoadOptions { Password = Password }))
        {
            wb.SaveAs(plainFile.Path);

            wb.Worksheet("Data").Cell("A1").Value = "still plain";
            wb.Save();
        }

        await Assert.That(IsEncryptedContainer(File.ReadAllBytes(plainFile.Path))).IsFalse();

        using var reopened = new XLWorkbook(plainFile.Path);
        await Assert.That(reopened.Worksheet("Data").Cell("A1").GetString()).IsEqualTo("still plain");

        // And the file it came from is untouched by any of that.
        using var original = new XLWorkbook(encryptedFile.Path, new LoadOptions { Password = Password });
        await Assert.That(original.Worksheet("Data").Cell("A1").GetString()).IsEqualTo("before");
    }

    /// <summary>
    /// Thrown by <see cref="FailingEncryptor"/>, so that a test asserting a save failed cannot be
    /// satisfied by some other failure that happens to carry a common exception type.
    /// </summary>
    private sealed class EncryptionFailure : Exception
    {
        public EncryptionFailure()
            : base("Simulated encryption failure.")
        {
        }
    }

    /// <summary>
    /// An encryption that fails before it produces anything, standing in for the failures the real
    /// one can suffer and a test cannot provoke: an out-of-memory on the two full-size allocations,
    /// or a cryptographic fault.
    /// </summary>
    private sealed class FailingEncryptor : IWorkbookEncryptor
    {
        public void Encrypt(Stream destination, byte[] package, string password) =>
            throw new EncryptionFailure();
    }

    /// <summary>
    /// Whether two sets of bytes are identical, as a single assertion rather than a diff of two
    /// buffers that run to kilobytes.
    /// </summary>
    private static async Task AssertBytesAreUnchanged(byte[] before, byte[] after)
    {
        await Assert.That(after.Length).IsEqualTo(before.Length);
        await Assert.That(after.SequenceEqual(before)).IsTrue();
    }

    [Test]
    public async Task Save_leaves_the_file_it_would_overwrite_untouched_when_the_encryption_fails()
    {
        using var file = new TemporaryFile();
        WriteEncryptedFile(file.Path, Password);
        var before = File.ReadAllBytes(file.Path);

        using var wb = new XLWorkbook(file.Path, new LoadOptions { Password = Password });
        wb.Worksheet("Data").Cell("A1").Value = "after";
        wb.Encryptor = new FailingEncryptor();

        // The whole point of encrypting into a buffer: the destination is opened for writing only
        // once the bytes that replace it exist, so a failure before that leaves the file alone.
        await Assert.That(() => wb.Save()).Throws<EncryptionFailure>();
        await AssertBytesAreUnchanged(before, File.ReadAllBytes(file.Path));
    }

    [Test]
    public async Task Save_leaves_the_stream_it_would_overwrite_untouched_when_the_encryption_fails()
    {
        using var origin = new MemoryStream();
        var encrypted = CreateEncryptedWorkbook(Password);
        origin.Write(encrypted, 0, encrypted.Length);
        origin.Position = 0;

        using var wb = new XLWorkbook(origin, new LoadOptions { Password = Password });
        wb.Worksheet("Data").Cell("A4").Value = "never written";
        wb.Encryptor = new FailingEncryptor();

        await Assert.That(() => wb.Save()).Throws<EncryptionFailure>();

        // Truncating first would have emptied this stream and then spent the encryption with
        // nothing in it, which is what the ordering in SaveEncrypted exists to avoid.
        await AssertBytesAreUnchanged(encrypted, origin.ToArray());

        // And what is still there is still a workbook, not merely the right number of bytes.
        using var reopened = new XLWorkbook(new MemoryStream(origin.ToArray()), new LoadOptions { Password = Password });
        await Assert.That(reopened.Worksheet("Data").Cell("A1").GetString()).IsEqualTo("Hello encrypted world");
    }

    [Test]
    public async Task SaveAs_leaves_the_file_it_would_overwrite_untouched_when_the_encryption_fails()
    {
        using var file = new TemporaryFile();
        WriteEncryptedFile(file.Path, Password, "the file that was already there");
        var before = File.ReadAllBytes(file.Path);

        using var wb = new XLWorkbook();
        wb.AddWorksheet("Data").Cell("A1").Value = "the workbook that failed to encrypt";
        wb.Encryptor = new FailingEncryptor();

        await Assert.That(() => wb.SaveAs(file.Path, new SaveOptions { Password = Password }))
            .Throws<EncryptionFailure>();
        await AssertBytesAreUnchanged(before, File.ReadAllBytes(file.Path));
    }

    [Test]
    public async Task SaveAs_to_a_stream_reports_a_failed_encryption()
    {
        using var destination = new MemoryStream();
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Data").Cell("A1").Value = "the workbook that failed to encrypt";
        wb.Encryptor = new FailingEncryptor();

        // This path encrypts straight into the caller's stream rather than into a buffer, because
        // there is nothing to lose: SaveAs to a stream writes from wherever the stream is and never
        // truncates, so it has no prior content of its own to protect. What it owes the caller is
        // the failure itself.
        await Assert.That(() => wb.SaveAs(destination, new SaveOptions { Password = Password }))
            .Throws<EncryptionFailure>();
    }

    [Test]
    public async Task Save_replaces_the_file_when_the_encryption_succeeds()
    {
        // The control for the three tests above: the destination is left alone because the
        // encryption failed, not because a save with this shape never writes anything.
        using var file = new TemporaryFile();
        WriteEncryptedFile(file.Path, Password);
        var before = File.ReadAllBytes(file.Path);

        using (var wb = new XLWorkbook(file.Path, new LoadOptions { Password = Password }))
        {
            wb.Worksheet("Data").Cell("A1").Value = "after";
            wb.Save();
        }

        var after = File.ReadAllBytes(file.Path);
        await Assert.That(after.SequenceEqual(before)).IsFalse();

        using var reopened = new XLWorkbook(file.Path, new LoadOptions { Password = Password });
        await Assert.That(reopened.Worksheet("Data").Cell("A1").GetString()).IsEqualTo("after");
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
