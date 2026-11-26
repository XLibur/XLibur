using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Saving;

public class MacroFreeSaveTests
{
    private const string VbaFixture = @"TryToLoad\LO\xlsm\vba-user-function.xlsm";
    private const string SignatureRelType = "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature";
    private const string SignatureContentType = "application/vnd.ms-office.vbaProjectSignature";

    [Test]
    [Arguments("xlsx")]
    [Arguments("xltx")]
    public async Task SaveAs_macro_free_type_refuses_a_workbook_loaded_from_a_file(string extension)
    {
        using var source = CopyFixtureToFile();
        using var target = new TemporaryFile(TempPath(extension));
        using var wb = new XLWorkbook(source.Path);

        var ex = await Assert.That(() => wb.SaveAs(target.Path)).Throws<InvalidOperationException>();

        await Assert.That(ex!.Message).Contains("/xl/vbaProject.bin");
        await Assert.That(ex.Message).Contains(nameof(SaveOptions.DiscardVbaProject));
        await Assert.That(File.Exists(target.Path)).IsFalse();
    }

    [Test]
    public async Task SaveAs_macro_free_type_refuses_a_workbook_loaded_from_a_stream()
    {
        using var source = FixtureStream();
        using var target = new TemporaryFile();
        using var wb = new XLWorkbook(source);

        var ex = await Assert.That(() => wb.SaveAs(target.Path)).Throws<InvalidOperationException>();

        await Assert.That(ex!.Message).Contains("/xl/vbaProject.bin");
        await Assert.That(File.Exists(target.Path)).IsFalse();
    }

    [Test]
    public async Task SaveAs_xlsx_with_DiscardVbaProject_writes_a_macro_free_package()
    {
        using var source = CopyFixtureToFile();
        using var target = new TemporaryFile();
        using (var wb = new XLWorkbook(source.Path))
            wb.SaveAs(target.Path, new SaveOptions { DiscardVbaProject = true });

        using (var package = SpreadsheetDocument.Open(target.Path, false))
            await Assert.That(package.DocumentType).IsEqualTo(SpreadsheetDocumentType.Workbook);
        await Assert.That(VbaTraces(target.Path)).IsEmpty();
    }

    [Test]
    public async Task SaveAs_xlsx_with_DiscardVbaProject_from_a_stream_writes_a_macro_free_package()
    {
        using var source = FixtureStream();
        using var target = new TemporaryFile();
        using (var wb = new XLWorkbook(source))
            wb.SaveAs(target.Path, new SaveOptions { DiscardVbaProject = true });

        using var package = SpreadsheetDocument.Open(target.Path, false);
        await Assert.That(package.DocumentType).IsEqualTo(SpreadsheetDocumentType.Workbook);
        await Assert.That(package.WorkbookPart!.VbaProjectPart).IsNull();
    }

    [Test]
    public async Task DiscardVbaProject_removes_the_signature_parts_with_the_project()
    {
        using var source = CopyFixtureToFile();
        AddVbaProjectSignature(source.Path);

        using var target = new TemporaryFile();
        using var wb = new XLWorkbook(source.Path);
        var ex = await Assert.That(() => wb.SaveAs(target.Path)).Throws<InvalidOperationException>();
        await Assert.That(ex!.Message).Contains("vbaProjectSignature");

        wb.SaveAs(target.Path, new SaveOptions { DiscardVbaProject = true });

        await Assert.That(VbaTraces(target.Path)).IsEmpty();
    }

    [Test]
    public async Task SaveAs_xlsm_keeps_the_vba_project()
    {
        using var source = CopyFixtureToFile();
        using var target = new TemporaryFile(TempPath("xlsm"));
        using (var wb = new XLWorkbook(source.Path))
            wb.SaveAs(target.Path);

        using var package = SpreadsheetDocument.Open(target.Path, false);
        await Assert.That(package.DocumentType).IsEqualTo(SpreadsheetDocumentType.MacroEnabledWorkbook);
        await Assert.That(package.WorkbookPart!.VbaProjectPart).IsNotNull();
    }

    [Test]
    public async Task SaveAs_stream_keeps_the_vba_project_and_the_macro_enabled_type()
    {
        using var source = FixtureStream();
        using var target = new MemoryStream();
        using (var wb = new XLWorkbook(source))
            wb.SaveAs(target);

        target.Position = 0;
        using var package = SpreadsheetDocument.Open(target, false);
        await Assert.That(package.DocumentType).IsEqualTo(SpreadsheetDocumentType.MacroEnabledWorkbook);
        await Assert.That(package.WorkbookPart!.VbaProjectPart).IsNotNull();
    }

    [Test]
    [Arguments(false)]
    [Arguments(true)]
    public async Task Refused_SaveAs_leaves_an_existing_target_untouched(bool fromStream)
    {
        using var source = CopyFixtureToFile();
        using var target = new TemporaryFile();
        byte[] sentinel = [1, 2, 3, 4, 5];
        File.WriteAllBytes(target.Path, sentinel);
        using var sourceStream = fromStream ? File.OpenRead(source.Path) : null;
        using var wb = sourceStream is null ? new XLWorkbook(source.Path) : new XLWorkbook(sourceStream);

        await Assert.That(() => wb.SaveAs(target.Path)).Throws<InvalidOperationException>();

        await Assert.That(File.ReadAllBytes(target.Path)).IsEquivalentTo(sentinel);
        await Assert.That(Directory.GetFiles(Path.GetDirectoryName(target.Path)!, Path.GetFileName(target.Path) + "*"))
            .IsEquivalentTo([target.Path]);
    }

    [Test]
    public async Task Save_after_a_discarding_SaveAs_keeps_the_file_macro_free_and_writes_the_edit()
    {
        using var source = CopyFixtureToFile();
        using var target = new TemporaryFile();
        using (var wb = new XLWorkbook(source.Path))
        {
            wb.SaveAs(target.Path, new SaveOptions { DiscardVbaProject = true });
            wb.Worksheets.First().Cell("Z1").Value = "edited";
            wb.Save();
        }

        using (var package = SpreadsheetDocument.Open(target.Path, false))
            await Assert.That(package.DocumentType).IsEqualTo(SpreadsheetDocumentType.Workbook);
        await Assert.That(VbaTraces(target.Path)).IsEmpty();
        await Assert.That(ReadZ1(target.Path)).IsEqualTo("edited");
    }

    [Test]
    public async Task Save_after_SaveAs_xlsm_from_a_stream_keeps_the_vba_project_and_writes_the_edit()
    {
        using var source = FixtureStream();
        using var target = new TemporaryFile(TempPath("xlsm"));
        using (var wb = new XLWorkbook(source))
        {
            wb.SaveAs(target.Path);
            wb.Worksheets.First().Cell("Z1").Value = "edited";
            wb.Save();
        }

        using (var package = SpreadsheetDocument.Open(target.Path, false))
        {
            await Assert.That(package.DocumentType).IsEqualTo(SpreadsheetDocumentType.MacroEnabledWorkbook);
            await Assert.That(package.WorkbookPart!.VbaProjectPart).IsNotNull();
        }

        await Assert.That(ReadZ1(target.Path)).IsEqualTo("edited");
    }

    [Test]
    public async Task Save_after_SaveAs_xltx_from_a_stream_keeps_the_template_type()
    {
        using var source = new MemoryStream();
        using (var created = new XLWorkbook())
        {
            created.AddWorksheet();
            created.SaveAs(source);
        }

        source.Position = 0;
        using var target = new TemporaryFile(TempPath("xltx"));
        using (var wb = new XLWorkbook(source))
        {
            wb.SaveAs(target.Path);
            wb.Save();
        }

        using var package = SpreadsheetDocument.Open(target.Path, false);
        await Assert.That(package.DocumentType).IsEqualTo(SpreadsheetDocumentType.Template);
    }

    [Test]
    [RunOn(TUnit.Core.Enums.OS.Windows)]
    [Arguments(false)]
    [Arguments(true)]
    public async Task SaveAs_onto_a_target_held_open_throws_IOException_and_leaves_it_intact(bool fromStream)
    {
        using var source = PlainWorkbookFile();
        using var target = new TemporaryFile();
        byte[] sentinel = [1, 2, 3, 4, 5];
        File.WriteAllBytes(target.Path, sentinel);
        using var sourceStream = fromStream ? File.OpenRead(source.Path) : null;
        using var wb = sourceStream is null ? new XLWorkbook(source.Path) : new XLWorkbook(sourceStream);
        using var tempFolder = new PrivateTempFolder();

        using (new FileStream(target.Path, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            var ex = await Assert.That(() => wb.SaveAs(target.Path)).Throws<IOException>();
            await Assert.That(ex!.Message).Contains(target.Path);
        }

        await Assert.That(File.ReadAllBytes(target.Path)).IsEquivalentTo(sentinel);
        await Assert.That(Directory.GetFiles(Path.GetDirectoryName(target.Path)!, Path.GetFileName(target.Path) + "*"))
            .IsEquivalentTo([target.Path]);
        await Assert.That(tempFolder.Files).IsEmpty();
    }

    [Test]
    [Arguments(false)]
    [Arguments(true)]
    public async Task SaveAs_leaves_no_temporary_file_behind(bool refused)
    {
        using var source = CopyFixtureToFile();
        using var target = new TemporaryFile();
        using var wb = new XLWorkbook(source.Path);
        using var tempFolder = new PrivateTempFolder();

        if (refused)
            await Assert.That(() => wb.SaveAs(target.Path)).Throws<InvalidOperationException>();
        else
            wb.SaveAs(target.Path, new SaveOptions { DiscardVbaProject = true });

        string[] expected = refused ? [] : [target.Path];
        await Assert.That(tempFolder.Files).IsEmpty();
        await Assert.That(Directory.GetFiles(Path.GetDirectoryName(target.Path)!, Path.GetFileName(target.Path) + "*"))
            .IsEquivalentTo(expected);
    }

    [Test]
    [RunOn(TUnit.Core.Enums.OS.Windows)]
    [Arguments("forward slashes")]
    [Arguments("upper case")]
    public async Task SaveAs_to_another_spelling_of_the_loaded_file_writes_the_edit(string spelling)
    {
        using var source = PlainWorkbookFile();
        var otherSpelling = spelling == "upper case"
            ? source.Path.ToUpperInvariant()
            : source.Path.Replace('\\', '/');

        using (var wb = new XLWorkbook(source.Path))
        {
            wb.Worksheets.First().Cell("Z1").Value = "edited";
            wb.SaveAs(otherSpelling);
        }

        await Assert.That(ReadZ1(source.Path)).IsEqualTo("edited");
    }

    [Test]
    [Arguments(false)]
    [Arguments(true)]
    public async Task Encrypted_SaveAs_xlsx_refuses_a_workbook_with_a_vba_project(bool fromStream)
    {
        using var source = CopyFixtureToFile();
        using var target = new TemporaryFile();
        using var sourceStream = fromStream ? File.OpenRead(source.Path) : null;
        using var wb = sourceStream is null ? new XLWorkbook(source.Path) : new XLWorkbook(sourceStream);

        var ex = await Assert.That(() => wb.SaveAs(target.Path, new SaveOptions { Password = EncryptionPassword }))
            .Throws<InvalidOperationException>();

        await Assert.That(ex!.Message).Contains("/xl/vbaProject.bin");
        await Assert.That(File.Exists(target.Path)).IsFalse();
    }

    [Test]
    public async Task Encrypted_SaveAs_xlsx_with_DiscardVbaProject_writes_a_macro_free_package()
    {
        using var source = CopyFixtureToFile();
        using var target = new TemporaryFile();
        using (var wb = new XLWorkbook(source.Path))
            wb.SaveAs(target.Path, new SaveOptions { Password = EncryptionPassword, DiscardVbaProject = true });

        using var package = OpenEncrypted(target.Path);
        await Assert.That(package.DocumentType).IsEqualTo(SpreadsheetDocumentType.Workbook);
        await Assert.That(package.WorkbookPart!.VbaProjectPart).IsNull();
    }

    [Test]
    public async Task Encrypted_SaveAs_xlsm_keeps_the_vba_project()
    {
        using var source = CopyFixtureToFile();
        using var target = new TemporaryFile(TempPath("xlsm"));
        using (var wb = new XLWorkbook(source.Path))
            wb.SaveAs(target.Path, new SaveOptions { Password = EncryptionPassword });

        using var package = OpenEncrypted(target.Path);
        await Assert.That(package.DocumentType).IsEqualTo(SpreadsheetDocumentType.MacroEnabledWorkbook);
        await Assert.That(package.WorkbookPart!.VbaProjectPart).IsNotNull();
    }

    [Test]
    [Arguments("xlsm", SpreadsheetDocumentType.MacroEnabledWorkbook)]
    [Arguments("xltx", SpreadsheetDocumentType.Template)]
    public async Task Encrypted_SaveAs_of_a_new_workbook_takes_the_document_type_of_the_extension(
        string extension, SpreadsheetDocumentType expected)
    {
        using var target = new TemporaryFile(TempPath(extension));
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet();
            wb.SaveAs(target.Path, new SaveOptions { Password = EncryptionPassword });
        }

        using var package = OpenEncrypted(target.Path);
        await Assert.That(package.DocumentType).IsEqualTo(expected);
    }

    private const string EncryptionPassword = "pw";

    private static SpreadsheetDocument OpenEncrypted(string path)
    {
        using var container = File.OpenRead(path);
        return SpreadsheetDocument.Open(XLibur.Excel.IO.Encryption.WorkbookEncryption.Decrypt(container, EncryptionPassword), false);
    }

    private static TemporaryFile PlainWorkbookFile()
    {
        var file = new TemporaryFile();
        using var wb = new XLWorkbook();
        wb.AddWorksheet().Cell("A1").Value = 1;
        wb.SaveAs(file.Path);
        return file;
    }

    /// <summary>
    /// Points <see cref="Path.GetTempPath"/> at an empty folder of its own, so a test can see every
    /// temporary file a save leaves there. The suite runs serially, so changing the process
    /// environment is safe.
    /// </summary>
    private sealed class PrivateTempFolder : IDisposable
    {
        private static readonly string[] Variables = ["TMP", "TEMP", "TMPDIR"];
        private readonly string?[] _previous = Variables.Select(Environment.GetEnvironmentVariable).ToArray();
        private readonly string _folder = Directory.CreateTempSubdirectory("xlibur-save-").FullName;

        public PrivateTempFolder()
        {
            foreach (var variable in Variables)
                Environment.SetEnvironmentVariable(variable, _folder);
        }

        public string[] Files => Directory.GetFileSystemEntries(_folder);

        public void Dispose()
        {
            for (var i = 0; i < Variables.Length; i++)
                Environment.SetEnvironmentVariable(Variables[i], _previous[i]);
            Directory.Delete(_folder, recursive: true);
        }
    }

    private static string ReadZ1(string path)
    {
        using var wb = new XLWorkbook(path);
        return wb.Worksheets.First().Cell("Z1").GetString();
    }

    private static Stream FixtureStream() =>TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(VbaFixture));

    private static TemporaryFile CopyFixtureToFile()
    {
        var file = new TemporaryFile(TempPath("xlsm"));
        using var resource = FixtureStream();
        using var output = File.Create(file.Path);
        resource.CopyTo(output);
        return file;
    }

    /// <summary>
    /// Signs the VBA project the way Excel lays it out: the signature part, its relationship from
    /// <c>vbaProject.bin</c>, and its content type.
    /// </summary>
    private static void AddVbaProjectSignature(string path)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update);

        using (var signature = zip.CreateEntry("xl/vbaProjectSignature.bin").Open())
            signature.Write([1, 2, 3]);

        using (var rels = new StreamWriter(zip.CreateEntry("xl/_rels/vbaProject.bin.rels").Open()))
            rels.Write($"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{SignatureRelType}" Target="vbaProjectSignature.bin"/></Relationships>""");

        var contentTypes = zip.GetEntry("[Content_Types].xml")!;
        string xml;
        using (var reader = new StreamReader(contentTypes.Open()))
            xml = reader.ReadToEnd();
        contentTypes.Delete();
        using var writer = new StreamWriter(zip.CreateEntry("[Content_Types].xml").Open());
        writer.Write(xml.Replace("</Types>", $"""<Override PartName="/xl/vbaProjectSignature.bin" ContentType="{SignatureContentType}"/></Types>"""));
    }

    private static string TempPath(string extension) =>
        Path.ChangeExtension(Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()), extension);

    /// <summary>
    /// A <c>Default</c> content type for <c>.bin</c> may stay: Excel opened the files saved for #902
    /// with it, over COM.
    /// </summary>
    private static string[] VbaTraces(string path)
    {
        using var zip = ZipFile.OpenRead(path);
        return zip.Entries
            .Where(e => MentionsVba(e.FullName)
                        || (e.FullName.EndsWith(".rels", StringComparison.Ordinal) && MentionsVba(Read(e)))
                        || (e.FullName == "[Content_Types].xml" && Read(e).Contains("PartName=\"/xl/vba", StringComparison.OrdinalIgnoreCase)))
            .Select(e => e.FullName)
            .ToArray();

        static bool MentionsVba(string text) => text.Contains("vba", StringComparison.OrdinalIgnoreCase);

        static string Read(ZipArchiveEntry entry)
        {
            using var reader = new StreamReader(entry.Open());
            return reader.ReadToEnd();
        }
    }
}
