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

    private static Stream FixtureStream() => TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(VbaFixture));

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
