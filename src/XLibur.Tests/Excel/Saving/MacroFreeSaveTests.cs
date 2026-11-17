using System;
using System.IO;
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
    }

    [Test]
    public async Task SaveAs_macro_free_type_refuses_a_workbook_loaded_from_a_stream()
    {
        using var source = FixtureStream();
        using var target = new TemporaryFile();
        using var wb = new XLWorkbook(source);

        var ex = await Assert.That(() => wb.SaveAs(target.Path)).Throws<InvalidOperationException>();

        await Assert.That(ex!.Message).Contains("/xl/vbaProject.bin");
    }

    [Test]
    public async Task SaveAs_xlsx_with_DiscardVbaProject_writes_a_macro_free_package()
    {
        using var source = CopyFixtureToFile();
        using var target = new TemporaryFile();
        using (var wb = new XLWorkbook(source.Path))
            wb.SaveAs(target.Path, new SaveOptions { DiscardVbaProject = true });

        using var package = SpreadsheetDocument.Open(target.Path, false);
        await Assert.That(package.DocumentType).IsEqualTo(SpreadsheetDocumentType.Workbook);
        await Assert.That(package.WorkbookPart!.VbaProjectPart).IsNull();
        await Assert.That(PartNames(package).Any(n => n.Contains("vba", StringComparison.OrdinalIgnoreCase))).IsFalse();
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
        using (var package = SpreadsheetDocument.Open(source.Path, true))
        {
            var signature = package.WorkbookPart!.VbaProjectPart!.AddExtendedPart(SignatureRelType, SignatureContentType, ".bin");
            using var data = new MemoryStream([1, 2, 3]);
            signature.FeedData(data);
        }

        using var target = new TemporaryFile();
        using var wb = new XLWorkbook(source.Path);
        var ex = await Assert.That(() => wb.SaveAs(target.Path)).Throws<InvalidOperationException>();
        await Assert.That(ex!.Message).Contains("vbaProjectSignature");

        wb.SaveAs(target.Path, new SaveOptions { DiscardVbaProject = true });

        using var saved = SpreadsheetDocument.Open(target.Path, false);
        await Assert.That(PartNames(saved).Any(n => n.Contains("vba", StringComparison.OrdinalIgnoreCase))).IsFalse();
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

    private static string TempPath(string extension) =>
        Path.ChangeExtension(Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()), extension);

    private static string[] PartNames(OpenXmlPackage package) =>
        package.GetAllParts().Select(p => p.Uri.OriginalString).ToArray();
}
