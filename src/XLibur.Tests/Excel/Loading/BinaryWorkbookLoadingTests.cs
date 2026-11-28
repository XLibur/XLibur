using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Loading;

public class BinaryWorkbookLoadingTests
{
    private const string BinaryWorkbookContentType = "application/vnd.ms-excel.sheet.binary.macroEnabled.main";
    private const string XmlWorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";

    [Test]
    public async Task Stream_holding_an_xlsb_is_refused_as_not_supported()
    {
        using var stream = BuildPackage("workbook.bin", BinaryWorkbookContentType);

        await AssertRefusedAsXlsb(() => { using var workbook = new XLWorkbook(stream); });
    }

    [Test]
    [Arguments("xlsb")]
    [Arguments("xlsx")]
    public async Task File_holding_an_xlsb_is_refused_as_not_supported(string extension)
    {
        using var file = WriteToFile(BuildPackage("workbook.bin", BinaryWorkbookContentType), extension);

        await AssertRefusedAsXlsb(() => { using var workbook = new XLWorkbook(file.Path); });
    }

    [Test]
    public async Task Template_holding_an_xlsb_is_refused_as_not_supported()
    {
        using var file = WriteToFile(BuildPackage("workbook.bin", BinaryWorkbookContentType), "xlsx");

        await AssertRefusedAsXlsb(() => XLWorkbook.OpenFromTemplate(file.Path));
    }

    [Test]
    public async Task Saving_as_an_xlsb_is_refused_as_not_supported()
    {
        using var workbook = new XLWorkbook();
        workbook.AddWorksheet();
        using var file = new TemporaryFile(Path.ChangeExtension(Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()), "xlsb"));

        await AssertRefusedAsXlsb(() => workbook.SaveAs(file.Path));
    }

    [Test]
    public async Task Xlsx_with_a_part_of_unexpected_content_type_is_not_reported_as_an_xlsb()
    {
        using var stream = BuildPackage("workbook.xml", XmlWorkbookContentType);

        var exception = await Assert.That(() => new XLWorkbook(stream)).Throws<Exception>();

        await Assert.That(exception).IsNotTypeOf<NotSupportedException>();
    }

    private static async Task AssertRefusedAsXlsb(Action action)
    {
        var exception = await Assert.That(action).ThrowsExactly<NotSupportedException>();

        await Assert.That(exception!.Message).Contains(".xlsb");
        await Assert.That(exception.Message).Contains(".xlsx");
    }

    private static TemporaryFile WriteToFile(MemoryStream package, string extension)
    {
        var file = new TemporaryFile(Path.ChangeExtension(Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()), extension));
        using (package)
            File.WriteAllBytes(file.Path, package.ToArray());

        return file;
    }

    private static MemoryStream BuildPackage(string workbookPartName, string workbookContentType)
    {
        var buffer = new MemoryStream();
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml",
                $"""
                 <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                 <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                   <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
                   <Override PartName="/xl/{workbookPartName}" ContentType="{workbookContentType}" />
                   <Override PartName="/xl/worksheets/sheet1.bin" ContentType="application/vnd.ms-excel.worksheet" />
                 </Types>
                 """);
            AddEntry(archive, "_rels/.rels",
                $"""
                 <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                 <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                   <Relationship Id="rId1"
                                 Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                                 Target="xl/{workbookPartName}" />
                 </Relationships>
                 """);
            AddEntry(archive, $"xl/_rels/{workbookPartName}.rels",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1"
                                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                                Target="worksheets/sheet1.bin" />
                </Relationships>
                """);
            AddEntry(archive, $"xl/{workbookPartName}", "placeholder");
            AddEntry(archive, "xl/worksheets/sheet1.bin", "placeholder");
        }

        buffer.Position = 0;
        return buffer;
    }

    private static void AddEntry(ZipArchive archive, string name, string content)
    {
        using var entryStream = archive.CreateEntry(name).Open();
        var bytes = Encoding.UTF8.GetBytes(content);
        entryStream.Write(bytes, 0, bytes.Length);
    }
}
