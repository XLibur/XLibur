using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.IO;

namespace XLibur.Tests.Excel.Loading;

/// <summary>
/// Loading a package whose structure is broken must produce a deliberate rejection, never a
/// fault. <see cref="NullReferenceException"/> and the package reader's own exception types both
/// leave a caller unable to tell "this file is bad" from "this library is bad".
///
/// The packages here are synthesised rather than committed as binary fixtures, so the defect's
/// shape is legible in the test: a reader can see exactly which part is present and which is not.
/// </summary>
public class MalformedPackageLoadingTests
{
    private const string ContentTypesXml =
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
          <Default Extension="xml" ContentType="application/xml" />
        </Types>
        """;

    private const string EmptyRelationshipsXml =
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
        """;

    private const string WorkbookRelationshipXml =
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
                        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                        Target="xl/workbook.xml" />
        </Relationships>
        """;

    /// <summary>
    /// A valid OPC package that declares no workbook at all. Found by the fuzz harness — five
    /// separate crash artifacts reduced to this one shape (D27).
    /// </summary>
    [Test]
    public async Task Package_without_a_workbook_part_is_rejected_rather_than_faulting()
    {
        using var package = BuildPackage(
            ("[Content_Types].xml", ContentTypesXml),
            ("_rels/.rels", EmptyRelationshipsXml));

        await Assert.That(() => new XLWorkbook(package)).Throws<PartStructureException>();
    }

    /// <summary>
    /// A package whose relationship names a workbook part that was never written. The package
    /// reader signals this with an exception type of its own; it must not reach the caller.
    /// </summary>
    [Test]
    public async Task Package_whose_relationship_names_an_absent_part_is_rejected_rather_than_faulting()
    {
        using var package = BuildPackage(
            ("[Content_Types].xml", ContentTypesXml),
            ("_rels/.rels", WorkbookRelationshipXml));

        await Assert.That(() => new XLWorkbook(package)).Throws<PartStructureException>();
    }

    /// <summary>
    /// A stream that is not an archive at all is rejected with <see cref="FileFormatException"/>,
    /// a BCL type raised by the packaging layer rather than by the OpenXml SDK.
    ///
    /// This is deliberately *not* converted to <see cref="PartStructureException"/>. The rule
    /// XLibur holds itself to is that no <c>DocumentFormat.OpenXml</c> type escapes a public
    /// constructor; a documented BCL format exception is already a perfectly legible rejection and
    /// wrapping it would hide a well-understood type behind a vaguer one. The test pins the
    /// behaviour so that a future change to the rejection type is a decision rather than a drift.
    /// </summary>
    [Test]
    public async Task Stream_that_is_not_a_package_is_rejected_with_a_format_exception()
    {
        using var notAPackage = new MemoryStream("this is not a spreadsheet"u8.ToArray(), writable: false);

        await Assert.That(() => new XLWorkbook(notAPackage)).Throws<FileFormatException>();
    }

    private static MemoryStream BuildPackage(params (string Name, string Content)[] entries)
    {
        var buffer = new MemoryStream();
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (name, content) in entries)
            {
                using var entryStream = archive.CreateEntry(name).Open();
                var bytes = Encoding.UTF8.GetBytes(content);
                entryStream.Write(bytes, 0, bytes.Length);
            }
        }

        buffer.Position = 0;
        return buffer;
    }
}
