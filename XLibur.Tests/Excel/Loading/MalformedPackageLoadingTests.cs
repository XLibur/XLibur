using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
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

    /// <summary>
    /// A workbook declaring a sheet whose relationship id names no relationship. The line above
    /// the failing one already guarded an *empty* relId, described as something non-Excel
    /// producers emit; a dangling non-empty one had not been considered, and OpenXml's
    /// <c>GetPartById</c> answers it with <see cref="ArgumentOutOfRangeException"/> naming a
    /// parameter the caller never passed (D28).
    ///
    /// The sheet is treated as one XLibur cannot load and copies through, which is how the same
    /// method already handles a relationship pointing at a chartsheet.
    /// </summary>
    [Test]
    public async Task Sheet_whose_relationship_id_names_no_part_does_not_fault()
    {
        using var package = BuildPackage(
            ("[Content_Types].xml", MinimalContentTypes),
            ("_rels/.rels", WorkbookRelationshipXml),
            ("xl/workbook.xml", WorkbookWithDanglingSheetRelationship),
            ("xl/_rels/workbook.xml.rels", EmptyRelationshipsXml));

        using var workbook = new XLWorkbook(package);

        // The declaration is preserved rather than dropped or fabricated into a real sheet.
        await Assert.That(workbook.Worksheets.Count).IsEqualTo(0);

        // Saving it is a separate, still-open problem (D31): a workbook made only of sheets
        // XLibur cannot model is refused by CheckForWorksheetsPresent, and widening that guard
        // only moves the failure to a relationship-id conflict in the write path. The load is
        // what this test pins; the save is pinned as currently-known-wrong so that fixing D31
        // fails here and forces this test to be revisited rather than silently diverging.
        using var output = new MemoryStream();
        await Assert.That(() => workbook.SaveAs(output)).Throws<InvalidOperationException>();
    }

    /// <summary>
    /// A cell whose numeric literal overflows a double. Well-formed XML, but no cell can hold
    /// infinity, so the load is refused with XLibur's own type rather than letting XLCellValue's
    /// precondition escape as an ArgumentException naming the parameter 'number' (D29).
    /// </summary>
    [Test]
    public async Task Cell_value_that_overflows_a_double_is_rejected_rather_than_faulting()
    {
        using var package = BuildSheetPackage("""<row r="1"><c r="A1"><v>1e309</v></c></row>""", cellFormatCount: 1);

        await Assert.That(() => new XLWorkbook(package)).Throws<PartStructureException>();
    }

    /// <summary>
    /// A cell naming a style index past the end of &lt;cellXfs&gt;. Excel repairs this by using
    /// the default format; a file Excel opens without complaint must not be one XLibur refuses
    /// (D30). The cost of the repair is cosmetic, where refusing costs the whole document.
    /// </summary>
    [Test]
    public async Task Cell_style_index_past_the_end_of_the_table_falls_back_to_the_default_format()
    {
        using var package = BuildSheetPackage("""<row r="1"><c r="A1" s="17"><v>1</v></c></row>""", cellFormatCount: 1);

        using var workbook = new XLWorkbook(package);

        await Assert.That(workbook.Worksheet("Sheet1").Cell("A1").GetDouble()).IsEqualTo(1d);
    }

    private const string MinimalContentTypes =
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
          <Default Extension="xml" ContentType="application/xml" />
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
        </Types>
        """;

    private const string WorkbookWithDanglingSheetRelationship =
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets>
            <sheet name="Ghost" sheetId="1" r:id="rIdThatDoesNotExist" />
          </sheets>
        </workbook>
        """;

    /// <summary>
    /// A one-sheet package whose sheet data is <paramref name="rowsXml"/> and whose stylesheet
    /// declares <paramref name="cellFormatCount"/> cell formats.
    /// </summary>
    private static MemoryStream BuildSheetPackage(string rowsXml, int cellFormatCount)
    {
        var cellFormats = string.Concat(
            Enumerable.Repeat("""<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />""", cellFormatCount));

        return BuildPackage(
            ("[Content_Types].xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
                  <Default Extension="xml" ContentType="application/xml" />
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
                  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />
                  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
                </Types>
                """),
            ("_rels/.rels", WorkbookRelationshipXml),
            ("xl/workbook.xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet1" /></sheets>
                </workbook>
                """),
            ("xl/_rels/workbook.xml.rels",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdSheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" />
                  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />
                </Relationships>
                """),
            ("xl/styles.xml",
                $"""
                 <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                 <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                   <fonts count="1"><font><sz val="11" /><name val="Calibri" /></font></fonts>
                   <fills count="1"><fill><patternFill patternType="none" /></fill></fills>
                   <borders count="1"><border /></borders>
                   <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" /></cellStyleXfs>
                   <cellXfs count="{cellFormatCount}">{cellFormats}</cellXfs>
                 </styleSheet>
                 """),
            ("xl/worksheets/sheet1.xml",
                $"""
                 <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                 <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                   <sheetData>{rowsXml}</sheetData>
                 </worksheet>
                 """));
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
