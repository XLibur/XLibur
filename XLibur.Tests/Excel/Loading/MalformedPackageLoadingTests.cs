using System;
using System.Globalization;
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

    /// <summary>
    /// A sheet whose legal name contains a doubled apostrophe loads, keeps its name, and keeps
    /// its cells.
    /// </summary>
    /// <remarks>
    /// D40. Excel forbids an apostrophe only as the first or last character, so <c>Ann''s</c> is a
    /// legal sheet name. The loader looked the sheet up through <c>TryGetWorksheet</c> and
    /// <c>Worksheet</c>, which call <c>UnescapeSheetName</c> to collapse <c>''</c> to <c>'</c> —
    /// correct for a name lifted out of a formula, wrong for the <c>name</c> attribute of
    /// <c>&lt;sheet&gt;</c>, which is the name itself.
    ///
    /// <para>
    /// The cell assertion is the important one. The lookup that reads sheet contents failed
    /// silently, into a branch commented "this shouldn't be possible", so the sheet loaded with
    /// its name intact and <em>no cells at all</em>. Asserting only that the load does not throw
    /// would pass against the unfixed code for the wrong reason, because the throwing symptom is
    /// in the pivot-table pass and needs a pivot table to reach.
    /// </para>
    /// </remarks>
    [Test]
    [Arguments("Ann''s")]
    [Arguments("a''''b")]
    [Arguments("It's")]
    public async Task Sheet_name_containing_apostrophes_keeps_its_name_and_its_cells(string sheetName)
    {
        var escaped = sheetName.Replace("'", "&apos;", StringComparison.Ordinal);
        using var package = BuildSheetPackage(
            """<row r="1"><c r="A1"><v>7</v></c></row>""",
            cellFormatCount: 1,
            sheetName: escaped);

        using var workbook = new XLWorkbook(package);

        await Assert.That(workbook.Worksheets.Count).IsEqualTo(1);

        var sheet = workbook.Worksheets.First();
        await Assert.That(sheet.Name).IsEqualTo(sheetName);
        await Assert.That(sheet.Cell("A1").GetDouble()).IsEqualTo(7d);

        // The writer had the same defect at a third site, in WorkbookPartWriter's sheet reorder,
        // and it was only reachable once the load-side fixes let such a workbook exist. Fixing
        // the load moved the ArgumentException from the constructor to SaveAs.
        using var saved = new MemoryStream();
        workbook.SaveAs(saved);

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var reloadedSheet = reloaded.Worksheets.First();
        await Assert.That(reloadedSheet.Name).IsEqualTo(sheetName);
        await Assert.That(reloadedSheet.Cell("A1").GetDouble()).IsEqualTo(7d);
    }

    /// <summary>
    /// A number that no date can be made from, in a cell whose format says it is a date.
    /// </summary>
    /// <remarks>
    /// D39, found by the structure-aware fuzz target in the <em>save</em> phase. The load typed the
    /// cell <c>DateTime</c> on the strength of the number format alone, without asking whether the
    /// number was in range. Nothing complained until <c>SaveAs</c>, where <c>CellXmlWriter</c> asks
    /// for <c>GetDateTime()</c> and <c>DateTime.FromOADate</c> throws
    /// <c>ArgumentException("Not a legal OleAut date.")</c> — a file XLibur read and then could not
    /// write. Excel keeps such a cell as a number and renders <c>####</c>.
    /// </remarks>
    [Test]
    [Arguments("1e308")]
    [Arguments("-1e308")]
    [Arguments("2958466")]
    [Arguments("-657436")]
    public async Task Out_of_range_number_in_a_date_formatted_cell_stays_a_number_and_still_saves(string literal)
    {
        using var package = BuildSheetPackage(
            $"""<row r="1"><c r="A1" s="0"><v>{literal}</v></c></row>""",
            cellFormatCount: 1,
            numberFormatId: 14);

        using var workbook = new XLWorkbook(package);
        var cell = workbook.Worksheet("Sheet1").Cell("A1");

        await Assert.That(cell.DataType).IsEqualTo(XLDataType.Number);
        await Assert.That(cell.GetDouble()).IsEqualTo(double.Parse(literal, CultureInfo.InvariantCulture));

        // The half that actually failed: a workbook that loads must be one that saves.
        using var saved = new MemoryStream();
        workbook.SaveAs(saved);

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("Sheet1").Cell("A1").DataType).IsEqualTo(XLDataType.Number);
    }

    /// <summary>
    /// A date-formatted cell holding a number a date <em>can</em> be made from is still a date.
    /// The range check must not cost the ordinary case.
    /// </summary>
    [Test]
    public async Task In_range_number_in_a_date_formatted_cell_is_still_a_date()
    {
        using var package = BuildSheetPackage(
            """<row r="1"><c r="A1" s="0"><v>44561</v></c></row>""",
            cellFormatCount: 1,
            numberFormatId: 14);

        using var workbook = new XLWorkbook(package);
        var cell = workbook.Worksheet("Sheet1").Cell("A1");

        await Assert.That(cell.DataType).IsEqualTo(XLDataType.DateTime);
        await Assert.That(cell.GetDateTime()).IsEqualTo(new DateTime(2021, 12, 31, 0, 0, 0, DateTimeKind.Unspecified));
    }

    /// <summary>
    /// A workbook declaring two sheets with the same name. The collection it loads into guards
    /// duplicates with an ArgumentException naming 'sheetName' — right for the public
    /// <c>AddWorksheet</c>, wrong for a name that came out of a file (D32).
    /// </summary>
    [Test]
    public async Task Workbook_declaring_two_sheets_with_the_same_name_is_rejected_rather_than_faulting()
    {
        // Both sheets must be genuinely loadable. With dangling relationship ids they would go to
        // UnsupportedSheets and never reach the worksheet collection, so the duplicate would not
        // be detected and the test would pass for the wrong reason.
        const string emptySheetXml =
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <sheetData />
            </worksheet>
            """;

        using var package = BuildPackage(
            ("[Content_Types].xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
                  <Default Extension="xml" ContentType="application/xml" />
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
                  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
                  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
                </Types>
                """),
            ("_rels/.rels", WorkbookRelationshipXml),
            ("xl/workbook.xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets>
                    <sheet name="Twin" sheetId="1" r:id="rIdA" />
                    <sheet name="Twin" sheetId="2" r:id="rIdB" />
                  </sheets>
                </workbook>
                """),
            ("xl/_rels/workbook.xml.rels",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" />
                  <Relationship Id="rIdB" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml" />
                </Relationships>
                """),
            ("xl/worksheets/sheet1.xml", emptySheetXml),
            ("xl/worksheets/sheet2.xml", emptySheetXml));

        await Assert.That(() => new XLWorkbook(package)).Throws<PartStructureException>();
    }

    /// <summary>
    /// A workbook declaring a sheet name the format does not permit — here one ending in an
    /// apostrophe. <c>XLHelper.ValidateSheetName</c> guards that rule with an ArgumentException
    /// naming 'sheetName', right for <c>AddWorksheet</c> and wrong for a name read from a file
    /// (D33). Third instance of the same shape, after D32 and the collection's duplicate guard.
    /// </summary>
    [Test]
    public async Task Workbook_declaring_an_illegal_sheet_name_is_rejected_rather_than_faulting()
    {
        using var package = BuildPackage(
            ("[Content_Types].xml", MinimalContentTypes),
            ("_rels/.rels", WorkbookRelationshipXml),
            ("xl/workbook.xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets>
                    <sheet name="trailing quote'" sheetId="1" r:id="rIdA" />
                  </sheets>
                </workbook>
                """),
            ("xl/_rels/workbook.xml.rels", EmptyRelationshipsXml));

        await Assert.That(() => new XLWorkbook(package)).Throws<PartStructureException>();
    }

    /// <summary>
    /// A loadable sheet and an unloadable one sharing a name. Neither collection alone sees the
    /// duplicate — the unloadable sheet goes to <c>UnsupportedSheets</c> — so the load used to
    /// succeed, and the save then wrote both declarations, producing a file XLibur itself refused
    /// to read (D34).
    ///
    /// This is the defect that only a round-trip check can find: the first load is clean, the save
    /// is clean, and the corruption exists solely in the bytes written out. It was found by the
    /// fuzz harness's load-save-load oracle rather than by any exception on the way in.
    /// </summary>
    [Test]
    public async Task Loadable_and_unloadable_sheets_sharing_a_name_are_rejected_rather_than_written_back_corrupt()
    {
        using var package = BuildPackage(
            ("[Content_Types].xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
                  <Default Extension="xml" ContentType="application/xml" />
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
                  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
                </Types>
                """),
            ("_rels/.rels", WorkbookRelationshipXml),
            ("xl/workbook.xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets>
                    <sheet name="Twin" sheetId="1" r:id="rIdGhost" />
                    <sheet name="Twin" sheetId="2" r:id="rIdReal" />
                  </sheets>
                </workbook>
                """),
            ("xl/_rels/workbook.xml.rels",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdReal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" />
                </Relationships>
                """),
            ("xl/worksheets/sheet1.xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData />
                </worksheet>
                """));

        await Assert.That(() => new XLWorkbook(package)).Throws<PartStructureException>();
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
    /// <param name="rowsXml">The contents of &lt;sheetData&gt;.</param>
    /// <param name="cellFormatCount">How many entries &lt;cellXfs&gt; declares.</param>
    /// <param name="numberFormatId">
    /// The built-in number format every cell format uses. <c>0</c> is General; <c>14</c> is the
    /// short date, which is what makes the reader type a numeric cell as a date.
    /// </param>
    /// <param name="sheetName">
    /// The sheet's name, already XML-escaped, since a name may legally contain an apostrophe.
    /// </param>
    private static MemoryStream BuildSheetPackage(
        string rowsXml, int cellFormatCount, int numberFormatId = 0, string sheetName = "Sheet1")
    {
        var cellFormats = string.Concat(
            Enumerable.Repeat(
                $"""<xf numFmtId="{numberFormatId}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" />""",
                cellFormatCount));

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
                $"""
                 <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                 <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                   <sheets><sheet name="{sheetName}" sheetId="1" r:id="rIdSheet1" /></sheets>
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
