using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using NUnit.Framework;
using XLibur.Excel;

namespace XLibur.Tests.Excel.RichText;

/// <summary>
/// Round-trip fidelity of rich text whose runs carry no explicit colour (issue #219). The input
/// can't be produced through the <see cref="XLWorkbook"/> API — the writer would normalise it
/// before it were ever saved — so these tests build a minimal spec-valid package in memory.
/// </summary>
[TestFixture]
public class RichTextColorRoundTripTests
{
    // si[0] - rich-text runs whose rPr has NO <color> element, plus a run with no rPr at all.
    // si[1] - plain text + phonetic guide (rPh/phoneticPr), i.e. NOT rich text.
    // Neither carries any colour.
    private const string SharedStrings =
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
          <si>
            <r><rPr><b/><sz val="11"/><rFont val="Calibri"/></rPr><t>Bold</t></r>
            <r><t xml:space="preserve"> plain</t></r>
          </si>
          <si>
            <t>&#28450;&#23383;</t>
            <rPh sb="0" eb="2"><t>&#12363;&#12435;&#12376;</t></rPh>
            <phoneticPr fontId="0" type="Hiragana"/>
          </si>
        </sst>
        """;

    [Test]
    public void ColorlessRuns_AreNotSerializedAsExplicitBlack()
    {
        var savedSharedStrings = RoundTripSharedStrings(SharedStrings);

        // The source had no <color> anywhere; the round-trip must not invent black.
        Assert.That(savedSharedStrings, Does.Not.Contain("FF000000").IgnoreCase,
            $"No explicit black colour was present in the source, but the save injected one.\n\n{savedSharedStrings}");
    }

    [Test]
    public void PhoneticOnlyString_IsNotPromotedToRun()
    {
        var savedSharedStrings = RoundTripSharedStrings(SharedStrings);

        Assert.Multiple(() =>
        {
            Assert.That(savedSharedStrings, Does.Contain("rPh"), "Phonetic run (rPh) should be preserved.");
            Assert.That(savedSharedStrings, Does.Contain("漢字"), "Phonetic base text should be preserved.");

            // Two <si> entries, only the first of which is rich text, so exactly two runs.
            Assert.That(CountOccurrences(savedSharedStrings, "<x:r>"), Is.EqualTo(2),
                $"Plain text with a phonetic guide must not be promoted to a rich-text run.\n\n{savedSharedStrings}");
        });
    }

    [Test]
    public void ExplicitRunColor_IsPreserved()
    {
        // The counterpart of the above: a colour that *was* written must survive the round-trip,
        // so the fix can't just drop every run colour.
        const string coloredSharedStrings =
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
              <si>
                <r><rPr><color rgb="FFFF0000"/><sz val="11"/><rFont val="Calibri"/></rPr><t>Red</t></r>
              </si>
            </sst>
            """;

        var savedSharedStrings = RoundTripSharedStrings(coloredSharedStrings, siCount: 1);

        Assert.That(savedSharedStrings, Does.Contain("FFFF0000").IgnoreCase,
            $"Explicitly written run colour was lost.\n\n{savedSharedStrings}");
    }

    [Test]
    public void ColorlessRunText_IsStillReadable()
    {
        var input = BuildWorkbook(SharedStrings, siCount: 2);

        using var wb = new XLWorkbook(new MemoryStream(input));
        var ws = wb.Worksheets.First();

        Assert.Multiple(() =>
        {
            Assert.That(ws.Cell("A1").GetString(), Is.EqualTo("Bold plain"));
            Assert.That(ws.Cell("A2").GetString(), Is.EqualTo("漢字"));
        });
    }

    [Test]
    public void RunWithoutRunProperties_StillInheritsAnExplicitCellFontColor()
    {
        // A run with no rPr inherits the cell font (ECMA-376 CT_RElt). Dropping the color is only
        // correct when it is the ambiguous black default - an explicitly styled cell font must
        // still be materialized onto the run.
        const string greenFontStyles =
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <fonts count="1"><font><sz val="11"/><color rgb="FF00FF00"/><name val="Calibri"/></font></fonts>
              <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
              <borders count="1"><border/></borders>
              <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
              <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>
            </styleSheet>
            """;
        const string bareRun =
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
              <si><r><t>Inherited</t></r></si>
            </sst>
            """;

        var savedSharedStrings = RoundTripSharedStrings(bareRun, siCount: 1, styles: greenFontStyles);

        Assert.That(savedSharedStrings, Does.Contain("FF00FF00").IgnoreCase,
            $"A run without rPr must keep the cell font's explicit color.\n\n{savedSharedStrings}");
    }

    [Test]
    public void EditingPhoneticOnlyString_MaterializesARunWithoutLosingText()
    {
        // The stored rich text has no runs, but the mutable API is run-based. Touching it must
        // materialize a run rather than write back an empty cell.
        var input = BuildWorkbook(SharedStrings, siCount: 2);

        using var wb = new XLWorkbook(new MemoryStream(input));
        var cell = wb.Worksheets.First().Cell("A2");

        var richText = cell.GetRichText();
        Assert.That(richText.Text, Is.EqualTo("漢字"), "Runless rich text lost its text on the mutable path.");

        richText.AddText("!");

        Assert.Multiple(() =>
        {
            Assert.That(cell.GetString(), Is.EqualTo("漢字!"));
            Assert.That(richText.Phonetics.Count, Is.EqualTo(1), "Phonetics were lost when a run was materialized.");
        });
    }

    [Test]
    public void PhoneticOnlyString_IsMeasuredWhenAdjustingToContents()
    {
        // Column sizing walks the rich text runs; a runless rich text must fall back to the plain
        // text and cell font instead of measuring as empty.
        var input = BuildWorkbook(SharedStrings, siCount: 2);

        using var wb = new XLWorkbook(new MemoryStream(input));
        var ws = wb.Worksheets.First();
        var defaultWidth = ws.ColumnWidth;

        ws.Column(1).AdjustToContents();

        Assert.That(ws.Column(1).Width, Is.GreaterThan(defaultWidth),
            "A phonetic-only cell was measured as if it had no text.");
    }

    private static string RoundTripSharedStrings(string sharedStrings, int siCount = 2, string styles = "")
    {
        var input = BuildWorkbook(sharedStrings, siCount, styles);

        using var outMs = new MemoryStream();
        using (var wb = new XLWorkbook(new MemoryStream(input)))
            wb.SaveAs(outMs);

        return ReadPart(outMs.ToArray(), "xl/sharedStrings.xml");
    }

    private static int CountOccurrences(string haystack, string needle)
    {
        var count = 0;
        var index = haystack.IndexOf(needle, System.StringComparison.Ordinal);
        while (index >= 0)
        {
            count++;
            index = haystack.IndexOf(needle, index + needle.Length, System.StringComparison.Ordinal);
        }

        return count;
    }

    private static string ReadPart(byte[] xlsx, string partPath)
    {
        using var zip = new ZipArchive(new MemoryStream(xlsx), ZipArchiveMode.Read);
        var entry = zip.GetEntry(partPath) ?? throw new AssertionException($"Missing part: {partPath}");
        using var r = new StreamReader(entry.Open());
        return r.ReadToEnd();
    }

    private static byte[] BuildWorkbook(string sharedStrings, int siCount, string styles = "")
    {
        var sheetRows = new StringBuilder();
        for (var i = 0; i < siCount; i++)
            sheetRows.Append($"<row r=\"{i + 1}\"><c r=\"A{i + 1}\" t=\"s\"><v>{i}</v></c></row>");

        var parts = new (string Path, string Content)[]
        {
            ("[Content_Types].xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
                  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
                </Types>
                """),
            ("_rels/.rels",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>
                """),
            ("xl/workbook.xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
                </workbook>
                """),
            ("xl/_rels/workbook.xml.rels",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
                  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """),
            ("xl/styles.xml", styles.Length > 0 ? styles :
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
                  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
                  <borders count="1"><border/></borders>
                  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
                  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
                </styleSheet>
                """),
            ("xl/sharedStrings.xml", sharedStrings),
            ("xl/worksheets/sheet1.xml",
                $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData>{sheetRows}</sheetData>
                </worksheet>
                """),
        };

        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (path, content) in parts)
            {
                var e = zip.CreateEntry(path, CompressionLevel.Optimal);
                using var w = new StreamWriter(e.Open(), new UTF8Encoding(false));
                w.Write(content.TrimStart());
            }
        }

        return ms.ToArray();
    }
}
