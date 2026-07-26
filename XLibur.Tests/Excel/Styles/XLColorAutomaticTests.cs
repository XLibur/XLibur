using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using NUnit.Framework;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Styles;

/// <summary>
/// The automatic color - ECMA-376 <c>CT_Color/@auto</c>, shown as "Automatic" in Excel's font color
/// picker. The application resolves the actual color from the context it is used in, so it must not
/// be pinned to a concrete value on save.
/// </summary>
[TestFixture]
public class XLColorAutomaticTests
{
    [Test]
    public void Automatic_IsTheFirstColorType()
    {
        // Deliberately ordinal 0 so a default XLColorKey describes itself as automatic instead of
        // masquerading as a fully transparent RGB black.
        Assert.That((int)XLColorType.Automatic, Is.Zero);
    }

    [Test]
    public void Automatic_AndNoColor_AreTheSameValue()
    {
        Assert.Multiple(() =>
        {
#pragma warning disable CS0618 // NoColor is deprecated; this test exists to pin the alias.
            Assert.That(XLColor.NoColor, Is.SameAs(XLColor.Automatic),
                "NoColor is only the GUI label some Excel pickers use for the automatic color.");
#pragma warning restore CS0618
            Assert.That(XLColor.Automatic.ColorType, Is.EqualTo(XLColorType.Automatic));
            Assert.That(XLColor.Automatic.IsAutomatic, Is.True);
            Assert.That(XLColor.FromArgb(0, 0, 0).IsAutomatic, Is.False,
                "An explicit black is a stated color, not an automatic one.");
        });
    }

    [Test]
    public void Automatic_ToString_IsNotAnRgbValue()
    {
        Assert.That(XLColor.Automatic.ToString(), Is.EqualTo("Automatic"));
    }

    [Test]
    public void Automatic_HasNoRgbValueToRead()
    {
        // The automatic color carries no color value; reading one would silently hand back the
        // meaningless all-zero ARGB that used to leak into saved files.
        Assert.Throws<InvalidOperationException>(() => _ = XLColor.Automatic.Color);
    }

    [Test]
    public void AutomaticFontColor_IsWrittenAsAutoRatherThanTransparentBlack()
    {
        // The automatic color has no rgb/indexed/theme, so it used to fall through to the RGB branch
        // on save and be written as rgb="00000000" - a fully transparent black that no Excel file
        // means to express, and which pins down a color the source left to the application.
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "Auto";
            ws.Cell("A1").Style.Font.FontColor = XLColor.Automatic;
            wb.SaveAs(ms);
        }

        // Assert against the <fonts> block alone: a solid fill writes its own <fgColor auto="1"/>,
        // which would make a whole-document match pass for the wrong reason.
        var fonts = FontsBlock(ReadPart(ms.ToArray(), "xl/styles.xml"));

        Assert.Multiple(() =>
        {
            Assert.That(fonts, Does.Not.Contain("00000000"),
                $"The automatic font color was written as a transparent black.\n\n{fonts}");
            Assert.That(fonts, Does.Contain("auto=\"1\""),
                $"The automatic font color was not written.\n\n{fonts}");
        });
    }

    [Test]
    public void AutomaticFontColor_SurvivesARoundTrip()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "Auto";
            ws.Cell("A1").Style.Font.FontColor = XLColor.Automatic;
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var reloaded = new XLWorkbook(ms);

        Assert.That(reloaded.Worksheets.First().Cell("A1").Style.Font.FontColor.IsAutomatic, Is.True,
            "The automatic font color did not survive a save/load round-trip.");
    }

    private static string FontsBlock(string stylesXml)
    {
        var match = System.Text.RegularExpressions.Regex.Match(stylesXml, "<x:fonts.*?</x:fonts>",
            System.Text.RegularExpressions.RegexOptions.Singleline);
        Assert.That(match.Success, Is.True, $"No <fonts> block in styles.xml.\n\n{stylesXml}");
        return match.Value;
    }

    [Test]
    public void FontWithAutoColor_LoadsAsAutomatic()
    {
        var input = BuildWorkbook();

        using var wb = new XLWorkbook(new MemoryStream(input));
        var color = wb.Worksheets.First().Cell("A1").Style.Font.FontColor;

        Assert.That(color.IsAutomatic, Is.True, "auto=\"1\" should load as the automatic color.");
    }

    private static string ReadPart(byte[] xlsx, string partPath)
    {
        using var zip = new ZipArchive(new MemoryStream(xlsx), ZipArchiveMode.Read);
        var entry = zip.GetEntry(partPath) ?? throw new AssertionException($"Missing part: {partPath}");
        using var r = new StreamReader(entry.Open());
        return r.ReadToEnd();
    }

    /// <summary>
    /// A minimal package whose only font states an explicitly automatic color. The input can't be
    /// produced through the <see cref="XLWorkbook"/> API, so it is built by hand.
    /// </summary>
    private static byte[] BuildWorkbook()
    {
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
                  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """),
            ("xl/styles.xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <fonts count="1"><font><sz val="11"/><color auto="1"/><name val="Calibri"/></font></fonts>
                  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
                  <borders count="1"><border/></borders>
                  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
                  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>
                </styleSheet>
                """),
            ("xl/worksheets/sheet1.xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData><row r="1"><c r="A1" s="0" t="str"><v>Auto</v></c></row></sheetData>
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
