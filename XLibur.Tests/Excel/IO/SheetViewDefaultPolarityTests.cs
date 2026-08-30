using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// A save/reload round trip cannot see a polarity error in a boolean sheet-view attribute: the
/// reader assigns whatever the writer wrote, whichever way round that is, so a flipped condition in
/// the writer round-trips through XLibur's own reader without ever producing a wrong in-memory
/// value. Only reading the raw bytes and checking them against the true OOXML default (per the
/// ECMA-376 <c>CT_SheetView</c> schema) can catch that — the same reasoning
/// <see cref="WritePathAgreementTests"/> applies to the two write paths applies here to the one
/// write path against the spec it targets.
/// </summary>
public class SheetViewDefaultPolarityTests
{
    /// <summary>
    /// (Name, OOXML default, expected written value when the worksheet property is set to the
    /// non-default value below.) Mirrors <see cref="XLibur.Excel.XLViewProperties"/>'s polarity
    /// column for the nine boolean sheet-view attributes.
    /// </summary>
    private static readonly (string Attribute, bool OoxmlDefault, Action<IXLWorksheet> SetNonDefault)[]
        BooleanAttributes =
        [
            ("showFormulas", false, ws => ws.ShowFormulas = true),
            ("showGridLines", true, ws => ws.ShowGridLines = false),
            ("showOutlineSymbols", true, ws => ws.ShowOutlineSymbols = false),
            ("showRowColHeaders", true, ws => ws.ShowRowColHeaders = false),
            ("showRuler", true, ws => ws.ShowRuler = false),
            ("showWhiteSpace", true, ws => ws.ShowWhiteSpace = false),
            ("showZeros", true, ws => ws.ShowZeros = false),
            ("rightToLeft", false, ws => ws.RightToLeft = true),
            ("tabSelected", false, ws => ws.TabSelected = true),
        ];

    [Test]
    public async Task Default_valued_attributes_are_omitted()
    {
        var sheetView = SheetViewTag(SaveDefaultWorksheet());

        foreach (var (attribute, _, _) in BooleanAttributes)
            await Assert.That(Attribute(sheetView, attribute)).IsNull()
                .Because($"{attribute} holds its OOXML default and should be omitted");
    }

    [Test]
    public async Task Non_default_attributes_are_written_with_the_non_default_value()
    {
        var sheetView = SheetViewTag(SaveNonDefaultWorksheet());

        foreach (var (attribute, ooxmlDefault, _) in BooleanAttributes)
        {
            // Bool attributes serialise as "1"/"0"; every one of these was flipped away from its
            // OOXML default, so the written value must be the opposite of that default.
            var expected = ooxmlDefault ? "0" : "1";
            await Assert.That(Attribute(sheetView, attribute)).IsEqualTo(expected)
                .Because($"{attribute}'s default is {ooxmlDefault}; setting it non-default should write \"{expected}\"");
        }
    }

    private static MemoryStream SaveDefaultWorksheet()
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("S");
            wb.SaveAs(ms);
        }

        return ms;
    }

    private static MemoryStream SaveNonDefaultWorksheet()
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            foreach (var (_, _, setNonDefault) in BooleanAttributes)
                setNonDefault(ws);

            wb.SaveAs(ms);
        }

        return ms;
    }

    private static string SheetViewTag(MemoryStream package)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        var entry = archive.Entries.First(e =>
            e.FullName.Equals("xl/worksheets/sheet1.xml", StringComparison.OrdinalIgnoreCase));

        using var reader = new StreamReader(entry.Open());
        var xml = reader.ReadToEnd();

        var match = Regex.Match(xml, "<(?:[A-Za-z_][\\w.-]*:)?sheetView\\b[^>]*>");
        return match.Success ? match.Value : string.Empty;
    }

    /// <summary>The attribute's value, or <c>null</c> when the attribute is absent.</summary>
    private static string? Attribute(string tag, string name)
    {
        var match = Regex.Match(tag, $"\\b{name}=\"([^\"]*)\"");
        return match.Success ? match.Groups[1].Value : null;
    }
}
