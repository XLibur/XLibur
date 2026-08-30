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

    /// <summary>
    /// The two tests above only ever write from a freshly-created, in-memory worksheet. On a
    /// re-save of a loaded file, the writer instead mutates the <c>SheetView</c> element the reader
    /// kept a reference to — so a polarity bug that only shows up when an attribute already present
    /// from the load needs to be cleared or overwritten, rather than written for the first time,
    /// would not be caught above. This loads an XLibur-authored fixture, flips every boolean and the
    /// view mode the opposite way, re-saves, and reads the bytes again.
    /// </summary>
    [Test]
    public async Task LoadMutateResave_flip_to_default_omits_every_attribute()
    {
        var fixture = SaveNonDefaultWorksheet(XLSheetViewOptions.PageLayout);

        fixture.Position = 0;
        var resaved = new MemoryStream();
        using (var wb = new XLWorkbook(fixture))
        {
            var ws = wb.Worksheets.First();
            ws.ShowFormulas = false;
            ws.ShowGridLines = true;
            ws.ShowOutlineSymbols = true;
            ws.ShowRowColHeaders = true;
            ws.ShowRuler = true;
            ws.ShowWhiteSpace = true;
            ws.ShowZeros = true;
            ws.RightToLeft = false;
            ws.TabSelected = false;
            ws.SheetView.SetView(XLSheetViewOptions.Normal);
            wb.SaveAs(resaved);
        }

        var sheetView = SheetViewTag(resaved);

        foreach (var (attribute, _, _) in BooleanAttributes)
            await Assert.That(Attribute(sheetView, attribute)).IsNull()
                .Because($"{attribute} was flipped back to its OOXML default after load and should be omitted on resave");

        await Assert.That(Attribute(sheetView, "view")).IsNull()
            .Because("view mode was flipped back to Normal after load and should be omitted on resave");
    }

    [Test]
    public async Task LoadMutateResave_flip_to_non_default_writes_correct_polarity()
    {
        var fixture = SaveDefaultWorksheet();

        fixture.Position = 0;
        var resaved = new MemoryStream();
        using (var wb = new XLWorkbook(fixture))
        {
            var ws = wb.Worksheets.First();
            foreach (var (_, _, setNonDefault) in BooleanAttributes)
                setNonDefault(ws);
            ws.SheetView.SetView(XLSheetViewOptions.PageLayout);
            wb.SaveAs(resaved);
        }

        var sheetView = SheetViewTag(resaved);

        foreach (var (attribute, ooxmlDefault, _) in BooleanAttributes)
        {
            var expected = ooxmlDefault ? "0" : "1";
            await Assert.That(Attribute(sheetView, attribute)).IsEqualTo(expected)
                .Because($"{attribute} was flipped away from default after load; resave should write \"{expected}\"");
        }

        await Assert.That(Attribute(sheetView, "view")).IsEqualTo("pageLayout")
            .Because("view mode was flipped to PageLayout after load and should be written on resave");
    }

    /// <summary>
    /// Excel writes <c>zoomScale</c> alone on a sheet whose other views have never been zoomed, so a
    /// re-save must not fabricate a zoom for a view the file never mentioned. This builds that
    /// shape by stripping the three named scales back out of an XLibur-saved package, loads it, and
    /// re-saves: the page-layout zoom the file does carry must land in
    /// <c>zoomScalePageLayoutView</c> (ECMA-376 18.3.1.87), and <c>zoomScaleSheetLayoutView</c> —
    /// Page Break Preview — must stay absent.
    /// </summary>
    [Test]
    public async Task LoadResave_does_not_invent_a_zoom_for_a_view_the_file_never_zoomed()
    {
        var fixture = SavePageLayoutZoom(140).RewriteSheet1(xml =>
            Regex.Replace(xml, "\\s(?:zoomScaleNormal|zoomScalePageLayoutView|zoomScaleSheetLayoutView)=\"[^\"]*\"", ""));

        // Guard the fixture itself: if the strip stopped matching, the test below would pass for
        // the wrong reason.
        await Assert.That(Attribute(SheetViewTag(fixture), "zoomScale")).IsEqualTo("140");
        await Assert.That(Attribute(SheetViewTag(fixture), "zoomScaleSheetLayoutView")).IsNull();
        await Assert.That(Attribute(SheetViewTag(fixture), "zoomScalePageLayoutView")).IsNull();

        fixture.Position = 0;
        var resaved = new MemoryStream();
        using (var wb = new XLWorkbook(fixture))
            wb.SaveAs(resaved);

        var sheetView = SheetViewTag(resaved);

        await Assert.That(Attribute(sheetView, "zoomScale")).IsEqualTo("140");
        await Assert.That(Attribute(sheetView, "zoomScalePageLayoutView")).IsEqualTo("140")
            .Because("the sheet is in Page Layout view, so its zoom is the page-layout zoom");
        await Assert.That(Attribute(sheetView, "zoomScaleSheetLayoutView")).IsNull()
            .Because("the file never carried a Page Break Preview zoom and a re-save must not invent one");
    }

    private static MemoryStream SavePageLayoutZoom(int zoomScale)
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.SheetView.SetView(XLSheetViewOptions.PageLayout);
            ws.SheetView.ZoomScale = zoomScale;
            wb.SaveAs(ms);
        }

        return ms;
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

    private static MemoryStream SaveNonDefaultWorksheet(XLSheetViewOptions view = XLSheetViewOptions.Normal)
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            foreach (var (_, _, setNonDefault) in BooleanAttributes)
                setNonDefault(ws);
            ws.SheetView.SetView(view);

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
