using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.IO;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// The load and save take shortcuts for the common shape of a worksheet part and a package, and
/// fall back to the general path for anything else. These tests hold the shortcuts to the same
/// results as the general path, and check the fallbacks are taken when they should be.
/// </summary>
public class LoadSaveFastPathTests
{
    private const string MalformedTarget = "http://exa mple.com/";

    #region Worksheet part read once

    [Test]
    public async Task A_sheet_loads_the_same_whether_or_not_its_part_can_be_split_around_the_cells()
    {
        using var package = SaveStructuredWorkbook();

        // A comment ahead of <sheetData> is legal, but it is where a byte search for the element
        // could be fooled, so SheetDataLocator refuses the part and the load streams it instead.
        using var withComment = Copy(package)
            .RewriteSheet1(xml => xml.Replace("<x:sheetData>", "<!-- <x:sheetData> --><x:sheetData>"));

        using var split = new XLWorkbook(package);
        using var streamed = new XLWorkbook(withComment);

        var splitSheet = (XLWorksheet)split.Worksheet(1);
        var streamedSheet = (XLWorksheet)streamed.Worksheet(1);
        await Assert.That(splitSheet.TakePartWithoutSheetData()).IsNotNull();
        await Assert.That(streamedSheet.TakePartWithoutSheetData()).IsNull();

        foreach (var address in new[] { "A1", "B2", "C3", "D4", "B7" })
        {
            var expected = streamed.Worksheet(1).Cell(address);
            var actual = split.Worksheet(1).Cell(address);
            await Assert.That(actual.Value).IsEqualTo(expected.Value);
            await Assert.That(actual.Style.Font.Bold).IsEqualTo(expected.Style.Font.Bold);
            await Assert.That(actual.Style.Fill.BackgroundColor).IsEqualTo(expected.Style.Fill.BackgroundColor);
            await Assert.That(actual.HasHyperlink).IsEqualTo(expected.HasHyperlink);
        }

        await Assert.That(splitSheet.Column(2).Style.Font.Italic).IsTrue();
        await Assert.That(splitSheet.MergedRanges.Count).IsEqualTo(streamedSheet.MergedRanges.Count);
        await Assert.That(splitSheet.PageSetup.PageOrientation).IsEqualTo(XLPageOrientation.Landscape);
    }

    [Test]
    public async Task A_split_part_saves_the_same_as_a_streamed_one()
    {
        using var package = SaveStructuredWorkbook();
        using var withComment = Copy(package)
            .RewriteSheet1(xml => xml.Replace("<x:sheetData>", "<!-- c --><x:sheetData>"));

        using var fromSplit = new MemoryStream();
        using (var wb = new XLWorkbook(package))
            wb.SaveAs(fromSplit);

        using var fromStreamed = new MemoryStream();
        using (var wb = new XLWorkbook(withComment))
            wb.SaveAs(fromStreamed);

        // The streamed input carried a comment, which its save rightly keeps; nothing else differs.
        await Assert.That(fromSplit.Sheet1Xml()).IsEqualTo(fromStreamed.Sheet1Xml().Replace("<!-- c -->", ""));
    }

    [Test]
    public async Task A_prefix_declared_on_sheetData_itself_is_kept_for_the_markup_around_the_cells()
    {
        using var package = SaveStructuredWorkbook();

        // Legal and rare: the root uses the default namespace, and <sheetData> declares the prefix
        // it is written with. Cutting the cells out must not cut that declaration with them.
        using var declaredOnSheetData = Copy(package).RewriteSheet1(PrefixDeclaredOnSheetDataOnly);

        using var plain = new XLWorkbook(package);
        using var declared = new XLWorkbook(declaredOnSheetData);

        var declaredSheet = (XLWorksheet)declared.Worksheet(1);
        await Assert.That(declaredSheet.TakePartWithoutSheetData()).IsNotNull();

        foreach (var address in new[] { "A1", "B2", "C3", "D4", "B7" })
        {
            var expected = plain.Worksheet(1).Cell(address);
            var actual = declared.Worksheet(1).Cell(address);
            await Assert.That(actual.Value).IsEqualTo(expected.Value);
            await Assert.That(actual.Style.Font.Bold).IsEqualTo(expected.Style.Font.Bold);
            await Assert.That(actual.HasHyperlink).IsEqualTo(expected.HasHyperlink);
        }

        await Assert.That(declaredSheet.Column(2).Style.Font.Italic).IsTrue();
        await Assert.That(declaredSheet.MergedRanges.Count).IsEqualTo(1);
        await Assert.That(declaredSheet.PageSetup.PageOrientation).IsEqualTo(XLPageOrientation.Landscape);
    }

    [Test]
    public async Task A_prefix_declared_on_sheetData_itself_saves_a_loadable_part()
    {
        using var package = SaveStructuredWorkbook();
        using var declaredOnSheetData = Copy(package).RewriteSheet1(PrefixDeclaredOnSheetDataOnly);

        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(declaredOnSheetData))
            wb.SaveAs(saved);

        using var reloaded = new XLWorkbook(saved);
        var sheet = reloaded.Worksheet(1);
        await Assert.That(sheet.Cell("A1").Value).IsEqualTo("Header");
        await Assert.That(sheet.Cell("B2").Value).IsEqualTo(2.5);
        await Assert.That(sheet.MergedRanges.Count).IsEqualTo(1);
        await Assert.That(sheet.PageSetup.PageOrientation).IsEqualTo(XLPageOrientation.Landscape);
    }

    [Test]
    public async Task The_part_kept_from_the_load_serves_one_save_only()
    {
        using var package = SaveStructuredWorkbook();
        using var wb = new XLWorkbook(package);
        var ws = (XLWorksheet)wb.Worksheet(1);

        using var first = new MemoryStream();
        wb.SaveAs(first);

        // The save has taken it: the next save starts from the part the first one wrote.
        await Assert.That(ws.TakePartWithoutSheetData()).IsNull();

        ws.PageSetup.PageOrientation = XLPageOrientation.Portrait;
        wb.Worksheet(1).Cell("E5").Value = 42;
        using var second = new MemoryStream();
        wb.SaveAs(second);

        using var reloaded = new XLWorkbook(second);
        var sheet = reloaded.Worksheet(1);
        await Assert.That(sheet.PageSetup.PageOrientation).IsEqualTo(XLPageOrientation.Portrait);
        await Assert.That(sheet.Cell("E5").Value).IsEqualTo(42);
        await Assert.That(sheet.Cell("B2").Value).IsEqualTo(2.5);
    }

    [Test]
    public async Task The_part_kept_from_the_load_is_dropped_when_the_sheet_points_at_another_part()
    {
        using var package = SaveStructuredWorkbook();
        using var wb = new XLWorkbook(package);
        var ws = (XLWorksheet)wb.Worksheet(1);

        ws.RelId = "rIdSomethingElse";

        await Assert.That(ws.TakePartWithoutSheetData()).IsNull();
    }

    #endregion Worksheet part read once

    #region Read-only package open

    [Test]
    public async Task A_package_is_opened_directly_unless_a_relationship_needs_rewriting()
    {
        using var plain = SaveStructuredWorkbook();
        using var malformed = WithMalformedHyperlink(SaveStructuredWorkbook());

        using (var package = ReadOnlyPackageOpener.TryOpen(plain))
            await Assert.That(package).IsNotNull();

        using (var package = ReadOnlyPackageOpener.TryOpen(malformed))
            await Assert.That(package).IsNull();

        using (var package = ReadOnlyPackageOpener.TryOpen(new MemoryStream(new byte[] { 1, 2, 3, 4, 5 })))
            await Assert.That(package).IsNull();
    }

    [Test]
    public async Task A_malformed_hyperlink_target_still_loads_from_a_stream_and_a_file()
    {
        await Assert.That(Uri.TryCreate(MalformedTarget, UriKind.RelativeOrAbsolute, out _)).IsFalse();

        using var malformed = WithMalformedHyperlink(SaveStructuredWorkbook());

        using (var fromStream = new XLWorkbook(new MemoryStream(malformed.ToArray())))
            await Assert.That(fromStream.Worksheet(1).Cell("B7").HasHyperlink).IsTrue();

        using var file = new TemporaryFile();
        File.WriteAllBytes(file.Path, malformed.ToArray());
        using (var fromFile = new XLWorkbook(file.Path))
            await Assert.That(fromFile.Worksheet(1).Cell("B7").HasHyperlink).IsTrue();
    }

    [Test]
    public async Task Loading_from_a_read_only_stream_leaves_it_usable_for_a_later_save()
    {
        using var package = SaveStructuredWorkbook();
        using var readOnly = new MemoryStream(package.ToArray(), writable: false);

        using var wb = new XLWorkbook(readOnly);
        wb.Properties.Title = "Amended";
        using var output = new MemoryStream();
        wb.SaveAs(output);

        using var reloaded = new XLWorkbook(output);
        await Assert.That(reloaded.Properties.Title).IsEqualTo("Amended");
        await Assert.That(reloaded.Worksheet(1).Cell("B2").Value).IsEqualTo(2.5);
    }

    #endregion Read-only package open

    #region Cells written as raw markup

    [Test]
    public async Task Value_cells_are_written_exactly_as_the_writer_would_write_them()
    {
        using var output = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").Value = 1.5;
            ws.Cell("B1").Value = true;
            ws.Cell("C1").Value = XLError.NoValueAvailable;
            ws.Cell("D1").Value = "shared";
            ws.Cell("E1").Value = new DateTime(2026, 1, 2);
            ws.Cell("F1").Value = TimeSpan.FromHours(6);
            ws.Cell("G1").Value = -1E+20;
            wb.SaveAs(output);
        }

        var xml = output.Sheet1Xml();
        await Assert.That(xml).Contains("<x:c r=\"A1\" s=\"0\"><x:v>1.5</x:v></x:c>");
        await Assert.That(xml).Contains("<x:c r=\"B1\" s=\"0\" t=\"b\"><x:v>1</x:v></x:c>");
        await Assert.That(xml).Contains("<x:c r=\"C1\" s=\"0\" t=\"e\"><x:v>#N/A</x:v></x:c>");
        await Assert.That(xml).Contains("<x:c r=\"D1\" s=\"0\" t=\"s\"><x:v>0</x:v></x:c>");
        await Assert.That(xml).Contains("<x:c r=\"E1\" s=\"1\"><x:v>46024</x:v></x:c>");
        await Assert.That(xml).Contains("<x:c r=\"F1\" s=\"2\"><x:v>0.25</x:v></x:c>");
        await Assert.That(xml).Contains("<x:c r=\"G1\" s=\"0\"><x:v>-1E+20</x:v></x:c>");
    }

    [Test]
    public async Task Cells_written_as_markup_and_through_the_writer_keep_their_order()
    {
        using var output = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").Value = 1;
            ws.Cell("B1").FormulaA1 = "A1+1";
            ws.Cell("C1").Value = 3;
            ws.Cell("D1").Value = "inline";
            ws.Cell("D1").ShareString = false;
            ws.Cell("E1").Value = 5;
            ws.Cell("F1").Style.Fill.BackgroundColor = XLColor.Red;
            ws.Cell("A2").Value = 6;
            wb.SaveAs(output);
        }

        var xml = output.Sheet1Xml();
        var order = new[] { "r=\"A1\"", "r=\"B1\"", "r=\"C1\"", "r=\"D1\"", "r=\"E1\"", "r=\"F1\"", "r=\"2\"", "r=\"A2\"" }
            .Select(attribute => xml.IndexOf(attribute, StringComparison.Ordinal)).ToArray();

        await Assert.That(order.All(i => i >= 0)).IsTrue();
        await Assert.That(order.SequenceEqual(order.Order())).IsTrue();
        await Assert.That(xml).Contains("<x:c r=\"D1\" s=\"0\" t=\"inlineStr\"><x:is><x:t>inline</x:t></x:is></x:c>");

        using var reloaded = new XLWorkbook(output);
        var sheet = reloaded.Worksheet(1);
        await Assert.That(sheet.Cell("B1").FormulaA1).IsEqualTo("A1+1");
        await Assert.That(sheet.Cell("E1").Value).IsEqualTo(5);
        await Assert.That(sheet.Cell("A2").Value).IsEqualTo(6);
    }

    #endregion Cells written as raw markup

    [Test]
    public async Task Disposing_a_workbook_drops_its_cells()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.Cell("A1").Value = "text";
        ws.Cell("B2").Value = 2;

        wb.Dispose();

        await Assert.That(((XLWorksheet)ws).Internals.CellsCollection.IsEmpty).IsTrue();
    }

    /// <summary>
    /// A workbook whose sheet carries structural elements on both sides of <c>&lt;sheetData&gt;</c>:
    /// column styles before it, and merges, a hyperlink and page setup after it.
    /// </summary>
    private static MemoryStream SaveStructuredWorkbook()
    {
        var output = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Column(2).Style.Font.Italic = true;
            ws.Cell("A1").Value = "Header";
            ws.Cell("A1").Style.Font.Bold = true;
            ws.Cell("B2").Value = 2.5;
            ws.Cell("C3").Value = new DateTime(2026, 9, 27);
            ws.Cell("D4").Value = true;
            ws.Cell("D4").Style.Fill.BackgroundColor = XLColor.Yellow;
            ws.Range("A5:C5").Merge();
            ws.Cell("B7").Value = "link";
            ws.Cell("B7").SetHyperlink(new XLHyperlink("http://example.com/"));
            ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
            wb.SaveAs(output);
        }

        output.Position = 0;
        return output;
    }

    /// <summary>
    /// Rewrites a sheet XLibur saved (every element prefixed <c>x:</c>, declared on the root) so the
    /// root and everything outside the cells use the default namespace, and <c>&lt;x:sheetData&gt;</c>
    /// declares <c>xmlns:x</c> for itself.
    /// </summary>
    private static string PrefixDeclaredOnSheetDataOnly(string xml)
    {
        const string mainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var sheetDataStart = xml.IndexOf("<x:sheetData>", StringComparison.Ordinal);
        var sheetDataEnd = xml.IndexOf("</x:sheetData>", StringComparison.Ordinal) + "</x:sheetData>".Length;
        if (sheetDataStart < 0 || sheetDataEnd < sheetDataStart)
            throw new InvalidOperationException("The saved sheet has no <x:sheetData> element to rewrite.");

        static string Unprefixed(string markup) => markup
            .Replace("<x:", "<", StringComparison.Ordinal)
            .Replace("</x:", "</", StringComparison.Ordinal);

        var before = Unprefixed(xml[..sheetDataStart])
            .Replace($"xmlns:x=\"{mainNs}\"", $"xmlns=\"{mainNs}\"", StringComparison.Ordinal);
        var cells = xml[sheetDataStart..sheetDataEnd]
            .Replace("<x:sheetData>", $"<x:sheetData xmlns:x=\"{mainNs}\">", StringComparison.Ordinal);
        var after = Unprefixed(xml[sheetDataEnd..]);

        if (!before.Contains($"xmlns=\"{mainNs}\"", StringComparison.Ordinal))
            throw new InvalidOperationException("The saved sheet does not declare the x prefix on its root.");

        return before + cells + after;
    }

    /// <summary>An expandable copy, which a package rewrite needs.</summary>
    private static MemoryStream Copy(MemoryStream package)
    {
        var copy = new MemoryStream();
        copy.Write(package.GetBuffer(), 0, (int)package.Length);
        copy.Position = 0;
        return copy;
    }

    private static MemoryStream WithMalformedHyperlink(MemoryStream package)
    {
        var rewritten = package.RewritePart("xl/worksheets/_rels/sheet1.xml.rels",
            xml => xml.Replace("http://example.com/", MalformedTarget));
        rewritten.Position = 0;
        return rewritten;
    }
}
