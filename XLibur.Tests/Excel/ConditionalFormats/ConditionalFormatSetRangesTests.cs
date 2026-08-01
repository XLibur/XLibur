using System;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.ConditionalFormats;

/// <summary>
/// Replacing a rule's coverage wholesale — the public way to widen one rule over blocks generated
/// from it, instead of leaving a copied rule per block.
/// </summary>
public class ConditionalFormatSetRangesTests
{
    private static IXLConditionalFormat FormatOver(IXLWorksheet sheet, string address)
    {
        var format = sheet.Range(address).AddConditionalFormat();
        format.WhenEquals("1").Fill.SetBackgroundColor(XLColor.Red);
        return format;
    }

    [Test]
    public async Task SetRanges_replaces_the_coverage()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");
        var format = FormatOver(sheet, "A1:B2");

        format.SetRanges([sheet.Range("A1:B2"), sheet.Range("A5:B6")]);

        var covered = format.Ranges.Select(r => r.RangeAddress.ToString()).ToList();
        await Assert.That(covered).IsEquivalentTo(new[] { "A1:B2", "A5:B6" });
    }

    [Test]
    public async Task SetRanges_discards_what_the_rule_covered_before()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");
        var format = FormatOver(sheet, "A1:B2");

        format.SetRanges([sheet.Range("D1:D4")]);

        var covered = format.Ranges.Select(r => r.RangeAddress.ToString()).ToList();
        await Assert.That(covered).IsEquivalentTo(new[] { "D1:D4" });
    }

    [Test]
    public async Task Mutating_what_Ranges_returns_does_nothing()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");
        var format = FormatOver(sheet, "A1:B2");

        // The reason SetRanges has to exist: Ranges is a fresh projection each call.
        format.Ranges.Add(sheet.Range("D1:D4"));

        await Assert.That(format.Ranges.Count).IsEqualTo(1);
    }

    [Test]
    public async Task SetRanges_returns_the_format_for_chaining()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");
        var format = FormatOver(sheet, "A1:B2");

        await Assert.That(format.SetRanges([sheet.Range("A1:A2")])).IsSameReferenceAs(format);
    }

    [Test]
    public async Task SetRanges_rejects_a_range_from_another_worksheet()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");
        var other = wb.AddWorksheet("Sheet2");
        var format = FormatOver(sheet, "A1:B2");

        // Coverage is stored as bare rectangles against the rule's own sheet, so accepting this
        // would silently move the rule rather than extending it.
        await Assert.That(() => format.SetRanges([other.Range("A1:B2")]))
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task SetRanges_rejects_an_empty_set()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");
        var format = FormatOver(sheet, "A1:B2");

        await Assert.That(() => format.SetRanges([])).Throws<ArgumentException>();
    }

    [Test]
    public async Task SetRanges_rejects_null()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");
        var format = FormatOver(sheet, "A1:B2");

        await Assert.That(() => format.SetRanges(null!)).Throws<ArgumentNullException>();
    }

    [Test]
    public async Task Coverage_survives_a_save_and_load()
    {
        using var stream = new System.IO.MemoryStream();

        using (var wb = new XLWorkbook())
        {
            var sheet = wb.AddWorksheet("Sheet1");
            var format = FormatOver(sheet, "A1:B2");
            format.SetRanges([sheet.Range("A1:B2"), sheet.Range("A5:B6")]);
            wb.SaveAs(stream);
        }

        stream.Position = 0;
        using var loaded = new XLWorkbook(stream);

        var covered = loaded.Worksheet("Sheet1").ConditionalFormats
            .Single()
            .Ranges.Select(r => r.RangeAddress.ToString())
            .ToList();

        await Assert.That(covered).IsEquivalentTo(new[] { "A1:B2", "A5:B6" });
    }
}
