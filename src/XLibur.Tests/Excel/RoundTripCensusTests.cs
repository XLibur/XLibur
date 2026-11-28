using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;

namespace XLibur.Tests.Excel;

public class RoundTripCensusTests
{
    private static readonly Dictionary<string, string> Unreadable = new()
    {
        ["Encrypted.Encrypted_XL365.xlsx"] = "an encrypted package, which is a compound file rather than a zip",
        ["TryToLoad.LO.xlsm.pass.CVE-2016-0122-1.xlsm"] = "not a package: its bytes start 19 1E D1 A8, which is neither a zip nor a compound file",
        ["TryToLoad.LO.xlsx.tdf100709.xlsx"] = "a package with no workbook part",
        ["TryToLoad.SheetsWithoutRelId.xlsx"] = "two sheets with an empty r:id, so their parts, and the pivot cache only they use, cannot be reached",
    };

    public static IEnumerable<string> Fixtures() =>
        TestHelper.ListResourceFiles(IsWorkbook).Where(f => !Unreadable.ContainsKey(f));

    [Test]
    public async Task The_census_reports_a_deleted_element_an_attribute_changed_to_its_default_and_a_cell_that_lost_its_style()
    {
        var losses = CensusOfDamage(SeedFixture, "xl/worksheets/sheet1.xml", sheet =>
        {
            var main = sheet.Root!.Name.Namespace;
            sheet.Descendants(main + "mergeCell").First().Remove();
            sheet.Descendants(main + "col").First(c => (string?)c.Attribute("customWidth") == "1").SetAttributeValue("customWidth", "0");
            sheet.Descendants(main + "c").First(c => (string?)c.Attribute("r") == "B2").SetAttributeValue("s", null);
        });

        await Assert.That(losses).IsEquivalentTo(new[]
        {
            "spreadsheetml.worksheet c@s: 21 -> 20",
            "spreadsheetml.worksheet c@s->alignment: 21 -> 20",
            "spreadsheetml.worksheet c@s->border: 21 -> 20",
            "spreadsheetml.worksheet c@s->fill: 21 -> 20",
            "spreadsheetml.worksheet c@s->font: 21 -> 20",
            "spreadsheetml.worksheet col@customWidth: 7 -> 6",
            "spreadsheetml.worksheet mergeCell: 9 -> 8",
            "spreadsheetml.worksheet mergeCell@ref: 9 -> 8",
        });
    }

    [Test]
    public async Task The_census_reports_a_conditional_format_whose_dxf_lost_its_fill_colour()
    {
        var losses = CensusOfDamage(@"Examples\ConditionalFormatting\CFEqualsNumber.xlsx", "xl/styles.xml", styles =>
            styles.Descendants(styles.Root!.Name.Namespace + "bgColor").Single().Remove());

        await Assert.That(losses).IsEquivalentTo(new[]
        {
            "spreadsheetml.styles bgColor: 1 -> 0",
            "spreadsheetml.styles bgColor@rgb: 1 -> 0",
            "spreadsheetml.worksheet cfRule@dxfId->fill: 1 -> 0",
        });
    }

    [Test]
    public async Task Saving_a_fixture_drops_nothing_the_allow_list_does_not_explain()
    {
        var unexplained = new StringBuilder();
        var lost = new Dictionary<(string ContentType, string Key), int>();
        foreach (var fixture in Fixtures())
        {
            using var original = TestHelper.OpenResource(fixture);
            var before = TakeCensus(original);
            foreach (var edit in new[] { false, true })
            {
                var after = TakeCensus(Save(original, edit));
                foreach (var loss in RoundTripCensus.Losses(before, after))
                {
                    var key = (loss.ContentType, RoundTripCensusAllowList.Normalise(loss.Key));
                    if (RoundTripCensusAllowList.Explains(loss.ContentType, loss.Key))
                        lost[key] = lost.GetValueOrDefault(key) + loss.Before - loss.After;
                    else
                        unexplained.AppendLine($"{fixture} ({(edit ? "edited" : "untouched")}): {loss}");
                }
            }
        }

        var stale = RoundTripCensusAllowList.Keys.Where(k => !lost.ContainsKey(k)).Select(k => $"{k.ContentType} {k.Key}");
        var exceeded = RoundTripCensusAllowList.Bounds
            .Where(b => lost.GetValueOrDefault(b.Key) > b.Value)
            .Select(b => $"{b.Key.ContentType} {b.Key.Key}: lost {lost[b.Key]}, bound {b.Value}{Environment.NewLine}");
        foreach (var (key, bound) in RoundTripCensusAllowList.Bounds.Where(b => lost.GetValueOrDefault(b.Key) < b.Value))
            Console.WriteLine($"Bound above the measured loss, lower it: {key.ContentType} {key.Key}: lost {lost.GetValueOrDefault(key)}, bound {bound}");

        using (Assert.Multiple())
        {
            await Assert.That(unexplained.ToString()).IsEmpty();
            await Assert.That(string.Concat(exceeded)).IsEmpty().Because("a bounded allow-list entry covers only as much loss as the corpus had when it was measured");
            await Assert.That(stale).IsEmpty().Because("an allow-list entry no fixture needs any more should be removed");
        }
    }

    /// <summary>The losses the census reports when <paramref name="damage"/> edits one part of a copy of a fixture.</summary>
    private static List<string> CensusOfDamage(string fixture, string part, Action<XDocument> damage)
    {
        var qualified = new HashSet<(string, string)> { ("spreadsheetml.worksheet", "col@customWidth=false") };
        using var original = TestHelper.OpenResource(fixture);
        var before = RoundTripCensus.Of(original, qualified);

        using var damaged = new MemoryStream();
        original.Position = 0;
        original.CopyTo(damaged);
        using (var zip = new ZipArchive(damaged, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry(part)!;
            XDocument document;
            using (var stream = entry.Open())
                document = XDocument.Load(stream);

            damage(document);
            entry.Delete();
            using var rewritten = zip.CreateEntry(part).Open();
            document.Save(rewritten);
        }

        damaged.Position = 0;
        return RoundTripCensus.Losses(before, RoundTripCensus.Of(damaged, qualified)).Select(l => (l with { Example = null }).ToString()).ToList();
    }

    private const string SeedFixture = @"Other\StyleReferenceFiles\MergedCellsBorder\inputfile.xlsx";

    private static bool IsWorkbook(string resource) =>
        resource.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || resource.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase);

    private static RoundTripCensus TakeCensus(Stream package)
    {
        package.Position = 0;
        return RoundTripCensus.Of(package, RoundTripCensusAllowList.Qualified);
    }

    private static MemoryStream Save(MemoryStream original, bool edit)
    {
        using var wb = new XLWorkbook(new MemoryStream(original.ToArray()));
        if (edit)
        {
            foreach (var ws in wb.Worksheets)
                ws.Cell(XLHelper.MaxRowNumber, XLHelper.MaxColumnNumber).Value = 1;
        }

        var saved = new MemoryStream();
        wb.SaveAs(saved);
        return saved;
    }
}
