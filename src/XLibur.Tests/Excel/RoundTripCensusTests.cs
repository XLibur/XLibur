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
        ["TryToLoad.LO.xlsm.pass.CVE-2016-0122-1.xlsm"] = "an encrypted package, which is a compound file rather than a zip",
        ["TryToLoad.LO.xlsx.tdf100709.xlsx"] = "a package with no workbook part",
        ["TryToLoad.SheetsWithoutRelId.xlsx"] = "two sheets with an empty r:id, so their parts, and the pivot cache only they use, cannot be reached",
    };

    public static IEnumerable<string> Fixtures() =>
        TestHelper.ListResourceFiles(IsWorkbook).Where(f => !Unreadable.ContainsKey(f));

    [Test]
    public async Task The_census_reports_a_deleted_element_and_an_attribute_changed_to_its_default()
    {
        var qualified = new HashSet<(string, string)> { ("spreadsheetml.worksheet", "col@customWidth=false") };
        using var original = TestHelper.OpenResource(SeedFixture);
        var before = RoundTripCensus.Of(original, qualified);

        using var damaged = new MemoryStream();
        original.Position = 0;
        original.CopyTo(damaged);
        using (var zip = new ZipArchive(damaged, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!;
            XDocument sheet;
            using (var stream = entry.Open())
                sheet = XDocument.Load(stream);

            var main = sheet.Root!.Name.Namespace;
            sheet.Descendants(main + "mergeCell").First().Remove();
            sheet.Descendants(main + "col").First(c => (string?)c.Attribute("customWidth") == "1").SetAttributeValue("customWidth", "0");

            entry.Delete();
            using var rewritten = zip.CreateEntry("xl/worksheets/sheet1.xml").Open();
            sheet.Save(rewritten);
        }

        damaged.Position = 0;
        var losses = RoundTripCensus.Losses(before, RoundTripCensus.Of(damaged, qualified)).Select(l => l.ToString());

        await Assert.That(losses).IsEquivalentTo(new[]
        {
            "spreadsheetml.worksheet col@customWidth: 7 -> 6",
            "spreadsheetml.worksheet mergeCell: 9 -> 8",
            "spreadsheetml.worksheet mergeCell@ref: 9 -> 8",
        });
    }

    [Test]
    public async Task Saving_a_fixture_drops_nothing_the_allow_list_does_not_explain()
    {
        var unexplained = new StringBuilder();
        var explained = new HashSet<(string, string)>();
        foreach (var fixture in Fixtures())
        {
            using var original = TestHelper.OpenResource(fixture);
            var before = TakeCensus(original);
            foreach (var edit in new[] { false, true })
            {
                var after = TakeCensus(Save(original, edit));
                foreach (var loss in RoundTripCensus.Losses(before, after))
                {
                    if (RoundTripCensusAllowList.Explains(loss.ContentType, loss.Key))
                        explained.Add((loss.ContentType, RoundTripCensusAllowList.Normalise(loss.Key)));
                    else
                        unexplained.AppendLine($"{fixture} ({(edit ? "edited" : "untouched")}): {loss}");
                }
            }
        }

        var stale = RoundTripCensusAllowList.Keys.Where(k => !explained.Contains(k)).Select(k => $"{k.ContentType} {k.Key}");

        await Assert.That(unexplained.ToString()).IsEmpty();
        await Assert.That(stale).IsEmpty().Because("an allow-list entry no fixture needs any more should be removed");
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
