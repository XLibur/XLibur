using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using XLibur.Tests.Utils;
using Hyperlink = DocumentFormat.OpenXml.Spreadsheet.Hyperlink;

namespace XLibur.Tests.Excel.Hyperlinks;

/// <summary>
/// An external hyperlink keeps the relationship ID it was loaded with, so a load and save of an
/// unchanged sheet writes the same bytes (#705).
/// </summary>
public class HyperlinkRelIdTests
{
    private const string Sheet1Rels = "xl/worksheets/_rels/sheet1.xml.rels";

    [Test]
    public async Task External_hyperlink_saves_the_same_sheet_and_rels_on_every_resave()
    {
        using var created = Save(wb =>
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Header";
            ws.Cell("B7").Value = "link";
            ws.Cell("B7").SetHyperlink(new XLHyperlink("http://example.com/"));
        });

        using var first = LoadAndSave(created);
        using var second = LoadAndSave(first);
        using var third = LoadAndSave(second);

        await Assert.That(first.Sheet1Xml()).IsEqualTo(created.Sheet1Xml());
        await Assert.That(first.ReadPart(Sheet1Rels)).IsEqualTo(created.ReadPart(Sheet1Rels));
        await Assert.That(third.Sheet1Xml()).IsEqualTo(first.Sheet1Xml());
        await Assert.That(third.ReadPart(Sheet1Rels)).IsEqualTo(first.ReadPart(Sheet1Rels));
    }

    [Test]
    public async Task Saving_one_workbook_twice_writes_the_same_relationship_id()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("B7").SetHyperlink(new XLHyperlink("http://example.com/"));

        using var first = new MemoryStream();
        wb.SaveAs(first);
        using var second = new MemoryStream();
        wb.SaveAs(second);

        await Assert.That(Links(second)).IsEquivalentTo(Links(first));
    }

    [Test]
    public async Task Changed_address_gets_a_relationship_to_the_new_address()
    {
        using var created = Save(wb =>
            wb.AddWorksheet("Data").Cell("B7").SetHyperlink(new XLHyperlink("http://example.com/")));

        using var output = new MemoryStream();
        using (var wb = new XLWorkbook(new MemoryStream(created.ToArray())))
        {
            wb.Worksheet("Data").Cell("B7").GetHyperlink().ExternalAddress = new Uri("http://example.org/");
            wb.SaveAs(output);
        }

        var links = Links(output);
        await Assert.That(links).HasSingleItem();
        await Assert.That(links[0].Target).IsEqualTo("http://example.org/");
        await Assert.That(output.ReadPart(Sheet1Rels)).DoesNotContain("example.com");
    }

    [Test]
    public async Task Hyperlink_copied_to_another_sheet_gets_its_own_valid_relationship()
    {
        using var created = Save(wb =>
        {
            wb.AddWorksheet("Data").Cell("B7").SetHyperlink(new XLHyperlink("http://example.com/"));
            wb.AddWorksheet("Other");
        });

        using var output = new MemoryStream();
        using (var wb = new XLWorkbook(new MemoryStream(created.ToArray())))
        {
            wb.Worksheet("Data").Cell("B7").CopyTo(wb.Worksheet("Other").Cell("C3"));
            wb.SaveAs(output);
        }

        using var loaded = new XLWorkbook(new MemoryStream(output.ToArray()));
        await Assert.That(loaded.Worksheet("Data").Cell("B7").GetHyperlink().ExternalAddress!.ToString())
            .IsEqualTo("http://example.com/");
        await Assert.That(loaded.Worksheet("Other").Cell("C3").GetHyperlink().ExternalAddress!.ToString())
            .IsEqualTo("http://example.com/");
    }

    [Test]
    public async Task Cells_loaded_from_one_hyperlink_range_each_get_a_valid_relationship()
    {
        // A <hyperlink ref="B7:C7"> is loaded as one hyperlink per cell, both carrying the same r:id.
        using var created = Save(wb =>
            wb.AddWorksheet("Data").Cell("B7").SetHyperlink(new XLHyperlink("http://example.com/")));
        created.RewriteSheet1(xml => xml.Replace("ref=\"B7\"", "ref=\"B7:C7\"", StringComparison.Ordinal));

        using var first = LoadAndSave(created);
        using var second = LoadAndSave(first);

        var links = Links(first);
        await Assert.That(links.Select(l => l.Ref)).IsEquivalentTo(["B7", "C7"]);
        await Assert.That(links.Select(l => l.Id).Distinct().Count()).IsEqualTo(2);
        await Assert.That(links.All(l => l.Target == "http://example.com/")).IsTrue();
        await Assert.That(second.ReadPart(Sheet1Rels)).IsEqualTo(first.ReadPart(Sheet1Rels));
    }

    [Test]
    public async Task Hyperlinks_example_keeps_its_relationship_ids()
    {
        using var input = new MemoryStream();
        using (var resource = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\Misc\Hyperlinks.xlsx")))
            resource.CopyTo(input);

        using var output = LoadAndSave(input);

        var before = Links(input);
        await Assert.That(before.Any(l => l.Id is not null)).IsTrue();
        await Assert.That(Links(output).Where(l => l.Id is not null))
            .IsEquivalentTo(before.Where(l => l.Id is not null));
    }

    private sealed record Link(string Sheet, string Ref, string? Id, string? Target);

    /// <summary>Every hyperlink in the package, with the target its relationship resolves to.</summary>
    private static Link[] Links(Stream package)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        var workbookPart = document.WorkbookPart!;
        return workbookPart.Workbook!.Sheets!.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
            .SelectMany(sheet =>
            {
                var part = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
                return part.Worksheet!.Descendants<Hyperlink>().Select(hl =>
                {
                    var id = hl.Id?.Value;
                    var target = id is null
                        ? null
                        : part.HyperlinkRelationships.Single(r => r.Id == id).Uri.ToString();
                    return new Link(sheet.Name!.Value!, hl.Reference!.Value!, id, target);
                });
            })
            .ToArray();
    }

    private static MemoryStream Save(Action<XLWorkbook> build)
    {
        var output = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            build(wb);
            wb.SaveAs(output);
        }

        output.Position = 0;
        return output;
    }

    private static MemoryStream LoadAndSave(MemoryStream package)
    {
        var output = new MemoryStream();
        using (var wb = new XLWorkbook(new MemoryStream(package.ToArray())))
            wb.SaveAs(output);

        output.Position = 0;
        return output;
    }
}
