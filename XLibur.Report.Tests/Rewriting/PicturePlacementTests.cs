using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Report.Tests.Rewriting;

/// <summary>
/// A picture in a template has to end up beside the same content after generation as before it.
/// </summary>
/// <remarks>
/// There is no picture code in the rewriter: a picture anchor is a live range, so expansion moves
/// it the way it moves everything else. These tests are what says so — the behaviour is inherited
/// rather than written, and inherited behaviour is the kind that disappears quietly.
/// </remarks>
public class PicturePlacementTests
{
    private const string OnePixelPng =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

    private static MemoryStream Image() => new(Convert.FromBase64String(OnePixelPng));

    private static List<SaleItem> Items(int count) => Enumerable.Range(1, count)
        .Select(i => new SaleItem { Product = "Product " + i, Quantity = i })
        .ToList();

    /// <summary>Rows 3 and 4 are the bound range; the picture is anchored at A8, below it.</summary>
    private static XLWorkbook Template(string anchor)
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");

        sheet.Cell("A2").Value = "Product";
        sheet.Cell("A3").Value = "{{ item.Product }}";
        sheet.Cell("B3").Value = "{{ item.Quantity }}";

        using var image = Image();
        sheet.AddPicture(image, "Logo").MoveTo(sheet.Cell(anchor));

        workbook.DefinedNames.Add("Items", sheet.Range("A3:B4"));
        return workbook;
    }

    private static void Generate(IXLWorkbook workbook, int itemCount)
    {
        using var template = new XLTemplate(workbook);
        template.AddVariable("Items", Items(itemCount));
        template.Generate();
    }

    [Test]
    public async Task APictureBelowTheRangeMovesBelowTheGeneratedRows()
    {
        using var workbook = Template("A8");

        Generate(workbook, itemCount: 4);

        // Three rows were inserted for the extra items and the empty options row was removed.
        await Assert.That(workbook.Worksheet("Report").Pictures.Single().TopLeftCell.Address.RowNumber)
            .IsEqualTo(10);
    }

    [Test]
    public async Task APictureAboveTheRangeStaysWhereItIs()
    {
        using var workbook = Template("A1");

        Generate(workbook, itemCount: 4);

        await Assert.That(workbook.Worksheet("Report").Pictures.Single().TopLeftCell.Address.RowNumber)
            .IsEqualTo(1);
    }

    /// <summary>An empty collection removes rows, and the picture comes back up with them.</summary>
    [Test]
    public async Task APictureBelowAnEmptyRangeMovesUp()
    {
        using var workbook = Template("A8");

        Generate(workbook, itemCount: 0);

        await Assert.That(workbook.Worksheet("Report").Pictures.Single().TopLeftCell.Address.RowNumber)
            .IsEqualTo(6);
    }

    [Test]
    public async Task TheMovedAnchorSurvivesASaveAndReload()
    {
        using var generated = new MemoryStream();

        using (var workbook = Template("A8"))
        {
            Generate(workbook, itemCount: 4);
            workbook.SaveAs(generated);
        }

        generated.Position = 0;
        using var reloaded = new XLWorkbook(generated);
        await Assert.That(reloaded.Worksheet("Report").Pictures.Single().TopLeftCell.Address.RowNumber)
            .IsEqualTo(10);
    }
}
