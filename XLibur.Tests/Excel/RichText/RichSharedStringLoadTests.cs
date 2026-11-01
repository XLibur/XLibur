using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.RichText;

/// <summary>
/// Cells that show one rich shared string. The load builds that string's rich text once for each
/// cell font it is shown in, and every cell takes the built text, rather than each cell building its
/// own run by run.
/// </summary>
public class RichSharedStringLoadTests
{
    [Test]
    public async Task A_cell_that_shows_a_rich_string_another_cell_showed_costs_little_to_load()
    {
        // The difference between loading 3,000 and 1,000 cells of one rich string is what each
        // further cell costs. Building the rich text again for each cell cost several kilobytes.
        using var small = SharedRichString(cells: 1_000);
        using var large = SharedRichString(cells: 3_000);

        var perCell = (LoadAllocation(large) - LoadAllocation(small)) / 2_000;

        await Assert.That(perCell).IsLessThan(400);
    }

    [Test]
    public async Task Cells_in_different_fonts_each_show_the_string_in_their_own_font()
    {
        // The first run has no rPr and inherits the whole cell font; the second states only bold,
        // and takes the rest of the cell font. So one string reads differently in each font.
        using var source = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            for (var row = 1; row <= 4; row++)
            {
                var cell = ws.Cell(row, 1);
                cell.Style.Font.FontName = row % 2 == 0 ? "Arial" : "Courier New";
                cell.Style.Font.FontSize = row % 2 == 0 ? 14 : 9;
                var richText = cell.GetRichText();
                richText.AddText("plain ");
                richText.AddText("bold").SetBold();
            }

            wb.SaveAs(source);
        }

        using var loaded = new XLWorkbook(source);
        var runs = loaded.Worksheet(1).Column(1).CellsUsed()
            .Select(c => string.Join(" | ", c.GetRichText().Select(r => $"{r.Text.Trim()} {r.FontName} {r.FontSize} {r.Bold}")));

        await Assert.That(string.Join(Environment.NewLine, runs)).IsEqualTo(string.Join(Environment.NewLine,
            "plain Courier New 9 False | bold Courier New 9 True",
            "plain Arial 14 False | bold Arial 14 True",
            "plain Courier New 9 False | bold Courier New 9 True",
            "plain Arial 14 False | bold Arial 14 True"));
    }

    [Test]
    public async Task Editing_one_cell_of_a_shared_rich_string_leaves_the_others_as_they_were()
    {
        using var source = SharedRichString(cells: 3);
        using var wb = new XLWorkbook(source);
        var ws = wb.Worksheet(1);
        ws.Cell("A2").GetRichText().First().SetItalic().Text = "Changed ";

        var texts = ws.Column(1).CellsUsed()
            .Select(c => string.Join("|", c.GetRichText().Select(r => $"{r.Text}{(r.Italic ? "/i" : "")}{(r.Bold ? "/b" : "")}")));

        await Assert.That(string.Join(", ", texts)).IsEqualTo(
            "Hello /b|world, Changed /i/b|world, Hello /b|world");
    }

    /// <summary>A saved workbook whose column A holds <paramref name="cells"/> cells of one rich string.</summary>
    private static MemoryStream SharedRichString(int cells)
    {
        var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            for (var row = 1; row <= cells; row++)
            {
                var richText = ws.Cell(row, 1).GetRichText();
                richText.AddText("Hello ").SetBold();
                richText.AddText("world");
            }

            wb.SaveAs(package);
        }

        package.Position = 0;
        return package;
    }

    /// <summary>The bytes this thread allocates loading <paramref name="package"/>, after one warm-up load.</summary>
    private static long LoadAllocation(MemoryStream package)
    {
        package.Position = 0;
        using (new XLWorkbook(package))
        {
        }

        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        package.Position = 0;
        var before = GC.GetAllocatedBytesForCurrentThread();
        using (new XLWorkbook(package))
        {
            return GC.GetAllocatedBytesForCurrentThread() - before;
        }
    }
}
