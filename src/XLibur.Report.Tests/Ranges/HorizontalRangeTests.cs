using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Report.Tests.Ranges;

/// <summary>
/// A range that repeats across: one column per item instead of one row.
/// </summary>
public class HorizontalRangeTests
{
    private static List<SaleItem> Items() => new()
    {
        new() { Product = "Widget", Quantity = 2, Region = "North" },
        new() { Product = "Gadget", Quantity = 5, Region = "South" },
        new() { Product = "Doohickey", Quantity = 1, Region = "North" },
    };

    /// <summary>
    /// A range over B4:C6 with the labels down column A. Column B repeats — product, quantity and a
    /// spare line — and column C is the options column.
    /// </summary>
    private static XLWorkbook Template(string c4 = "<<Horizontal>>", string c5 = "", string c6 = "")
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");

        sheet.Cell("A4").Value = "Product";
        sheet.Cell("A5").Value = "Quantity";
        sheet.Cell("A6").Value = "Region";

        sheet.Cell("B4").Value = "{{ item.Product }}";
        sheet.Cell("B5").Value = "{{ item.Quantity }}";
        sheet.Cell("B6").Value = "{{ item.Region }}";

        foreach (var (address, text) in new[] { ("C4", c4), ("C5", c5), ("C6", c6) })
        {
            if (text.Length > 0)
            {
                sheet.Cell(address).Value = text;
            }
        }

        workbook.DefinedNames.Add("Items", sheet.Range("B4:C6"));
        return workbook;
    }

    private static XLGenerateResult Generate(IXLWorkbook workbook, List<SaleItem>? items = null)
    {
        using var template = new XLTemplate(workbook);
        template.AddVariable("Items", items ?? Items());
        return template.Generate();
    }

    private static List<string> Row(IXLWorksheet sheet, int row, string firstColumn, int count) =>
        Enumerable.Range(XLHelper.GetColumnNumberFromLetter(firstColumn), count)
            .Select(column => sheet.Cell(row, column).Value.ToString() ?? string.Empty)
            .ToList();

    [Test]
    public async Task EachItemGetsAColumn()
    {
        using var workbook = Template();

        Generate(workbook);

        await Assert.That(Row(workbook.Worksheet("Report"), 4, "B", 3))
            .IsEquivalentTo(new[] { "Widget", "Gadget", "Doohickey" });
    }

    [Test]
    public async Task EveryLineOfTheItemIsWritten()
    {
        using var workbook = Template();

        Generate(workbook);

        var sheet = workbook.Worksheet("Report");
        await Assert.That(Row(sheet, 5, "B", 3)).IsEquivalentTo(new[] { "2", "5", "1" });
        await Assert.That(Row(sheet, 6, "B", 3)).IsEquivalentTo(new[] { "North", "South", "North" });
    }

    [Test]
    public async Task TheLabelsBesideTheRangeStayPut()
    {
        using var workbook = Template();

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").Cell("A4").Value.GetText()).IsEqualTo("Product");
    }

    [Test]
    public async Task TheHorizontalTagTextNeverReachesTheReport()
    {
        using var workbook = Template();

        Generate(workbook);

        // The options column ended up at E, and nothing was written into it.
        await Assert.That(workbook.Worksheet("Report").Cell("E4").Value.IsBlank).IsTrue();
    }

    /// <summary>An options column holding nothing is dropped, exactly as an options row is.</summary>
    [Test]
    public async Task AnEmptyOptionsColumnIsRemoved()
    {
        using var workbook = Template();

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").RangeUsed()!.RangeAddress.LastAddress.ColumnNumber)
            .IsEqualTo(4);
    }

    [Test]
    public async Task TheDefinedNameIsRePointedAtTheGeneratedColumns()
    {
        using var workbook = Template();

        Generate(workbook);

        await Assert.That(workbook.DefinedName("Items")!.RefersTo).Contains("$B$4:$D$6");
    }

    // ── Tags, across ────────────────────────────────────────────────────

    /// <summary>A summary in a line totals along that line, across the generated columns.</summary>
    [Test]
    public async Task ASummaryTotalsAcrossTheGeneratedColumns()
    {
        using var workbook = Template(c5: "<<Sum>>");

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").Cell("E5").FormulaA1).IsEqualTo("SUBTOTAL(9,B5:D5)");
    }

    [Test]
    public async Task ASummaryCanTotalAnotherLine()
    {
        using var workbook = Template(c4: "<<Horizontal>><<Sum over=5>>");

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").Cell("E4").FormulaA1).IsEqualTo("SUBTOTAL(9,B5:D5)");
    }

    /// <summary>A line named as a column letter means nothing across; the tag says so.</summary>
    [Test]
    public async Task ASummaryOverAColumnLetterIsReported()
    {
        using var workbook = Template(c5: "<<Sum over=B>>");

        var result = Generate(workbook);

        await Assert.That(result.HasErrors).IsTrue();
        await Assert.That(result.ParsingErrors[0].Message).Contains("does not name a row");
    }

    [Test]
    public async Task SortOrdersTheColumnsByTheLineItSitsIn()
    {
        using var workbook = Template(c4: "<<Horizontal>><<Sort>>");

        Generate(workbook);

        await Assert.That(Row(workbook.Worksheet("Report"), 4, "B", 3))
            .IsEquivalentTo(new[] { "Doohickey", "Gadget", "Widget" });
    }

    [Test]
    public async Task HiddenHidesTheRowItSitsIn()
    {
        using var workbook = Template(c6: "<<Hidden>>");

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").Row(6).IsHidden).IsTrue();
    }

    [Test]
    public async Task DeleteRemovesTheRowItSitsIn()
    {
        using var workbook = Template(c6: "<<Delete>>");

        Generate(workbook);

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("B4").Value.GetText()).IsEqualTo("Widget");
        await Assert.That(sheet.Cell("A6").Value.IsBlank).IsTrue();
    }

    /// <summary>A test in a repeated column drops that column, the mirror of dropping a row.</summary>
    [Test]
    public async Task ATestInARepeatedColumnDropsThatColumn()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        sheet.Cell("A4").Value = "Product";
        sheet.Cell("B4").Value = "{{ item.Product }}";
        sheet.Cell("B5").Value = "<<If test=\"item.Quantity > 1\">>";
        sheet.Cell("C4").Value = "<<Horizontal>>";
        workbook.DefinedNames.Add("Items", sheet.Range("B4:C5"));

        Generate(workbook);

        await Assert.That(Row(sheet, 4, "B", 2)).IsEquivalentTo(new[] { "Widget", "Gadget" });
    }

    [Test]
    public async Task AnEmptyCollectionRemovesTheColumns()
    {
        using var workbook = Template();

        Generate(workbook, new List<SaleItem>());

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("B4").Value.IsBlank).IsTrue();
        await Assert.That(sheet.Cell("A4").Value.GetText()).IsEqualTo("Product");
    }

    /// <summary>
    /// Grouping across is not supported and says so, rather than doing something surprising with an
    /// outline nobody asked for.
    /// </summary>
    [Test]
    public async Task GroupingAcrossIsReported()
    {
        using var workbook = Template(c6: "<<Group>>");

        var result = Generate(workbook);

        await Assert.That(result.HasErrors).IsTrue();
        await Assert.That(result.ParsingErrors[0].Message).Contains("not supported in a range that repeats across");

        // Generation carried on regardless, ungrouped.
        await Assert.That(Row(workbook.Worksheet("Report"), 4, "B", 3))
            .IsEquivalentTo(new[] { "Widget", "Gadget", "Doohickey" });
    }

    /// <summary>Excel filters rows, so an autofilter across has nothing to do and says so.</summary>
    [Test]
    public async Task AutoFilterAcrossIsReported()
    {
        using var workbook = Template(c6: "<<AutoFilter>>");

        var result = Generate(workbook);

        await Assert.That(result.HasErrors).IsTrue();
        await Assert.That(result.ParsingErrors[0].Message).Contains("filters rows, not columns");
    }

    // ── Conditional formatting, across ──────────────────────────────────

    [Test]
    public async Task ConditionalFormattingIsStretchedAcrossNotDuplicated()
    {
        using var workbook = Template();
        workbook.Worksheet("Report").Range("B5:B5").AddConditionalFormat()
            .WhenGreaterThan(1).Fill.SetBackgroundColor(XLColor.Red);

        Generate(workbook);

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.ConditionalFormats.Count()).IsEqualTo(1);
        await Assert.That(sheet.ConditionalFormats.Single().Ranges.Single().RangeAddress.ToStringRelative())
            .IsEqualTo("B5:D5");
    }

    // ── The ledger, across ──────────────────────────────────────────────

    /// <summary>Content to the right of a range that grew across moves with it.</summary>
    [Test]
    public async Task ContentToTheRightMovesAcross()
    {
        using var workbook = Template();
        workbook.Worksheet("Report").Cell("F4").Value = "after";

        Generate(workbook);

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("F4").Value.IsBlank).IsTrue();
        await Assert.That(sheet.Cell("G4").Value.GetText()).IsEqualTo("after");
    }

    /// <summary>
    /// A chart plotting the one column a horizontal range repeats has to plot every column it
    /// generated — the mirror of what the rewriter does for a vertical range.
    /// </summary>
    [Test]
    public async Task AChartSeriesIsStretchedAcross()
    {
        using var workbook = Template();
        var sheet = workbook.Worksheet("Report");
        var chart = sheet.Charts.Add(XLChartType.ColumnClustered);
        chart.Series.Add("Quantity", "Report!$B$5:$B$5", "Report!$B$4:$B$4");

        Generate(workbook);

        var series = sheet.Charts.Single().Series.Single();
        await Assert.That(series.ValueReferences).IsEqualTo("Report!$B$5:$D$5");
        await Assert.That(series.CategoryReferences).IsEqualTo("Report!$B$4:$D$4");
    }

    /// <summary>A series to the right of the range moves rather than stretching.</summary>
    [Test]
    public async Task AChartSeriesToTheRightMovesAcross()
    {
        using var workbook = Template();
        var sheet = workbook.Worksheet("Report");
        var chart = sheet.Charts.Add(XLChartType.ColumnClustered);
        chart.Series.Add("Later", "Report!$H$5:$J$5");

        Generate(workbook);

        await Assert.That(sheet.Charts.Single().Series.Single().ValueReferences).IsEqualTo("Report!$I$5:$K$5");
    }

    /// <summary>A series in a row the range does not cross is not stretched by it.</summary>
    [Test]
    public async Task AChartSeriesInAnotherRowIsNotStretched()
    {
        using var workbook = Template();
        var sheet = workbook.Worksheet("Report");
        var chart = sheet.Charts.Add(XLChartType.ColumnClustered);
        chart.Series.Add("Elsewhere", "Report!$B$20:$B$20");

        Generate(workbook);

        await Assert.That(sheet.Charts.Single().Series.Single().ValueReferences).IsEqualTo("Report!$B$20:$B$20");
    }

    [Test]
    public async Task AHorizontalReportSurvivesASaveAndReload()
    {
        using var workbook = Template(c5: "<<Sum>>");

        Generate(workbook);

        using var stream = new System.IO.MemoryStream();
        workbook.SaveAs(stream, validate: true);
        stream.Position = 0;

        using var reloaded = new XLWorkbook(stream);
        var sheet = reloaded.Worksheet("Report");
        await Assert.That(sheet.Cell("D4").Value.GetText()).IsEqualTo("Doohickey");
        await Assert.That(sheet.Cell("E5").FormulaA1).IsEqualTo("SUBTOTAL(9,B5:D5)");
    }
}
