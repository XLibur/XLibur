using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Report.Tests.Tags;

public class GroupTagTests
{
    /// <summary>
    /// Four sales over two regions and two categories, deliberately shuffled so that a test only
    /// passes if grouping ordered them itself.
    /// </summary>
    private static List<SaleItem> Items() => new()
    {
        new() { Region = "North", Category = "Tools", Product = "Widget", Quantity = 2 },
        new() { Region = "South", Category = "Toys", Product = "Gadget", Quantity = 5 },
        new() { Region = "North", Category = "Toys", Product = "Doohickey", Quantity = 1 },
        new() { Region = "South", Category = "Tools", Product = "Sprocket", Quantity = 3 },
    };

    /// <summary>
    /// Builds a four-column range over A3:D4 — region, category, product, quantity — where row 3 is
    /// the data row and row 4 the options row.
    /// </summary>
    private static IXLWorkbook Template(string a = "", string b = "", string c = "", string d = "")
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");

        sheet.Cell("A2").Value = "Region";
        sheet.Cell("B2").Value = "Category";
        sheet.Cell("C2").Value = "Product";
        sheet.Cell("D2").Value = "Quantity";

        sheet.Cell("A3").Value = "{{ item.Region }}";
        sheet.Cell("B3").Value = "{{ item.Category }}";
        sheet.Cell("C3").Value = "{{ item.Product }}";
        sheet.Cell("D3").Value = "{{ item.Quantity }}";

        foreach (var (column, text) in new[] { ("A", a), ("B", b), ("C", c), ("D", d) })
        {
            if (text.Length > 0)
            {
                sheet.Cell(column + "4").Value = text;
            }
        }

        workbook.DefinedNames.Add("Items", sheet.Range("A3:D4"));
        return workbook;
    }

    private static XLGenerateResult Generate(IXLWorkbook workbook, List<SaleItem>? items = null)
    {
        using var template = new XLTemplate(workbook);
        template.AddVariable("Items", items ?? Items());
        return template.Generate();
    }

    private static List<string> Column(IXLWorksheet sheet, string column, int firstRow, int lastRow) =>
        Enumerable.Range(firstRow, lastRow - firstRow + 1)
            .Select(row => sheet.Cell(column + row).Value.ToString() ?? string.Empty)
            .ToList();

    [Test]
    public async Task GroupOrdersRowsSoThatEachGroupIsContiguous()
    {
        using var workbook = Template(a: "<<Group>>");

        Generate(workbook);

        // Rows 5 and 8 are the subtotal rows the groups produced.
        var sheet = workbook.Worksheet("Report");
        await Assert.That(Column(sheet, "C", 3, 4)).IsEquivalentTo(new[] { "Widget", "Doohickey" });
        await Assert.That(Column(sheet, "C", 6, 7)).IsEquivalentTo(new[] { "Gadget", "Sprocket" });
    }

    [Test]
    public async Task EachGroupGetsASubtotalRowOverItsOwnRows()
    {
        using var workbook = Template(a: "<<Group>>", d: "<<Sum>>");

        Generate(workbook);

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("D5").FormulaA1).IsEqualTo("SUBTOTAL(9,D3:D4)");
        await Assert.That(sheet.Cell("D8").FormulaA1).IsEqualTo("SUBTOTAL(9,D6:D7)");
    }

    [Test]
    public async Task TheSubtotalRowIsLabelledWithItsGroupKey()
    {
        using var workbook = Template(a: "<<Group>>", d: "<<Sum>>");

        Generate(workbook);

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("A5").Value.GetText()).IsEqualTo("North Total");
        await Assert.That(sheet.Cell("A8").Value.GetText()).IsEqualTo("South Total");
    }

    [Test]
    public async Task TheLabelCanBeGiven()
    {
        using var workbook = Template(a: "<<Group totalLabel=\"All of {0}\">>", d: "<<Sum>>");

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").Cell("A5").Value.GetText()).IsEqualTo("All of North");
    }

    /// <summary>
    /// The grand total spans the group totals, and <c>SUBTOTAL</c> ignoring nested
    /// <c>SUBTOTAL</c>s is what keeps that from counting the data twice.
    /// </summary>
    [Test]
    public async Task TheGrandTotalDoesNotCountTheGroupTotalsAgain()
    {
        using var workbook = Template(a: "<<Group>>", d: "<<Sum>>");

        Generate(workbook);

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("D9").FormulaA1).IsEqualTo("SUBTOTAL(9,D3:D8)");

        workbook.RecalculateAllFormulas();
        await Assert.That(sheet.Cell("D9").Value.GetNumber()).IsEqualTo(11d);
        await Assert.That(sheet.Cell("D5").Value.GetNumber()).IsEqualTo(3d);
        await Assert.That(sheet.Cell("D8").Value.GetNumber()).IsEqualTo(8d);
    }

    [Test]
    public async Task DataRowsAreOutlinedBelowTheirSubtotalRow()
    {
        using var workbook = Template(a: "<<Group>>", d: "<<Sum>>");

        Generate(workbook);

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Row(3).OutlineLevel).IsEqualTo(1);
        await Assert.That(sheet.Row(4).OutlineLevel).IsEqualTo(1);
        await Assert.That(sheet.Row(5).OutlineLevel).IsEqualTo(0);
        await Assert.That(sheet.Row(6).OutlineLevel).IsEqualTo(1);
    }

    [Test]
    public async Task SummaryAbovePutsTheSubtotalRowFirst()
    {
        using var workbook = Template(a: "<<Group summaryAbove>>", d: "<<Sum>>");

        Generate(workbook);

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("A3").Value.GetText()).IsEqualTo("North Total");
        await Assert.That(sheet.Cell("D3").FormulaA1).IsEqualTo("SUBTOTAL(9,D4:D5)");
        await Assert.That(sheet.Cell("A6").Value.GetText()).IsEqualTo("South Total");
        await Assert.That(sheet.Outline.SummaryVLocation).IsEqualTo(XLOutlineSummaryVLocation.Top);
    }

    /// <summary>The range-wide form of the same option, for a template with several levels.</summary>
    [Test]
    public async Task SummaryAboveCanBeSetForTheWholeRange()
    {
        using var workbook = Template(a: "<<Group>>", b: "<<SummaryAbove>>", d: "<<Sum>>");

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").Cell("A3").Value.GetText()).IsEqualTo("North Total");
    }

    [Test]
    public async Task GroupsNestLeftmostOutermost()
    {
        using var workbook = Template(a: "<<Group>>", b: "<<Group>>", d: "<<Sum>>");

        Generate(workbook);

        var sheet = workbook.Worksheet("Report");

        // North/Tools, its total, North/Toys, its total, North's total, then the same for South:
        // the inner total comes first so that the totals read outwards from the rows they cover.
        await Assert.That(Column(sheet, "C", 3, 3)).IsEquivalentTo(new[] { "Widget" });
        await Assert.That(sheet.Cell("B4").Value.GetText()).IsEqualTo("Tools Total");
        await Assert.That(sheet.Cell("B6").Value.GetText()).IsEqualTo("Toys Total");
        await Assert.That(sheet.Cell("A7").Value.GetText()).IsEqualTo("North Total");
        await Assert.That(sheet.Cell("A12").Value.GetText()).IsEqualTo("South Total");
    }

    [Test]
    public async Task ANestedSubtotalCoversOnlyItsOwnLevel()
    {
        using var workbook = Template(a: "<<Group>>", b: "<<Group>>", d: "<<Sum>>");

        Generate(workbook);

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("D4").FormulaA1).IsEqualTo("SUBTOTAL(9,D3:D3)");
        await Assert.That(sheet.Cell("D7").FormulaA1).IsEqualTo("SUBTOTAL(9,D3:D6)");

        workbook.RecalculateAllFormulas();
        await Assert.That(sheet.Cell("D7").Value.GetNumber()).IsEqualTo(3d);
        await Assert.That(sheet.Cell("D13").Value.GetNumber()).IsEqualTo(11d);
    }

    [Test]
    public async Task NestedGroupsOutlineOneLevelApart()
    {
        using var workbook = Template(a: "<<Group>>", b: "<<Group>>", d: "<<Sum>>");

        Generate(workbook);

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Row(3).OutlineLevel).IsEqualTo(2);
        await Assert.That(sheet.Row(4).OutlineLevel).IsEqualTo(1);
        await Assert.That(sheet.Row(7).OutlineLevel).IsEqualTo(0);
    }

    [Test]
    public async Task DescendingReversesTheGroupOrder()
    {
        using var workbook = Template(a: "<<Group desc>>", d: "<<Sum>>");

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").Cell("A5").Value.GetText()).IsEqualTo("South Total");
    }

    /// <summary>
    /// Grouping orders rows itself, but stably, so a sort on another column still decides what
    /// happens inside a group.
    /// </summary>
    [Test]
    public async Task SortDecidesTheOrderWithinAGroup()
    {
        using var workbook = Template(a: "<<Group>>", c: "<<Sort>>");

        Generate(workbook);

        await Assert.That(Column(workbook.Worksheet("Report"), "C", 3, 4))
            .IsEquivalentTo(new[] { "Doohickey", "Widget" });
    }

    [Test]
    public async Task NoSortGroupsTheRowsAsTheyCome()
    {
        using var workbook = Template(a: "<<Group nosort>>", d: "<<Sum>>");

        Generate(workbook);

        // Four runs of one row each, because the regions alternate in the source order.
        var sheet = workbook.Worksheet("Report");
        await Assert.That(Column(sheet, "A", 3, 6))
            .IsEquivalentTo(new[] { "North", "North Total", "South", "South Total" });
    }

    [Test]
    public async Task MergeLabelsCollapsesAGroupsRepeatedLabelIntoOne()
    {
        using var workbook = Template(a: "<<Group merge>>", d: "<<Sum>>");

        Generate(workbook);

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.MergedRanges.Count).IsEqualTo(2);
        await Assert.That(sheet.MergedRanges.Select(r => r.RangeAddress.ToStringRelative()))
            .IsEquivalentTo(new[] { "A3:A4", "A6:A7" });
    }

    [Test]
    public async Task DisableSubtotalsOutlinesWithoutAddingRows()
    {
        using var workbook = Template(a: "<<Group disableSubtotals>>", d: "<<Sum>>");

        Generate(workbook);

        var sheet = workbook.Worksheet("Report");
        await Assert.That(Column(sheet, "C", 3, 6))
            .IsEquivalentTo(new[] { "Widget", "Doohickey", "Gadget", "Sprocket" });
        await Assert.That(sheet.Row(3).OutlineLevel).IsEqualTo(1);
        await Assert.That(sheet.Cell("D7").FormulaA1).IsEqualTo("SUBTOTAL(9,D3:D6)");
    }

    [Test]
    public async Task DisableGrandTotalKeepsTheGroupTotalsAndDropsTheReportTotal()
    {
        using var workbook = Template(a: "<<Group>>", b: "<<DisableGrandTotal>>", d: "<<Sum>>");

        Generate(workbook);

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("D5").FormulaA1).IsEqualTo("SUBTOTAL(9,D3:D4)");

        // Nothing was written into the options row, so it was removed with the rest of the tags.
        await Assert.That(sheet.RangeUsed()!.RangeAddress.LastAddress.RowNumber).IsEqualTo(8);
    }

    [Test]
    public async Task PageBreaksSeparateTheGroupsButNotTheLastOne()
    {
        using var workbook = Template(a: "<<Group pageBreaks>>", d: "<<Sum>>");

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").PageSetup.RowBreaks).IsEquivalentTo(new[] { 5 });
    }

    [Test]
    public async Task CollapseHidesTheRowsUnderEachSubtotal()
    {
        using var workbook = Template(a: "<<Group collapse>>", d: "<<Sum>>");

        Generate(workbook);

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Row(3).IsHidden).IsTrue();
        await Assert.That(sheet.Row(4).IsHidden).IsTrue();
        await Assert.That(sheet.Row(5).IsHidden).IsFalse();
    }

    [Test]
    public async Task ASubtotalRowTakesTheOptionsRowsStyling()
    {
        using var workbook = Template(a: "<<Group>>", d: "<<Sum>>");
        workbook.Worksheet("Report").Row(4).Style.Font.Bold = true;

        Generate(workbook);

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("D5").Style.Font.Bold).IsTrue();
        await Assert.That(sheet.Cell("D3").Style.Font.Bold).IsFalse();
    }

    /// <summary>
    /// Grouping inserts rows into a block whose conditional formats expansion already widened, so
    /// the rule count has to survive it — the upstream #216 complaint applies just as much here.
    /// </summary>
    [Test]
    public async Task GroupingDoesNotMultiplyConditionalFormats()
    {
        using var workbook = Template(a: "<<Group>>", d: "<<Sum>>");
        workbook.Worksheet("Report").Range("D3:D3").AddConditionalFormat().WhenGreaterThan(2).Fill.SetBackgroundColor(XLColor.Red);

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").ConditionalFormats.Count()).IsEqualTo(1);
    }

    /// <summary>
    /// Inserting a subtotal row above the first group puts it above everything the range generated,
    /// which is the one place grouping can push a conditional format around instead of through it.
    /// </summary>
    [Test]
    public async Task ConditionalFormatsSurviveASubtotalRowAboveTheBlock()
    {
        using var workbook = Template(a: "<<Group summaryAbove>>", d: "<<Sum>>");
        workbook.Worksheet("Report").Range("D3:D3").AddConditionalFormat().WhenGreaterThan(2).Fill.SetBackgroundColor(XLColor.Red);

        Generate(workbook);

        // Rows 3 and 6 hold the two subtotals, rows 4, 5, 7 and 8 the data. The rule still covers
        // every data row, having been pushed down by the first subtotal row and stretched by the
        // second — and there is still only one of it.
        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.ConditionalFormats.Count()).IsEqualTo(1);
        await Assert.That(sheet.ConditionalFormats.Single().Ranges.Single().RangeAddress.ToStringRelative())
            .IsEqualTo("D4:D8");
    }

    /// <summary>
    /// Two template rows per item, so that every row number grouping works out is a multiple of the
    /// block size rather than of one.
    /// </summary>
    [Test]
    public async Task GroupingHandlesItemsThatTakeMoreThanOneRow()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        sheet.Cell("A1").Value = "{{ item.Region }}";
        sheet.Cell("B1").Value = "{{ item.Quantity }}";
        sheet.Cell("A2").Value = "{{ item.Product }}";
        sheet.Cell("A3").Value = "<<Group>>";
        sheet.Cell("B3").Value = "<<Sum>>";
        workbook.DefinedNames.Add("Items", sheet.Range("A1:B3"));

        Generate(workbook);

        // Two rows per item, two items per region, then the region's subtotal row.
        await Assert.That(Column(sheet, "A", 1, 5))
            .IsEquivalentTo(new[] { "North", "Widget", "North", "Doohickey", "North Total" });
        await Assert.That(sheet.Cell("B5").FormulaA1).IsEqualTo("SUBTOTAL(9,B1:B4)");
        await Assert.That(sheet.Cell("B10").FormulaA1).IsEqualTo("SUBTOTAL(9,B6:B9)");
        await Assert.That(sheet.Cell("B11").FormulaA1).IsEqualTo("SUBTOTAL(9,B1:B10)");
    }

    [Test]
    public async Task GroupByAnExpressionTheReportDoesNotShow()
    {
        using var workbook = Template(c: "<<Group by=item.Region>>", d: "<<Sum>>");

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").Cell("C5").Value.GetText()).IsEqualTo("North Total");
    }

    [Test]
    public async Task GroupWithNothingToGroupByIsReportedWithoutStoppingGeneration()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        sheet.Cell("A1").Value = "fixed text";
        sheet.Cell("A2").Value = "<<Group>>";
        workbook.DefinedNames.Add("Items", sheet.Range("A1:A2"));

        using var template = new XLTemplate(workbook);
        template.AddVariable("Items", Items());
        var result = template.Generate();

        await Assert.That(result.HasErrors).IsTrue();
        await Assert.That(result.ParsingErrors[0].Message).Contains("nothing to group by");
        await Assert.That(sheet.Cell("A1").Value.GetText()).IsEqualTo("fixed text");
    }

    /// <summary>
    /// No items means no groups, and the range is removed exactly as an ungrouped one is.
    /// </summary>
    [Test]
    public async Task AnEmptyCollectionGroupsIntoNothing()
    {
        using var workbook = Template(a: "<<Group>>", d: "<<Sum>>");

        var result = Generate(workbook, new List<SaleItem>());

        var sheet = workbook.Worksheet("Report");
        await Assert.That(result.HasErrors).IsFalse();
        await Assert.That(sheet.RangeUsed()!.RangeAddress.LastAddress.RowNumber).IsEqualTo(2);
    }

    [Test]
    public async Task GroupedRowsSurviveASaveAndReload()
    {
        using var workbook = Template(a: "<<Group>>", d: "<<Sum>>");

        Generate(workbook);

        using var stream = new System.IO.MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;

        using var reloaded = new XLWorkbook(stream);
        var sheet = reloaded.Worksheet("Report");
        await Assert.That(sheet.Cell("A5").Value.GetText()).IsEqualTo("North Total");
        await Assert.That(sheet.Cell("D5").FormulaA1).IsEqualTo("SUBTOTAL(9,D3:D4)");
        await Assert.That(sheet.Row(3).OutlineLevel).IsEqualTo(1);
    }
}
