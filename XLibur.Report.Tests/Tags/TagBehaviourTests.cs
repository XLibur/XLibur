using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Report.Tags;

namespace XLibur.Report.Tests.Tags;

public class TagBehaviourTests
{
    private static List<SaleItem> Items() => new()
    {
        new() { Product = "Widget", Quantity = 2, UnitPrice = 5m, SoldOn = new DateTime(2026, 3, 1) },
        new() { Product = "Gadget", Quantity = 5, UnitPrice = 10m, SoldOn = new DateTime(2026, 1, 1) },
        new() { Product = "Doohickey", Quantity = 1, UnitPrice = 7.5m, SoldOn = new DateTime(2026, 2, 1) },
    };

    /// <summary>
    /// Builds a two-column range over A3:B4, where row 3 is the data row and row 4 the options row
    /// holding <paramref name="optionsA"/> and <paramref name="optionsB"/>.
    /// </summary>
    private static XLWorkbook Template(string optionsA = "", string optionsB = "")
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");

        sheet.Cell("A2").Value = "Product";
        sheet.Cell("B2").Value = "Quantity";
        sheet.Cell("A3").Value = "{{ item.Product }}";
        sheet.Cell("B3").Value = "{{ item.Quantity }}";

        if (optionsA.Length > 0)
        {
            sheet.Cell("A4").Value = optionsA;
        }

        if (optionsB.Length > 0)
        {
            sheet.Cell("B4").Value = optionsB;
        }

        workbook.DefinedNames.Add("Items", sheet.Range("A3:B4"));
        return workbook;
    }

    private static XLGenerateResult Generate(IXLWorkbook workbook)
    {
        using var template = new XLTemplate(workbook);
        template.AddVariable("Items", Items());
        return template.Generate();
    }

    private static List<string> ColumnText(IXLWorksheet sheet, string column, int firstRow, int count) =>
        Enumerable.Range(firstRow, count).Select(r => sheet.Cell(column + r).Value.ToString() ?? string.Empty).ToList();

    [Test]
    public async Task SortOrdersByTheColumnItSitsUnder()
    {
        using var workbook = Template(optionsA: "<<Sort>>");

        Generate(workbook);

        await Assert.That(ColumnText(workbook.Worksheet("Report"), "A", 3, 3))
            .IsEquivalentTo(new[] { "Doohickey", "Gadget", "Widget" });
    }

    [Test]
    public async Task SortDescendingReversesTheOrder()
    {
        using var workbook = Template(optionsA: "<<Sort desc>>");

        Generate(workbook);

        await Assert.That(ColumnText(workbook.Worksheet("Report"), "A", 3, 3))
            .IsEquivalentTo(new[] { "Widget", "Gadget", "Doohickey" });
    }

    [Test]
    public async Task DescTagSortsDescending()
    {
        using var workbook = Template(optionsB: "<<Desc>>");

        Generate(workbook);

        await Assert.That(ColumnText(workbook.Worksheet("Report"), "A", 3, 3))
            .IsEquivalentTo(new[] { "Gadget", "Widget", "Doohickey" });
    }

    [Test]
    public async Task SortUsesANumericColumnNumerically()
    {
        using var workbook = Template(optionsB: "<<Sort>>");

        Generate(workbook);

        await Assert.That(ColumnText(workbook.Worksheet("Report"), "A", 3, 3))
            .IsEquivalentTo(new[] { "Doohickey", "Widget", "Gadget" });
    }

    /// <summary>Sorting by something the report does not display needs an explicit key.</summary>
    [Test]
    public async Task SortByAnExplicitExpression()
    {
        using var workbook = Template(optionsA: "<<Sort by=item.SoldOn>>");

        Generate(workbook);

        await Assert.That(ColumnText(workbook.Worksheet("Report"), "A", 3, 3))
            .IsEquivalentTo(new[] { "Gadget", "Doohickey", "Widget" });
    }

    [Test]
    public async Task SumWritesASubtotalOverTheGeneratedRows()
    {
        using var workbook = Template(optionsB: "<<Sum>>");

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").Cell("B6").FormulaA1).IsEqualTo("SUBTOTAL(9,B3:B5)");
    }

    [Test]
    [Arguments("<<Avg>>", "SUBTOTAL(1,B3:B5)")]
    [Arguments("<<Count>>", "SUBTOTAL(2,B3:B5)")]
    [Arguments("<<Max>>", "SUBTOTAL(4,B3:B5)")]
    [Arguments("<<Min>>", "SUBTOTAL(5,B3:B5)")]
    public async Task SummaryTagsWriteTheirSubtotalNumber(string tag, string expected)
    {
        using var workbook = Template(optionsB: tag);

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").Cell("B6").FormulaA1).IsEqualTo(expected);
    }

    [Test]
    public async Task SumCanTotalAnotherColumn()
    {
        using var workbook = Template(optionsA: "<<Sum over=B>>");

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").Cell("A6").FormulaA1).IsEqualTo("SUBTOTAL(9,B3:B5)");
    }

    /// <summary>An options row holding a total is not an empty row, so it survives.</summary>
    [Test]
    public async Task OptionsRowSurvivesWhenATagWroteIntoIt()
    {
        using var workbook = Template(optionsB: "<<Sum>>");

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").RangeUsed()!.RangeAddress.LastAddress.RowNumber).IsEqualTo(6);
    }

    /// <summary>An options row that held nothing but tags is removed once they are read.</summary>
    [Test]
    public async Task OptionsRowIsRemovedWhenItsTagsLeaveNothingBehind()
    {
        using var workbook = Template(optionsA: "<<Sort>>");

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").RangeUsed()!.RangeAddress.LastAddress.RowNumber).IsEqualTo(5);
    }

    [Test]
    public async Task TagTextNeverReachesTheReport()
    {
        using var workbook = Template(optionsA: "Total <<Sort>>");

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").Cell("A6").Value.GetText()).IsEqualTo("Total");
    }

    /// <summary>
    /// A summary tag writes into its own cell, so a label belongs in a neighbouring column rather
    /// than alongside the tag.
    /// </summary>
    [Test]
    public async Task ALabelInOneColumnCanSitBesideATotalInAnother()
    {
        using var workbook = Template(optionsA: "Total", optionsB: "<<Sum>>");

        Generate(workbook);

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("A6").Value.GetText()).IsEqualTo("Total");
        await Assert.That(sheet.Cell("B6").FormulaA1).IsEqualTo("SUBTOTAL(9,B3:B5)");
    }

    [Test]
    public async Task AutoFilterCoversTheHeaderAndTheGeneratedRows()
    {
        using var workbook = Template(optionsA: "<<AutoFilter>>");

        Generate(workbook);

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.AutoFilter.IsEnabled).IsTrue();
        await Assert.That(sheet.AutoFilter.Range.RangeAddress.ToStringRelative()).IsEqualTo("A2:B5");
    }

    [Test]
    public async Task HiddenHidesTheColumnItSitsIn()
    {
        using var workbook = Template(optionsB: "<<Hidden>>");

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").Column(2).IsHidden).IsTrue();
    }

    [Test]
    public async Task DeleteRemovesTheColumnItSitsIn()
    {
        using var workbook = Template(optionsB: "<<Delete>>");

        Generate(workbook);

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("A3").Value.GetText()).IsEqualTo("Widget");
        await Assert.That(sheet.Cell("B3").Value.IsBlank).IsTrue();
    }

    /// <summary>
    /// Sorting by a column and then removing it is the reason Delete runs last.
    /// </summary>
    [Test]
    public async Task AColumnCanBeSortedByAndThenRemoved()
    {
        using var workbook = Template(optionsB: "<<Sort desc>><<Delete>>");

        Generate(workbook);

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("A3").Value.GetText()).IsEqualTo("Gadget");
        await Assert.That(sheet.Cell("B3").Value.IsBlank).IsTrue();
    }

    [Test]
    public async Task DeleteCanBeKept()
    {
        using var workbook = Template(optionsB: "<<Delete keep=true>>");

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").Cell("B3").Value.GetNumber()).IsEqualTo(2d);
    }

    [Test]
    public async Task UnknownTagIsReportedWithoutStoppingGeneration()
    {
        using var workbook = Template(optionsA: "<<NoSuchTag>>");

        var result = Generate(workbook);

        await Assert.That(result.HasErrors).IsTrue();
        await Assert.That(result.ParsingErrors[0].Message).Contains("NoSuchTag");
        await Assert.That(workbook.Worksheet("Report").Cell("A3").Value.GetText()).IsEqualTo("Widget");
    }

    [Test]
    public async Task SortWithNothingToSortByIsReported()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        sheet.Cell("A1").Value = "fixed text";
        sheet.Cell("A2").Value = "<<Sort>>";
        workbook.DefinedNames.Add("Items", sheet.Range("A1:A2"));

        using var template = new XLTemplate(workbook);
        template.AddVariable("Items", Items());
        var result = template.Generate();

        await Assert.That(result.HasErrors).IsTrue();
        await Assert.That(result.ParsingErrors[0].Message).Contains("nothing to sort by");
    }

    [Test]
    public async Task ACustomTagCanBeRegistered()
    {
        TagsRegister.Add<StampTag>("Stamp");
        using var workbook = Template(optionsA: "<<Stamp>>");

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").Cell("A6").Value.GetText()).IsEqualTo("stamped 3");
    }

    [Test]
    public async Task BuiltInTagsAreRegistered()
    {
        await Assert.That(TagsRegister.Contains("Sum")).IsTrue();
        await Assert.That(TagsRegister.Contains("sort")).IsTrue();
        await Assert.That(TagsRegister.Contains("Nonsense")).IsFalse();
    }

    private sealed class StampTag : OptionTag
    {
        public override void Execute(ProcessingContext context)
        {
            if (context.OptionsRow is null)
            {
                return;
            }

            context.Worksheet
                .Cell(context.OptionsRow.RangeAddress.FirstAddress.RowNumber, Column)
                .Value = $"stamped {context.Items.Count}";
        }
    }
}
