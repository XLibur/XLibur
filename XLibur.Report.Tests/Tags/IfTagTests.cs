using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Report.Tests.Tags;

public class IfTagTests
{
    private static List<SaleItem> Items() => new()
    {
        new() { Region = "North", Product = "Widget", Quantity = 2 },
        new() { Region = "South", Product = "Gadget", Quantity = 0 },
        new() { Region = "North", Product = "Doohickey", Quantity = 5 },
        new() { Region = "South", Product = "Sprocket", Quantity = 0 },
    };

    /// <summary>
    /// A three-column range over A3:C4: row 3 repeats, row 4 is the options row. Anything given for
    /// <paramref name="c3"/> lands in a repeated row, anything for the row-4 arguments in the options
    /// row.
    /// </summary>
    private static IXLWorkbook Template(string c3 = "", string a4 = "", string b4 = "", string c4 = "")
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");

        sheet.Cell("A2").Value = "Region";
        sheet.Cell("B2").Value = "Product";
        sheet.Cell("A3").Value = "{{ item.Region }}";
        sheet.Cell("B3").Value = "{{ item.Product }}";

        foreach (var (address, text) in new[] { ("C3", c3), ("A4", a4), ("B4", b4), ("C4", c4) })
        {
            if (text.Length > 0)
            {
                sheet.Cell(address).Value = text;
            }
        }

        workbook.DefinedNames.Add("Items", sheet.Range("A3:C4"));
        return workbook;
    }

    private static XLGenerateResult Generate(IXLWorkbook workbook, List<SaleItem>? items = null)
    {
        using var template = new XLTemplate(workbook);
        template.AddVariable("Items", items ?? Items());
        template.AddVariable("ShowAll", false);
        return template.Generate();
    }

    private static List<string> Products(IXLWorksheet sheet, int firstRow, int count) =>
        Enumerable.Range(firstRow, count)
            .Select(row => sheet.Cell("B" + row).Value.ToString() ?? string.Empty)
            .ToList();

    private static int LastUsedRow(IXLWorksheet sheet) =>
        sheet.RangeUsed()?.RangeAddress.LastAddress.RowNumber ?? 0;

    // ── Row level ───────────────────────────────────────────────────────

    [Test]
    public async Task ATestInARepeatedRowKeepsOnlyTheRowsThatAnswerYes()
    {
        using var workbook = Template(c3: "<<If test=\"item.Quantity > 0\">>");

        Generate(workbook);

        await Assert.That(Products(workbook.Worksheet("Report"), 3, 2))
            .IsEquivalentTo(new[] { "Widget", "Doohickey" });
        await Assert.That(LastUsedRow(workbook.Worksheet("Report"))).IsEqualTo(4);
    }

    /// <summary>What survives the test is what everything else sees.</summary>
    [Test]
    public async Task SurvivingRowsAreTheOnesSorted()
    {
        using var workbook = Template(c3: "<<If test=\"item.Quantity > 0\">>", b4: "<<Sort desc>>");

        Generate(workbook);

        await Assert.That(Products(workbook.Worksheet("Report"), 3, 2))
            .IsEquivalentTo(new[] { "Widget", "Doohickey" });
    }

    [Test]
    public async Task SurvivingRowsAreTheOnesGroupedAndTotalled()
    {
        using var workbook = Template(
            c3: "<<If test=\"item.Quantity > 0\">>",
            a4: "<<Group>>",
            c4: "<<Sum over=C>>");

        Generate(workbook);

        // Both survivors are in North, so there is one group of two and one grand total.
        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("A5").Value.GetText()).IsEqualTo("North Total");
        await Assert.That(sheet.Cell("C5").FormulaA1).IsEqualTo("SUBTOTAL(9,C3:C4)");
        await Assert.That(sheet.Cell("C6").FormulaA1).IsEqualTo("SUBTOTAL(9,C3:C5)");
    }

    /// <summary>
    /// The documented surprise: only <c>null</c> and <c>false</c> are false, so a bare quantity keeps
    /// the zeroes and the comparison has to be written out.
    /// </summary>
    [Test]
    public async Task ZeroReadsAsTrue()
    {
        using var workbook = Template(c3: "<<If test=\"item.Quantity\">>");

        Generate(workbook);

        await Assert.That(Products(workbook.Worksheet("Report"), 3, 4))
            .IsEquivalentTo(new[] { "Widget", "Gadget", "Doohickey", "Sprocket" });
    }

    [Test]
    public async Task NullReadsAsFalse()
    {
        using var workbook = Template(c3: "<<If test=\"item.Notes\">>");

        Generate(workbook);

        // No item sets Notes, so every row is dropped and the range goes with them.
        await Assert.That(LastUsedRow(workbook.Worksheet("Report"))).IsEqualTo(2);
    }

    /// <summary>A test may share its cell with the expression the column displays.</summary>
    [Test]
    public async Task ATestCanShareACellWithAnExpression()
    {
        using var workbook = Template(c3: "{{ item.Quantity }}<<If test=\"item.Quantity > 0\">>");

        Generate(workbook);

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("C3").Value.GetNumber()).IsEqualTo(2d);
        await Assert.That(sheet.Cell("C4").Value.GetNumber()).IsEqualTo(5d);
        await Assert.That(LastUsedRow(sheet)).IsEqualTo(4);
    }

    [Test]
    public async Task TagTextNeverReachesTheReport()
    {
        using var workbook = Template(c3: "<<If test=\"item.Quantity > 0\">>");

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").Cell("C3").Value.IsBlank).IsTrue();
    }

    // ── Range level ─────────────────────────────────────────────────────

    [Test]
    public async Task ATestInTheOptionsRowThatAnswersYesKeepsEverything()
    {
        using var workbook = Template(c4: "<<If test=\"true\">>");

        Generate(workbook);

        await Assert.That(Products(workbook.Worksheet("Report"), 3, 4))
            .IsEquivalentTo(new[] { "Widget", "Gadget", "Doohickey", "Sprocket" });
    }

    /// <summary>A no renders the range exactly as an empty collection does.</summary>
    [Test]
    public async Task ATestInTheOptionsRowThatAnswersNoRemovesTheRange()
    {
        using var workbook = Template(c4: "<<If test=\"ShowAll\">>");

        Generate(workbook);

        await Assert.That(LastUsedRow(workbook.Worksheet("Report"))).IsEqualTo(2);
    }

    /// <summary>The range-level question can be about the collection itself.</summary>
    [Test]
    public async Task ARangeLevelTestCanAskAboutTheCollection()
    {
        using var workbook = Template(c4: "<<If test=\"items.size > 10\">>");

        Generate(workbook);

        await Assert.That(LastUsedRow(workbook.Worksheet("Report"))).IsEqualTo(2);
    }

    [Test]
    public async Task ARangeLevelNoLeavesAnOptionsRowTotalBehavingAsOverNoData()
    {
        using var workbook = Template(a4: "Total <<If test=\"ShowAll\">>", c4: "<<Sum>>");

        Generate(workbook);

        // The options row still holds content, so it survives with the total over nothing.
        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("A3").Value.GetText()).IsEqualTo("Total");
        await Assert.That(LastUsedRow(sheet)).IsEqualTo(3);
    }

    // ── Errors ──────────────────────────────────────────────────────────

    [Test]
    public async Task AnIfWithNothingToTestIsReported()
    {
        using var workbook = Template(c3: "<<If>>");

        var result = Generate(workbook);

        await Assert.That(result.HasErrors).IsTrue();
        await Assert.That(result.ParsingErrors[0].Message).Contains("needs something to test");
    }

    /// <summary>
    /// A test that will not evaluate answers no, and says why, rather than aborting a report that is
    /// otherwise fine.
    /// </summary>
    [Test]
    public async Task ATestThatWillNotEvaluateIsReportedAndAnswersNo()
    {
        using var workbook = Template(c3: "<<If test=\"item.Quantity >\">>");

        var result = Generate(workbook);

        await Assert.That(result.HasErrors).IsTrue();
        await Assert.That(LastUsedRow(workbook.Worksheet("Report"))).IsEqualTo(2);
    }

    // ── The same machinery under Delete ─────────────────────────────────

    /// <summary>
    /// <c>&lt;&lt;Delete keep&gt;&gt;</c> documented an interpolated value long before one worked;
    /// the tag now evaluates its parameter the way <c>&lt;&lt;If&gt;&gt;</c> evaluates its test.
    /// </summary>
    [Test]
    public async Task DeleteKeepsTheColumnWhenAnInterpolatedValueSaysSo()
    {
        using var workbook = Template(c3: "{{ item.Quantity }}", c4: "<<Delete keep=\"{{ !ShowAll }}\">>");

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").Cell("C3").Value.GetNumber()).IsEqualTo(2d);
    }

    [Test]
    public async Task DeleteRemovesTheColumnWhenAnInterpolatedValueSaysNo()
    {
        using var workbook = Template(c3: "{{ item.Quantity }}", c4: "<<Delete keep=\"{{ ShowAll }}\">>");

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").Cell("C3").Value.IsBlank).IsTrue();
    }

    [Test]
    public async Task DeleteKeepsTheColumnOnTheBareFlag()
    {
        using var workbook = Template(c3: "{{ item.Quantity }}", c4: "<<Delete keep>>");

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Report").Cell("C3").Value.GetNumber()).IsEqualTo(2d);
    }
}
