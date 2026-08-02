using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Report.Tests.Tags;

/// <summary>
/// <c>&lt;&lt;Pivot&gt;&gt;</c> lays a pivot table out over the rows the report generated, from
/// column tags saying what each column is for.
/// </summary>
public class PivotTagTests
{
    private static List<SaleItem> Items(int count = 4) => Enumerable.Range(1, count)
        .Select(i => new SaleItem
        {
            Product = "Product " + i,
            Region = i % 2 == 0 ? "South" : "North",
            Category = i <= 2 ? "Retail" : "Trade",
            Quantity = i,
        })
        .ToList();

    /// <summary>
    /// A <c>Data</c> sheet with headings in row 1, the repeated row in row 2 and the options row in
    /// row 3, plus an empty <c>Summary</c> sheet for the pivot to be built onto.
    /// </summary>
    private static XLWorkbook Template(string options = "<<Pivot dest=\"Summary!A1\">>", string a3 = "<<Row>>", string c3 = "<<Data>>")
    {
        var workbook = new XLWorkbook();
        var data = workbook.AddWorksheet("Data");

        data.Cell("A1").Value = "Region";
        data.Cell("B1").Value = "Category";
        data.Cell("C1").Value = "Quantity";
        data.Cell("A2").Value = "{{ item.Region }}";
        data.Cell("B2").Value = "{{ item.Category }}";
        data.Cell("C2").Value = "{{ item.Quantity }}";

        data.Cell("A3").Value = a3;
        data.Cell("B3").Value = options;
        data.Cell("C3").Value = c3;

        workbook.DefinedNames.Add("Items", data.Range("A2:C3"));
        workbook.AddWorksheet("Summary");

        return workbook;
    }

    private static XLGenerateResult Generate(XLWorkbook workbook, int itemCount = 4)
    {
        using var template = new XLTemplate(workbook);
        template.AddVariable("Items", Items(itemCount));
        return template.Generate();
    }

    private static IXLPivotTable Pivot(XLWorkbook workbook) =>
        workbook.Worksheet("Summary").PivotTables.Single();

    private static string SourceArea(IXLPivotTable pivot)
    {
        var reference = (XLPivotSourceReference)((XLPivotCache)pivot.PivotCache).Source;
        return reference.UsesName ? reference.Name : reference.Area.Value.Area.ToString();
    }

    [Test]
    public async Task APivotIsBuiltAtTheDestinationTheTagNames()
    {
        using var workbook = Template();

        var result = Generate(workbook);

        await Assert.That(result.HasErrors).IsFalse();
        await Assert.That(Pivot(workbook).TargetCell.Address.ToString()).IsEqualTo("A1");
    }

    /// <summary>
    /// The source is the heading row plus every generated row — the point of the tag being that the
    /// template does not know how many there will be.
    /// </summary>
    [Test]
    public async Task TheSourceCoversTheHeadingRowAndEveryGeneratedRow()
    {
        using var workbook = Template();

        Generate(workbook);

        await Assert.That(SourceArea(Pivot(workbook))).IsEqualTo("A1:C5");
    }

    [Test]
    public async Task TheCacheHoldsOneRecordPerItem()
    {
        using var workbook = Template();

        Generate(workbook);

        await Assert.That(((XLPivotCache)Pivot(workbook).PivotCache).RecordCount).IsEqualTo(4);
    }

    /// <summary>The field tags name their fields from the heading above the column they sit in.</summary>
    [Test]
    public async Task AColumnTaggedRowBecomesARowLabel()
    {
        using var workbook = Template();

        Generate(workbook);

        await Assert.That(Pivot(workbook).RowLabels.Select(f => f.SourceName)).IsEquivalentTo(new[] { "Region" });
    }

    [Test]
    public async Task AColumnTaggedDataBecomesAValueField()
    {
        using var workbook = Template();

        Generate(workbook);

        await Assert.That(Pivot(workbook).Values.Single().SourceName).IsEqualTo("Quantity");
    }

    [Test]
    public async Task AColumnTaggedColumnBecomesAColumnLabel()
    {
        using var workbook = Template();
        workbook.Worksheet("Data").Cell("B3").Value = "<<Column>><<Pivot dest=\"Summary!A1\">>";

        Generate(workbook);

        await Assert.That(Pivot(workbook).ColumnLabels.Select(f => f.SourceName)).Contains("Category");
    }

    /// <summary><c>&lt;&lt;Col&gt;&gt;</c> is the same tag, spelled the way a template author might.</summary>
    [Test]
    public async Task ColIsAnAliasForColumn()
    {
        using var workbook = Template();
        workbook.Worksheet("Data").Cell("B3").Value = "<<Col>><<Pivot dest=\"Summary!A1\">>";

        Generate(workbook);

        await Assert.That(Pivot(workbook).ColumnLabels.Select(f => f.SourceName)).Contains("Category");
    }

    [Test]
    public async Task AColumnTaggedPageBecomesAReportFilter()
    {
        using var workbook = Template();
        workbook.Worksheet("Data").Cell("B3").Value = "<<Page>><<Pivot dest=\"Summary!A1\">>";

        Generate(workbook);

        await Assert.That(Pivot(workbook).ReportFilters.Select(f => f.SourceName)).IsEquivalentTo(new[] { "Category" });
    }

    [Test]
    public async Task AValueFieldSumsByDefault()
    {
        using var workbook = Template();

        Generate(workbook);

        await Assert.That(Pivot(workbook).Values.Single().SummaryFormula).IsEqualTo(XLPivotSummary.Sum);
    }

    [Test]
    [Arguments("<<Data avg>>", XLPivotSummary.Average)]
    [Arguments("<<Data func=avg>>", XLPivotSummary.Average)]
    [Arguments("<<Data average>>", XLPivotSummary.Average)]
    [Arguments("<<Data max>>", XLPivotSummary.Maximum)]
    [Arguments("<<Data min>>", XLPivotSummary.Minimum)]
    [Arguments("<<Data count>>", XLPivotSummary.Count)]
    [Arguments("<<Data product>>", XLPivotSummary.Product)]
    [Arguments("<<Data stddev>>", XLPivotSummary.StandardDeviation)]
    [Arguments("<<Data varp>>", XLPivotSummary.PopulationVariance)]
    public async Task AValueFieldSummarisesTheWayTheTagSays(string tag, XLPivotSummary expected)
    {
        using var workbook = Template(c3: tag);

        Generate(workbook);

        await Assert.That(Pivot(workbook).Values.Single().SummaryFormula).IsEqualTo(expected);
    }

    [Test]
    public async Task AValueFieldCanBeGivenItsOwnTitle()
    {
        using var workbook = Template(c3: "<<Data title=\"Units sold\">>");

        Generate(workbook);

        await Assert.That(Pivot(workbook).Values.Single().CustomName).IsEqualTo("Units sold");
    }

    [Test]
    public async Task FieldsAreAddedInTheOrderTheirColumnsRun()
    {
        using var workbook = Template();
        workbook.Worksheet("Data").Cell("B3").Value = "<<Row>><<Pivot dest=\"Summary!A1\">>";

        Generate(workbook);

        await Assert.That(Pivot(workbook).RowLabels.Select(f => f.SourceName).ToList())
            .IsEquivalentTo(new[] { "Region", "Category" });
    }

    /// <summary>
    /// Excel is the authority on its own pivot layout, so a generated pivot asks it to re-read the
    /// source on open just as a template's own pivot is made to.
    /// </summary>
    [Test]
    public async Task AGeneratedPivotIsMarkedToRefreshOnOpen()
    {
        using var workbook = Template();

        Generate(workbook);

        await Assert.That(Pivot(workbook).PivotCache.RefreshDataOnOpen).IsTrue();
    }

    /// <summary>
    /// The rewriter re-points what the template drew. A pivot built during generation is already
    /// reading the generated rows, and replaying the expansion over it would move it off them.
    /// </summary>
    [Test]
    public async Task AGeneratedPivotIsNotRePointedASecondTime()
    {
        using var workbook = Template();

        Generate(workbook);

        // A1:C5, not the A1:C8 that re-applying the four-row expansion would give.
        await Assert.That(SourceArea(Pivot(workbook))).IsEqualTo("A1:C5");
    }

    /// <summary>
    /// Built after the rewriter has re-pointed what the template drew, so a pivot the template already
    /// had over the same rows shares its cache rather than the workbook carrying two identical ones.
    /// </summary>
    [Test]
    public async Task AGeneratedPivotSharesACacheWithATemplatePivotOverTheSameRows()
    {
        using var workbook = Template();
        var data = workbook.Worksheet("Data");

        // The template's own pivot, over the template's geometry — which the rewriter grows to A1:C5,
        // exactly what the generated one is built over.
        var drawn = workbook.Worksheet("Summary").PivotTables.Add("Drawn", workbook.Worksheet("Summary").Cell("F1"), data.Range("A1:C3"));
        drawn.RowLabels.Add("Region");
        drawn.Values.Add("Quantity");

        Generate(workbook);

        await Assert.That(workbook.PivotCaches.Count()).IsEqualTo(1);
    }

    /// <summary>
    /// A destination named as a cell on the range's own sheet, and one named by a defined name, both
    /// work — a template author picks whichever survives editing.
    /// </summary>
    [Test]
    public async Task ADestinationMayBeUnqualifiedMeaningTheRangesOwnSheet()
    {
        using var workbook = Template(options: "<<Pivot dest=\"F1\">>");

        var result = Generate(workbook);

        await Assert.That(result.HasErrors).IsFalse();
        await Assert.That(workbook.Worksheet("Data").PivotTables.Single().TargetCell.Address.ToString())
            .IsEqualTo("F1");
    }

    [Test]
    public async Task ADestinationMayBeADefinedNameCoveringOneCell()
    {
        using var workbook = Template(options: "<<Pivot dest=\"PivotHere\">>");
        workbook.DefinedNames.Add("PivotHere", workbook.Worksheet("Summary").Range("B4:B4"));

        var result = Generate(workbook);

        await Assert.That(result.HasErrors).IsFalse();
        await Assert.That(Pivot(workbook).TargetCell.Address.ToString()).IsEqualTo("B4");
    }

    /// <summary>
    /// The options row is dropped once the tags have run and nothing was written into it, taking a row
    /// out from under anything below. A pivot's position would not follow, so the tag applies the
    /// shift itself.
    /// </summary>
    [Test]
    public async Task ADestinationBelowTheRangeAllowsForTheOptionsRowBeingDropped()
    {
        using var workbook = Template(options: "<<Pivot dest=\"A10\">>");

        Generate(workbook);

        // Three rows were inserted for the extra items and the empty options row was removed, so the
        // template's row 10 is the report's row 12.
        await Assert.That(workbook.Worksheet("Data").PivotTables.Single().TargetCell.Address.ToString())
            .IsEqualTo("A12");
    }

    /// <summary>The same destination, but with a total keeping the options row alive.</summary>
    [Test]
    public async Task ADestinationBelowARangeWhoseOptionsRowSurvivesMovesOneFurther()
    {
        using var workbook = Template(options: "<<Pivot dest=\"A10\">><<Sum>>");

        Generate(workbook);

        await Assert.That(workbook.Worksheet("Data").PivotTables.Single().TargetCell.Address.ToString())
            .IsEqualTo("A13");
    }

    [Test]
    public async Task APivotCanBeGivenAName()
    {
        using var workbook = Template(options: "<<Pivot dest=\"Summary!A1\" name=\"ByRegion\">>");

        Generate(workbook);

        await Assert.That(Pivot(workbook).Name).IsEqualTo("ByRegion");
    }

    /// <summary>
    /// Two pivots in one workbook need two names, and a template generating them cannot be made to
    /// supply both.
    /// </summary>
    [Test]
    public async Task GeneratedPivotsAreNamedApartFromEachOther()
    {
        using var workbook = Template();
        var data = workbook.Worksheet("Data");

        data.Cell("A6").Value = "Region";
        data.Cell("B6").Value = "Quantity";
        data.Cell("A7").Value = "{{ item.Region }}";
        data.Cell("B7").Value = "{{ item.Quantity }}";
        data.Cell("A8").Value = "<<Row>>";
        data.Cell("B8").Value = "<<Data>><<Pivot dest=\"Summary!F1\">>";
        workbook.DefinedNames.Add("Second", data.Range("A7:B8"));

        using var template = new XLTemplate(workbook);
        template.AddVariable("Items", Items());
        template.AddVariable("Second", Items());
        var result = template.Generate();

        await Assert.That(result.HasErrors).IsFalse();
        await Assert.That(workbook.Worksheet("Summary").PivotTables.Select(p => p.Name).Distinct().Count())
            .IsEqualTo(2);
    }

    [Test]
    public async Task ADestMissingIsReported()
    {
        using var workbook = Template(options: "<<Pivot>>");

        var result = Generate(workbook);

        await Assert.That(result.HasErrors).IsTrue();
        await Assert.That(result.ParsingErrors[0].Message).Contains("needs a dest");
    }

    [Test]
    public async Task ADestOnASheetThatIsNotThereIsReported()
    {
        using var workbook = Template(options: "<<Pivot dest=\"Nowhere!A1\">>");

        var result = Generate(workbook);

        await Assert.That(result.HasErrors).IsTrue();
        await Assert.That(result.ParsingErrors[0].Message).Contains("not a cell this library can find");
    }

    [Test]
    public async Task APivotWithNoFieldsIsReported()
    {
        using var workbook = Template(a3: string.Empty, c3: string.Empty);

        var result = Generate(workbook);

        await Assert.That(result.HasErrors).IsTrue();
        await Assert.That(result.ParsingErrors[0].Message).Contains("no fields to lay out");
    }

    /// <summary>
    /// A pivot names its fields from the heading row above the range, so a range starting in row 1 has
    /// nothing to name them with.
    /// </summary>
    [Test]
    public async Task ARangeWithNoHeadingRowIsReported()
    {
        using var workbook = new XLWorkbook();
        var data = workbook.AddWorksheet("Data");
        data.Cell("A1").Value = "{{ item.Region }}";
        data.Cell("A2").Value = "<<Row>><<Pivot dest=\"Summary!A1\">>";
        workbook.DefinedNames.Add("Items", data.Range("A1:A2"));
        workbook.AddWorksheet("Summary");

        var result = Generate(workbook);

        await Assert.That(result.HasErrors).IsTrue();
        await Assert.That(result.ParsingErrors[0].Message).Contains("heading row");
    }

    [Test]
    public async Task AFieldOverAColumnWithNoHeadingIsReported()
    {
        using var workbook = Template();
        workbook.Worksheet("Data").Cell("C1").Clear(XLClearOptions.Contents);

        var result = Generate(workbook);

        await Assert.That(result.HasErrors).IsTrue();
        await Assert.That(result.ParsingErrors[0].Message).Contains("no heading");
    }

    /// <summary>
    /// Grouping writes subtotal rows into the generated block, and a pivot over them would count each
    /// group twice. Refused rather than half-done.
    /// </summary>
    [Test]
    public async Task APivotInAGroupedRangeIsRefused()
    {
        using var workbook = Template(a3: "<<Row>><<Group>>");

        var result = Generate(workbook);

        await Assert.That(result.HasErrors).IsTrue();
        await Assert.That(result.ParsingErrors[0].Message).Contains("cannot be used in a grouped range");
        await Assert.That(workbook.Worksheet("Summary").PivotTables.Count()).IsEqualTo(0);
    }

    [Test]
    public async Task APivotInAHorizontalRangeIsRefused()
    {
        using var workbook = new XLWorkbook();
        var data = workbook.AddWorksheet("Data");
        data.Cell("A2").Value = "Region";
        data.Cell("B2").Value = "{{ item.Region }}";
        data.Cell("C2").Value = "<<Horizontal>><<Pivot dest=\"Summary!A1\">>";
        workbook.DefinedNames.Add("Items", data.Range("B2:C2"));
        workbook.AddWorksheet("Summary");

        var result = Generate(workbook);

        await Assert.That(result.HasErrors).IsTrue();
        await Assert.That(result.ParsingErrors[0].Message).Contains("repeats across");
    }

    /// <summary>
    /// A field tag with no pivot to read it is a template mistake worth naming: silently doing nothing
    /// looks like the pivot failing for some other reason.
    /// </summary>
    [Test]
    public async Task AFieldTagWithNoPivotIsReported()
    {
        using var workbook = Template(options: string.Empty);

        var result = Generate(workbook);

        await Assert.That(result.HasErrors).IsTrue();
        await Assert.That(result.ParsingErrors[0].Message).Contains("only means something to a <<Pivot>>");
    }

    /// <summary>
    /// <c>&lt;&lt;Pivot&gt;&gt;</c> runs after <c>&lt;&lt;Delete&gt;&gt;</c>, so a column removed for
    /// the reader's benefit is one the pivot never sees — its cache source is a plain rectangle that a
    /// later deletion would not correct.
    /// </summary>
    [Test]
    public async Task APivotIsBuiltOverTheColumnsThatSurviveDelete()
    {
        using var workbook = Template();
        workbook.Worksheet("Data").Cell("B3").Value = "<<Delete>><<Pivot dest=\"Summary!A1\">>";

        var result = Generate(workbook);

        await Assert.That(result.HasErrors).IsFalse();
        await Assert.That(SourceArea(Pivot(workbook))).IsEqualTo("A1:B5");
    }

    /// <summary>
    /// ClosedXML.Report's #200 has produced pivot output Excel refuses to open since 2021, so a
    /// generated pivot goes through the OpenXML validator rather than only a round-trip.
    /// </summary>
    [Test]
    public async Task AGeneratedPivotPassesTheOpenXmlValidator()
    {
        using var workbook = Template();
        workbook.Worksheet("Data").Cell("B3").Value = "<<Column>><<Pivot dest=\"Summary!A1\">>";

        Generate(workbook);

        using var stream = new MemoryStream();
        await Assert.That(() => workbook.SaveAs(stream, validate: true)).ThrowsNothing();
    }

    [Test]
    public async Task AGeneratedPivotSurvivesASaveAndReload()
    {
        using var templateFile = new MemoryStream();

        using (var workbook = Template())
        {
            workbook.SaveAs(templateFile);
        }

        templateFile.Position = 0;
        using var generated = new MemoryStream();

        using (var template = new XLTemplate(templateFile))
        {
            template.AddVariable("Items", Items());
            template.Generate();
            template.SaveAs(generated);
        }

        generated.Position = 0;
        using var reloaded = new XLWorkbook(generated);
        var pivot = reloaded.Worksheet("Summary").PivotTables.Single();

        await Assert.That(SourceArea(pivot)).IsEqualTo("A1:C5");
        await Assert.That(pivot.RowLabels.Select(f => f.SourceName)).Contains("Region");
        await Assert.That(pivot.Values.Single().SourceName).IsEqualTo("Quantity");
    }
}
