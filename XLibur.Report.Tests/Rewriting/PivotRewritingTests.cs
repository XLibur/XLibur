using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.Coordinates;

namespace XLibur.Report.Tests.Rewriting;

/// <summary>
/// A pivot table in a template has to reflect what the report generated, not what the template
/// held.
/// </summary>
public class PivotRewritingTests
{
    private static List<SaleItem> Items(int count) => Enumerable.Range(1, count)
        .Select(i => new SaleItem
        {
            Product = "Product " + i,
            Region = i % 2 == 0 ? "South" : "North",
            Quantity = i,
        })
        .ToList();

    /// <summary>
    /// A <c>Data</c> sheet whose heading is row 1 and whose bound range is rows 2 (data) and 3
    /// (options), plus a pivot table over the source <paramref name="source"/> chooses.
    /// </summary>
    private static (XLWorkbook Workbook, IXLPivotTable Pivot) Template(
        Func<IXLWorksheet, IXLRange> source,
        string pivotSheet = "Pivot")
    {
        var workbook = new XLWorkbook();
        var data = workbook.AddWorksheet("Data");

        data.Cell("A1").Value = "Region";
        data.Cell("B1").Value = "Quantity";
        data.Cell("A2").Value = "{{ item.Region }}";
        data.Cell("B2").Value = "{{ item.Quantity }}";

        workbook.DefinedNames.Add("Items", data.Range("A2:B3"));

        var target = pivotSheet == "Data" ? data : workbook.AddWorksheet(pivotSheet);
        var pivot = target.PivotTables.Add("pvt", target.Cell(pivotSheet == "Data" ? "D1" : "A1"), source(data));
        pivot.RowLabels.Add("Region");
        pivot.Values.Add("Quantity");

        return (workbook, pivot);
    }

    private static XLGenerateResult Generate(XLWorkbook workbook, int itemCount = 4)
    {
        using var template = new XLTemplate(workbook);
        template.AddVariable("Items", Items(itemCount));
        return template.Generate();
    }

    private static XLPivotCache Cache(IXLPivotTable pivot) => (XLPivotCache)pivot.PivotCache;

    private static string SourceArea(IXLPivotTable pivot)
    {
        var reference = (XLPivotSourceReference)Cache(pivot).Source;
        return reference.UsesName ? reference.Name : reference.Area.Value.Area.ToString();
    }

    /// <summary>
    /// The template's source covers the heading and the one repeated row; after generation it has to
    /// cover the heading and all four.
    /// </summary>
    [Test]
    public async Task AnAreaSourcedPivotIsRePointedAtTheGeneratedRows()
    {
        var (workbook, pivot) = Template(sheet => sheet.Range("A1:B2"));
        using var _ = workbook;

        Generate(workbook);

        await Assert.That(SourceArea(pivot)).IsEqualTo("A1:B5");
        await Assert.That(Cache(pivot).RecordCount).IsEqualTo(4);
    }

    [Test]
    public async Task ARePointedCacheHoldsTheGeneratedValues()
    {
        var (workbook, pivot) = Template(sheet => sheet.Range("A1:B2"));
        using var _ = workbook;

        Generate(workbook);

        await Assert.That(Cache(pivot).FieldNames).IsEquivalentTo(new[] { "Region", "Quantity" });
    }

    /// <summary>
    /// The spec's documented happy path. The name grows with the report on its own; what the
    /// rewriter supplies is the refresh — and losing exactly that is what upstream 0.2.12 regressed
    /// in its issue #399.
    /// </summary>
    [Test]
    public async Task ANameSourcedPivotIsRefreshedAndItsSourceLeftAlone()
    {
        var (workbook, pivot) = Template(sheet => sheet.Range("A1:B2"));
        using var _ = workbook;

        workbook.DefinedNames.Add("SourceData", workbook.Worksheet("Data").Range("A1:B3"));
        Cache(pivot).Source = new XLPivotSourceReference("SourceData");

        Generate(workbook);

        await Assert.That(SourceArea(pivot)).IsEqualTo("SourceData");
        await Assert.That(Cache(pivot).RecordCount).IsEqualTo(4);
    }

    [Test]
    public async Task ATableSourcedPivotIsRefreshedAndItsSourceLeftAlone()
    {
        var (workbook, pivot) = Template(sheet => sheet.Range("A1:B3").CreateTable("SourceTable"));
        using var _ = workbook;

        Generate(workbook);

        await Assert.That(SourceArea(pivot)).IsEqualTo("SourceTable");
        await Assert.That(Cache(pivot).RecordCount).IsEqualTo(4);
    }

    /// <summary>
    /// Excel is the authority on its own pivot layout, so it is asked to re-read the source when it
    /// opens the generated file.
    /// </summary>
    [Test]
    public async Task EveryTouchedCacheIsMarkedToRefreshOnOpen()
    {
        var (workbook, pivot) = Template(sheet => sheet.Range("A1:B2"));
        using var _ = workbook;

        Cache(pivot).RefreshDataOnOpen = false;

        Generate(workbook);

        await Assert.That(pivot.PivotCache.RefreshDataOnOpen).IsTrue();
    }

    /// <summary>
    /// A pivot below a bound range ends up below the rows the report generated, and moves exactly
    /// <em>once</em> doing it.
    /// <para>
    /// The expectation is unchanged; what produces it moved. <c>XLPivotTable</c> reacts to the row
    /// inserts expansion performs as of spec 33, so <c>PivotRewriter</c>'s own mover was deleted —
    /// with both in place this landed on row 12 rather than 10, because each moved it by the same
    /// two rows. That is the regression this test caught, and it is the reason it asserts a
    /// position rather than merely that something moved.
    /// </para>
    /// </summary>
    [Test]
    public async Task APivotBelowTheRangeMovesBelowTheGeneratedRows()
    {
        var (workbook, pivot) = Template(sheet => sheet.Range("A1:B2"), pivotSheet: "Data");
        using var _ = workbook;

        // The pivot's target starts at D1 on the data sheet; move it below the bound range.
        pivot.TargetCell = workbook.Worksheet("Data").Cell("D8");

        Generate(workbook);

        // Three rows were inserted for the extra items and the empty options row was removed.
        await Assert.That(pivot.TargetCell.Address.RowNumber).IsEqualTo(10);
    }

    [Test]
    public async Task APivotAboveTheRangeStaysWhereItIs()
    {
        var (workbook, pivot) = Template(sheet => sheet.Range("A1:B2"), pivotSheet: "Data");
        using var _ = workbook;

        Generate(workbook);

        await Assert.That(pivot.TargetCell.Address.RowNumber).IsEqualTo(1);
    }

    /// <summary>A pivot on its own sheet is never in the way, so it does not move.</summary>
    [Test]
    public async Task APivotOnAnotherSheetIsNotMoved()
    {
        var (workbook, pivot) = Template(sheet => sheet.Range("A1:B2"));
        using var _ = workbook;

        Generate(workbook);

        await Assert.That(pivot.TargetCell.Address.RowNumber).IsEqualTo(1);
    }

    /// <summary>A cache reading a sheet the report did not touch is left entirely alone.</summary>
    [Test]
    public async Task ACacheOverUntouchedDataIsNotRefreshed()
    {
        var (workbook, pivot) = Template(sheet => sheet.Range("A1:B2"));
        using var _ = workbook;

        var other = workbook.AddWorksheet("Static");
        other.Cell("A1").Value = "Kind";
        other.Cell("B1").Value = "Count";
        other.Cell("A2").Value = "Fixed";
        other.Cell("B2").Value = 7;

        var staticPivot = workbook.Worksheet("Pivot").PivotTables
            .Add("static", workbook.Worksheet("Pivot").Cell("F1"), other.Range("A1:B2"));
        staticPivot.RowLabels.Add("Kind");
        staticPivot.Values.Add("Count");
        Cache(staticPivot).RefreshDataOnOpen = false;

        Generate(workbook);

        await Assert.That(staticPivot.PivotCache.RefreshDataOnOpen).IsFalse();
        await Assert.That(SourceArea(staticPivot)).IsEqualTo("A1:B2");
        await Assert.That(SourceArea(pivot)).IsEqualTo("A1:B5");
    }

    /// <summary>An empty collection shrinks the source rather than leaving it over dead rows.</summary>
    [Test]
    public async Task AnAreaSourcedPivotOverAnEmptyRangeShrinks()
    {
        var (workbook, pivot) = Template(sheet => sheet.Range("A1:B2"));
        using var _ = workbook;

        var result = Generate(workbook, itemCount: 0);

        await Assert.That(result.HasErrors).IsFalse();
        await Assert.That(SourceArea(pivot)).IsEqualTo("A1:B1");
    }

    /// <summary>
    /// A pivot built on a name the report removed cannot be refreshed. That is a template problem to
    /// report, not an exception to abort a finished report with.
    /// </summary>
    [Test]
    public async Task APivotWhoseSourceNameDisappearsIsReportedRatherThanThrown()
    {
        var (workbook, pivot) = Template(sheet => sheet.Range("A1:B2"));
        using var _ = workbook;

        // Bound to the range itself, so an empty collection deletes the name along with the rows.
        workbook.DefinedNames.Add("Vanishing", workbook.Worksheet("Data").Range("A2:B3"));
        Cache(pivot).Source = new XLPivotSourceReference("Vanishing");

        var result = Generate(workbook, itemCount: 0);

        await Assert.That(result.HasErrors).IsTrue();
        await Assert.That(result.ParsingErrors[0].Message).Contains("pivot table's source data");
    }

    /// <summary>
    /// ClosedXML.Report's #200 has produced pivot output Excel refuses to open since 2021, so a
    /// generated pivot is saved through the OpenXML validator rather than only round-tripped.
    /// </summary>
    [Test]
    public async Task AGeneratedPivotPassesTheOpenXmlValidator()
    {
        var (workbook, _) = Template(sheet => sheet.Range("A1:B2"));
        using var _unused = workbook;

        Generate(workbook);

        using var stream = new MemoryStream();
        await Assert.That(() => workbook.SaveAs(stream, validate: true)).ThrowsNothing();
    }

    [Test]
    public async Task ARePointedSourceSurvivesASaveAndReload()
    {
        using var templateFile = new MemoryStream();

        var (workbook, _) = Template(sheet => sheet.Range("A1:B2"));
        using (workbook)
        {
            workbook.SaveAs(templateFile);
        }

        templateFile.Position = 0;
        using var generated = new MemoryStream();

        using (var template = new XLTemplate(templateFile))
        {
            template.AddVariable("Items", Items(4));
            template.Generate();
            template.SaveAs(generated);
        }

        generated.Position = 0;
        using var reloaded = new XLWorkbook(generated);
        var pivot = reloaded.Worksheet("Pivot").PivotTables.Single();

        await Assert.That(SourceArea(pivot)).IsEqualTo("A1:B5");
        await Assert.That(pivot.PivotCache.RefreshDataOnOpen).IsTrue();
    }
}
