using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Report.Tests.Rewriting;

/// <summary>
/// A chart whose series plot a bound range have to plot every row the report generated, not the one
/// row the template repeated.
/// </summary>
public class ChartRewritingTests
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
    /// A sheet whose rows 3 (data) and 4 (options) are named <c>Items</c>, with a column chart
    /// plotting the quantity column against the product column.
    /// </summary>
    private static XLWorkbook Template(string valueReferences = "Report!$B$3:$B$3", string? categoryReferences = "Report!$A$3:$A$3")
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");

        sheet.Cell("A2").Value = "Product";
        sheet.Cell("B2").Value = "Quantity";
        sheet.Cell("A3").Value = "{{ item.Product }}";
        sheet.Cell("B3").Value = "{{ item.Quantity }}";

        var chart = sheet.Charts.Add(XLChartType.ColumnClustered);
        chart.Series.Add("Quantity", valueReferences, categoryReferences);

        workbook.DefinedNames.Add("Items", sheet.Range("A3:B4"));
        return workbook;
    }

    private static XLGenerateResult Generate(IXLWorkbook workbook, int itemCount = 4)
    {
        using var template = new XLTemplate(workbook);
        template.AddVariable("Items", Items(itemCount));
        return template.Generate();
    }

    private static IXLChartSeries Series(IXLWorkbook workbook) =>
        workbook.Worksheet("Report").Charts.Single().Series.Single();

    [Test]
    public async Task ASeriesOnTheRepeatedRowStretchesOverEveryGeneratedRow()
    {
        using var workbook = Template();

        Generate(workbook);

        await Assert.That(Series(workbook).ValueReferences).IsEqualTo("Report!$B$3:$B$6");
    }

    [Test]
    public async Task TheCategoryReferenceStretchesTheSameWay()
    {
        using var workbook = Template();

        Generate(workbook);

        await Assert.That(Series(workbook).CategoryReferences).IsEqualTo("Report!$A$3:$A$6");
    }

    /// <summary>A series covering the heading as well keeps covering it.</summary>
    [Test]
    public async Task ASeriesStartingAboveTheRangeKeepsItsStart()
    {
        using var workbook = Template(valueReferences: "Report!$B$2:$B$3", categoryReferences: null);

        Generate(workbook);

        await Assert.That(Series(workbook).ValueReferences).IsEqualTo("Report!$B$2:$B$6");
    }

    [Test]
    public async Task ASeriesBelowTheRangeMovesDownWithTheRowsItPlots()
    {
        using var workbook = Template(valueReferences: "Report!$B$10:$B$12", categoryReferences: null);

        Generate(workbook);

        // Two rows net: three were inserted for the second, third and fourth items, and the empty
        // options row was removed again.
        await Assert.That(Series(workbook).ValueReferences).IsEqualTo("Report!$B$12:$B$14");
    }

    /// <summary>
    /// The same, but with a total in the options row so that the row survives. Everything below the
    /// range then moves one further than it does when the options row is dropped — which the ledger
    /// has to account for, and did not until this test said so.
    /// </summary>
    [Test]
    public async Task ASeriesBelowARangeWhoseOptionsRowSurvivesMovesOneFurther()
    {
        using var workbook = Template(valueReferences: "Report!$B$10:$B$12", categoryReferences: null);
        workbook.Worksheet("Report").Cell("B4").Value = "<<Sum>>";

        Generate(workbook);

        // Four generated rows plus the options row holding the total, where the template had two
        // rows: three more than before.
        await Assert.That(Series(workbook).ValueReferences).IsEqualTo("Report!$B$13:$B$15");
    }

    [Test]
    public async Task ASeriesAboveTheRangeIsLeftAlone()
    {
        using var workbook = Template(valueReferences: "Report!$B$1:$B$2", categoryReferences: null);

        Generate(workbook);

        await Assert.That(Series(workbook).ValueReferences).IsEqualTo("Report!$B$1:$B$2");
    }

    [Test]
    public async Task ASeriesOnAnotherSheetIsLeftAlone()
    {
        using var workbook = Template(valueReferences: "Elsewhere!$B$3:$B$3", categoryReferences: null);
        workbook.AddWorksheet("Elsewhere");

        Generate(workbook);

        await Assert.That(Series(workbook).ValueReferences).IsEqualTo("Elsewhere!$B$3:$B$3");
    }

    /// <summary>A chart may sit anywhere; what matters is the sheet its references name.</summary>
    [Test]
    public async Task AChartOnAnotherSheetStillFollowsTheDataItPlots()
    {
        using var workbook = Template();
        var dashboard = workbook.AddWorksheet("Dashboard");
        var chart = dashboard.Charts.Add(XLChartType.Line);
        chart.Series.Add("Quantity", "Report!$B$3:$B$3");

        Generate(workbook);

        await Assert.That(dashboard.Charts.Single().Series.Single().ValueReferences)
            .IsEqualTo("Report!$B$3:$B$6");
    }

    [Test]
    public async Task AnUnqualifiedReferenceIsReadAgainstTheChartsOwnSheet()
    {
        using var workbook = Template(valueReferences: "$B$3:$B$3", categoryReferences: null);

        Generate(workbook);

        await Assert.That(Series(workbook).ValueReferences).IsEqualTo("$B$3:$B$6");
    }

    [Test]
    public async Task AReferenceInAColumnBesideTheRangeIsNotStretched()
    {
        using var workbook = Template(valueReferences: "Report!$E$3:$E$3", categoryReferences: null);

        Generate(workbook);

        await Assert.That(Series(workbook).ValueReferences).IsEqualTo("Report!$E$3:$E$3");
    }

    /// <summary>
    /// A column beside the range is not stretched by it, but the rows below the range still move,
    /// because the insert that grew the range was a full-row one. A reference that starts above the
    /// range and ends below it therefore keeps its start and follows the rows its tail sat on.
    /// </summary>
    [Test]
    public async Task AReferenceBesideTheRangeSpanningItFollowsTheRowsBelowIt()
    {
        using var workbook = Template(valueReferences: "Report!$E$1:$E$10", categoryReferences: null);

        Generate(workbook);

        // Two rows net, the same shift a reference wholly below the range gets.
        await Assert.That(Series(workbook).ValueReferences).IsEqualTo("Report!$E$1:$E$12");
    }

    /// <summary>An empty collection shrinks the range, and the series shrinks with it.</summary>
    [Test]
    public async Task ASeriesOverAnEmptyRangeCollapses()
    {
        using var workbook = Template();

        using (var template = new XLTemplate(workbook))
        {
            template.AddVariable("Items", new List<SaleItem>());
            template.Generate();
        }

        // The rendered area is empty, so the series has nothing left to plot and says so rather
        // than pointing at rows that are now something else.
        await Assert.That(Series(workbook).ValueReferences).IsEqualTo("Report!$B$3:$B$3");
    }

    /// <summary>
    /// The re-pointed reference has to survive the file format, which for a chart read from a
    /// template means the patcher writing it back into the existing chart part.
    /// </summary>
    [Test]
    public async Task ARewrittenReferenceSurvivesASaveAndReload()
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
            template.AddVariable("Items", Items(4));
            template.Generate();
            template.SaveAs(generated);
        }

        generated.Position = 0;
        using var reloaded = new XLWorkbook(generated);
        await Assert.That(Series(reloaded).ValueReferences).IsEqualTo("Report!$B$3:$B$6");
        await Assert.That(Series(reloaded).CategoryReferences).IsEqualTo("Report!$A$3:$A$6");
    }

    /// <summary>
    /// With two ranges on a sheet, a reference into the lower one is both moved by what the upper
    /// one did and stretched by what its own did — which only comes out right if the expansions are
    /// replayed in the order they ran.
    /// </summary>
    [Test]
    public async Task AReferenceIsMovedByAnEarlierRangeAndStretchedByItsOwn()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        sheet.Cell("A1").Value = "{{ item.Product }}";
        sheet.Cell("B1").Value = "{{ item.Quantity }}";
        sheet.Cell("A4").Value = "{{ item.Product }}";
        sheet.Cell("B4").Value = "{{ item.Quantity }}";
        workbook.DefinedNames.Add("First", sheet.Range("A1:B2"));
        workbook.DefinedNames.Add("Second", sheet.Range("A4:B5"));

        var chart = sheet.Charts.Add(XLChartType.ColumnClustered);
        chart.Series.Add("Second", "Report!$B$4:$B$4");

        using (var template = new XLTemplate(workbook))
        {
            template.AddVariable("First", Items(3));
            template.AddVariable("Second", Items(2));
            template.Generate();
        }

        // The first range grew from one row to three and dropped its options row, moving the second
        // range down one; the second then grew from one row to two, stretching the series over both.
        await Assert.That(Series(workbook).ValueReferences).IsEqualTo("Report!$B$5:$B$6");
    }

    [Test]
    public async Task AChartInATemplateWithNoBoundRangeIsUntouched()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        sheet.Cell("A1").Value = "{{ Company }}";
        sheet.Cell("B1").Value = 5;

        var chart = sheet.Charts.Add(XLChartType.ColumnClustered);
        chart.Series.Add("Values", "Report!$B$1:$B$1");

        using (var template = new XLTemplate(workbook))
        {
            template.AddVariable("Company", "Contoso");
            template.Generate();
        }

        await Assert.That(Series(workbook).ValueReferences).IsEqualTo("Report!$B$1:$B$1");
    }
}
