using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Report.Examples;

namespace XLibur.Report.Tests;

/// <summary>
/// Runs every worked example, so that one cannot rot into a snippet that no longer compiles or no
/// longer works.
/// </summary>
/// <remarks>
/// <para>
/// The examples are documentation, and documentation that is never executed is documentation that is
/// eventually wrong. These tests are deliberately shallow — they assert that each example generates
/// without complaint and writes the pair of workbooks it promises, not what is in them. What is in
/// them is asserted by the tests of the features the examples demonstrate.
/// </para>
/// <para>
/// Every example runs once for the class rather than once per test: between them they generate a
/// couple of dozen workbooks, and there is nothing in any test here that a shared run makes less
/// truthful.
/// </para>
/// </remarks>
public class ExampleSmokeTests
{
    private static readonly Lazy<IReadOnlyList<ExampleRun>> Runs = new(RunAll);

    /// <summary>
    /// The one example whose purpose is to end with errors. Anything else reporting one is a failure.
    /// </summary>
    private static bool ExpectsErrors(string name) => name == new ErrorsAreReportedNotThrown().Name;

    private static List<ExampleRun> RunAll()
    {
        var directory = Path.Combine(Path.GetTempPath(), "XLiburReportExamples", Guid.NewGuid().ToString("N"));

        return AllExamples.Ordered.Select(example => example.Run(directory)).ToList();
    }

    [Test]
    public async Task EveryExampleGeneratesWithoutErrors()
    {
        var failed = Runs.Value
            .Where(run => run.Result.HasErrors && !ExpectsErrors(run.Name))
            .Select(run => $"{run.Name}: {string.Join("; ", run.Result.ParsingErrors.Select(error => error.ToString()))}")
            .ToList();

        await Assert.That(failed).IsEmpty();
    }

    /// <summary>
    /// Both workbooks, every time: the pair is the example. Each is saved with the OpenXML validator
    /// on, so reaching this point at all means neither file is malformed.
    /// </summary>
    [Test]
    public async Task EveryExampleWritesATemplateAndAReport()
    {
        foreach (var run in Runs.Value)
        {
            await Assert.That(File.Exists(run.TemplatePath)).IsTrue();
            await Assert.That(File.Exists(run.ReportPath)).IsTrue();
            await Assert.That(new FileInfo(run.ReportPath).Length).IsGreaterThan(0);
        }
    }

    /// <summary>
    /// The error-handling example is only worth having if it still produces errors — and if the report
    /// is still generated around them.
    /// </summary>
    [Test]
    public async Task TheErrorExampleReportsErrorsAndGeneratesAnyway()
    {
        var run = Runs.Value.Single(r => ExpectsErrors(r.Name));

        await Assert.That(run.Result.HasErrors).IsTrue();
        await Assert.That(File.Exists(run.ReportPath)).IsTrue();

        using var report = new XLWorkbook(run.ReportPath);
        var sheet = report.Worksheet("Errors");

        // The cells around the broken ones were generated normally, which is the whole claim.
        await Assert.That(sheet.Cell("B7").GetFormattedString()).IsEqualTo("Contoso Ltd");
        await Assert.That(sheet.Cell("A18").GetFormattedString()).IsEqualTo("Rotary hoe");
    }

    /// <summary>Every example is listed, named and described — the menu is the reader's way in.</summary>
    [Test]
    public async Task EveryExampleIsNamedAndDescribed()
    {
        await Assert.That(AllExamples.Ordered).IsNotEmpty();

        foreach (var example in AllExamples.Ordered)
        {
            await Assert.That(example.Name).IsNotNullOrEmpty();
            await Assert.That(example.Summary).IsNotNullOrEmpty();
            await Assert.That(AllExamples.ByName(example.Name)).IsNotNull();
        }

        var names = AllExamples.Ordered.Select(example => example.Name).ToList();
        await Assert.That(names.Distinct(StringComparer.OrdinalIgnoreCase).Count()).IsEqualTo(names.Count);
    }

    /// <summary>
    /// The flagship carries the claims the spec's acceptance criteria make, so they are asserted here
    /// rather than left to a reader opening the file.
    /// </summary>
    [Test]
    public async Task TheAnnualSalesReportShowsWhatItPromises()
    {
        var run = Runs.Value.Single(r => r.Name == new AnnualSalesReport().Name);

        using var report = new XLWorkbook(run.ReportPath);
        var sheet = report.Worksheet("Annual sales");

        // Twelve sales, three region subtotals, a grand total, over a template that had one data row.
        await Assert.That(sheet.RangeUsed()!.RangeAddress.LastAddress.RowNumber).IsEqualTo(21);

        // Two rules in the template and two in the report, however many rows it generated. One per row
        // is the upstream behaviour this exists to not have.
        await Assert.That(sheet.ConditionalFormats.Count()).IsEqualTo(2);

        // The title's three merges, plus one merged label per region.
        await Assert.That(sheet.MergedRanges.Count).IsEqualTo(6);

        // Grouped, so the data rows are outlined one level in from the subtotals.
        await Assert.That(sheet.Rows().Max(row => row.OutlineLevel)).IsEqualTo(1);
    }
}
