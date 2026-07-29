using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using XLibur.Excel;
using XLibur.Report.Expressions;

namespace XLibur.Report.Tests.Infrastructure;

/// <summary>
/// Runs a <see cref="ReportFixture"/>: generates its template and compares the result against the
/// fixture's expectation.
/// </summary>
public static class GoldenFile
{
    /// <summary>
    /// Verifies <paramref name="fixture"/>, throwing with every difference found. The workbook
    /// actually produced is left in <see cref="ReportResources.DiagnosticsDirectory"/> so a
    /// failure can be opened in Excel.
    /// </summary>
    public static void Verify(ReportFixture fixture, IExpressionEngine? engine = null)
    {
        if (ReportResources.Regenerating || !ReportResources.TemplateExists(fixture.Name))
        {
            Regenerate(fixture);
        }

        using var template = LoadTemplate(fixture);
        AssertCommittedTemplateIsCurrent(fixture, template);

        using var report = new XLTemplate(template, engine);
        fixture.Bind(report);
        var result = report.Generate();

        AssertErrorExpectation(fixture, result, template);

        using var expected = new XLWorkbook();
        fixture.BuildExpected(expected);

        var differences = WorkbookComparer.Compare(expected, template, fixture.Options);
        if (differences.Count == 0)
        {
            return;
        }

        var actualPath = ReportResources.WriteDiagnostic(fixture.Name + "-actual", template);
        var expectedPath = ReportResources.WriteDiagnostic(fixture.Name + "-expected", expected);

        throw new GoldenFileMismatchException(Describe(fixture.Name, differences, actualPath, expectedPath));
    }

    private static IXLWorkbook LoadTemplate(ReportFixture fixture)
    {
        using var stream = ReportResources.OpenTemplate(fixture.Name);
        return new XLWorkbook(stream);
    }

    private static void Regenerate(ReportFixture fixture)
    {
        using var workbook = new XLWorkbook();
        fixture.BuildTemplate(workbook);
        ReportResources.WriteTemplate(fixture.Name, workbook);
    }

    /// <summary>
    /// Guards against the committed template drifting from the code that defines it — the code is
    /// the source of truth, and a stale binary would quietly test the wrong thing.
    /// </summary>
    private static void AssertCommittedTemplateIsCurrent(ReportFixture fixture, IXLWorkbook committed)
    {
        using var fromCode = new XLWorkbook();
        fixture.BuildTemplate(fromCode);

        // Compared before generation runs, so `committed` is still the untouched template.
        var differences = WorkbookComparer.Compare(fromCode, committed, fixture.Options);
        if (differences.Count == 0)
        {
            return;
        }

        throw new GoldenFileMismatchException(
            $"The committed template '{fixture.Name}.xlsx' no longer matches the code that defines it. " +
            $"Re-run with XLIBUR_REPORT_REGEN=1 to rewrite it.{Environment.NewLine}" +
            string.Join(Environment.NewLine, differences));
    }

    private static void AssertErrorExpectation(ReportFixture fixture, XLGenerateResult result, IXLWorkbook produced)
    {
        if (fixture.ExpectsErrors == result.HasErrors)
        {
            return;
        }

        if (result.HasErrors)
        {
            var path = ReportResources.WriteDiagnostic(fixture.Name + "-actual", produced);
            var errors = new StringBuilder();
            foreach (var error in result.ParsingErrors)
            {
                errors.AppendLine("  " + error);
            }

            throw new GoldenFileMismatchException(
                $"'{fixture.Name}' generated with errors:{Environment.NewLine}{errors}Produced workbook: {path}");
        }

        throw new GoldenFileMismatchException($"'{fixture.Name}' was expected to generate with errors, but did not.");
    }

    private static string Describe(string name, IReadOnlyList<string> differences, string actualPath, string expectedPath)
    {
        var message = new StringBuilder();
        message.AppendLine($"'{name}' does not match its expectation ({differences.Count} difference(s)):");

        foreach (var difference in differences)
        {
            message.AppendLine("  " + difference);
        }

        message.AppendLine();
        message.AppendLine("Actual:   " + actualPath);
        message.AppendLine("Expected: " + expectedPath);
        return message.ToString();
    }
}

/// <summary>Thrown when a generated workbook does not match its fixture's expectation.</summary>
public sealed class GoldenFileMismatchException : Exception
{
    /// <inheritdoc cref="GoldenFileMismatchException"/>
    public GoldenFileMismatchException(string message)
        : base(message)
    {
    }
}
