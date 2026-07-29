using System;
using XLibur.Excel;

namespace XLibur.Report.Tests.Infrastructure;

/// <summary>
/// One golden-file case: how to build its template, what data to bind, and what the generated
/// workbook should look like.
/// </summary>
/// <remarks>
/// Both the template and the expectation are defined as code. The template is additionally
/// committed as an <c>.xlsx</c> so the suite exercises a workbook that has been through the file
/// format; the expectation never needs to be, because comparing two in-memory workbooks says
/// everything a second binary would.
/// </remarks>
public sealed class ReportFixture
{
    /// <summary>Creates a fixture. <paramref name="name"/> names its committed template.</summary>
    public ReportFixture(
        string name,
        Action<IXLWorkbook> buildTemplate,
        Action<IXLTemplate> bind,
        Action<IXLWorkbook> buildExpected)
    {
        Name = name ?? throw new ArgumentNullException(nameof(name));
        BuildTemplate = buildTemplate ?? throw new ArgumentNullException(nameof(buildTemplate));
        Bind = bind ?? throw new ArgumentNullException(nameof(bind));
        BuildExpected = buildExpected ?? throw new ArgumentNullException(nameof(buildExpected));
    }

    /// <summary>The fixture's name, which is also its committed template's file name.</summary>
    public string Name { get; }

    /// <summary>Builds the template workbook.</summary>
    public Action<IXLWorkbook> BuildTemplate { get; }

    /// <summary>Adds the variables the template binds.</summary>
    public Action<IXLTemplate> Bind { get; }

    /// <summary>Builds the workbook generation is expected to produce.</summary>
    public Action<IXLWorkbook> BuildExpected { get; }

    /// <summary>Which dimensions the comparison checks.</summary>
    public WorkbookComparisonOptions Options { get; init; } = WorkbookComparisonOptions.Default;

    /// <summary>Whether generation is expected to report errors.</summary>
    public bool ExpectsErrors { get; init; }
}
