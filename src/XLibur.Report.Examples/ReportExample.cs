using System;
using System.Collections.Generic;
using System.IO;
using XLibur.Excel;

namespace XLibur.Report.Examples;

/// <summary>
/// One worked example: a template it authors, the data it binds, and the report that comes out.
/// </summary>
/// <remarks>
/// <para>
/// Every example writes <em>both</em> workbooks, named so they sort next to each other. Opening the
/// pair is the point — a template language is much easier to read as a before and after than as a
/// listing, and the template is where the interesting part of each example lives.
/// </para>
/// <para>
/// Override <see cref="BuildTemplate"/> to author the template and <see cref="AddData"/> to bind the
/// variables it refers to. Everything else — saving, validating, reporting what happened — is the
/// same for every example and is done here.
/// </para>
/// </remarks>
public abstract class ReportExample
{
    /// <summary>The example's name, used for its file names and in the menu.</summary>
    public abstract string Name { get; }

    /// <summary>One line saying what the example shows.</summary>
    public abstract string Summary { get; }

    /// <summary>
    /// Authors the template. Write it exactly as a report author would in Excel: placeholder
    /// expressions in cells, a defined name over the rows to repeat, tags in the options row.
    /// </summary>
    protected abstract void BuildTemplate(IXLWorkbook workbook);

    /// <summary>Binds the variables the template refers to.</summary>
    protected abstract void AddData(IXLTemplate template);

    /// <summary>
    /// Anything worth pointing out about the generated workbook that reading the code does not tell
    /// you. Printed under the example when it runs.
    /// </summary>
    protected virtual void Describe(IXLWorkbook generated, TextWriter output)
    {
    }

    /// <summary>
    /// Writes the template and the report it generates into <paramref name="directory"/>, and returns
    /// what happened.
    /// </summary>
    /// <param name="directory">Where to write the pair. Created if it is not there.</param>
    /// <param name="output">Where to print the example's notes, or <c>null</c> for nowhere.</param>
    public ExampleRun Run(string directory, TextWriter? output = null)
    {
        Directory.CreateDirectory(directory);

        var templatePath = Path.Combine(directory, Name + "-1-template.xlsx");
        var reportPath = Path.Combine(directory, Name + "-2-report.xlsx");

        using (var workbook = new XLWorkbook())
        {
            BuildTemplate(workbook);

            // validate: true runs the file through the OpenXML schema validator on the way out. An
            // example that produces something Excel would refuse should fail here, loudly, rather
            // than in the reader's hands.
            workbook.SaveAs(templatePath, validate: true);
        }

        using var template = new XLTemplate(templatePath);
        AddData(template);

        var result = template.Generate();

        template.Workbook.SaveAs(reportPath, validate: true);

        if (output is not null)
        {
            Report(result, template.Workbook, output);
        }

        return new ExampleRun(Name, templatePath, reportPath, result);
    }

    private void Report(XLGenerateResult result, IXLWorkbook generated, TextWriter output)
    {
        output.WriteLine(Name);
        output.WriteLine(new string('─', Name.Length));
        output.WriteLine("  " + Summary);
        output.WriteLine();

        Describe(generated, output);

        if (result.HasErrors)
        {
            output.WriteLine($"  {result.ParsingErrors.Count} error(s) reported:");
            foreach (var error in result.ParsingErrors)
            {
                output.WriteLine("    " + error);
            }
        }

        output.WriteLine();
    }

    /// <summary>
    /// Adds a heading row, styled the way a report author would. Every example that has a table of
    /// anything wants this, and it is not what any of them is about.
    /// </summary>
    protected static void Headings(IXLWorksheet sheet, int row, params string[] headings)
    {
        for (var i = 0; i < headings.Length; i++)
        {
            sheet.Cell(row, i + 1).Value = headings[i];
        }

        sheet.Range(row, 1, row, headings.Length).Style
            .Font.SetBold()
            .Fill.SetBackgroundColor(XLColor.LightGray)
            .Border.SetBottomBorder(XLBorderStyleValues.Thin);
    }
}

/// <summary>What one example produced.</summary>
/// <param name="Name">The example's name.</param>
/// <param name="TemplatePath">The template it authored.</param>
/// <param name="ReportPath">The report it generated from that template.</param>
/// <param name="Result">What generation reported.</param>
public sealed record ExampleRun(string Name, string TemplatePath, string ReportPath, XLGenerateResult Result);

/// <summary>Every example, in the order a reader should meet them.</summary>
public static class AllExamples
{
    /// <summary>
    /// The examples, simplest first. The annual sales report is the flagship and comes after the
    /// smallest possible template, so that a reader has seen the shape before seeing it used in
    /// earnest.
    /// </summary>
    public static IReadOnlyList<ReportExample> Ordered { get; } = new ReportExample[]
    {
        new MinimalReport(),
        new AnnualSalesReport(),
        new ExcelFunctionsInExpressions(),
        new ConditionalRows(),
        new HorizontalReport(),
        new CustomTagRegistration(),
        new ErrorsAreReportedNotThrown(),
        new ChartOverGeneratedRows(),
        new PivotOverGeneratedRows(),
        new EverythingAtOnce(),
    };

    /// <summary>The example of the given name, or <c>null</c>.</summary>
    public static ReportExample? ByName(string name)
    {
        foreach (var example in Ordered)
        {
            if (string.Equals(example.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                return example;
            }
        }

        return null;
    }
}
