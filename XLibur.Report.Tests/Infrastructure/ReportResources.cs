using System;
using System.IO;
using System.Linq;
using XLibur.Excel;

namespace XLibur.Report.Tests.Infrastructure;

/// <summary>
/// Finds the committed template workbooks, and rewrites them when a fixture's defining code
/// changes.
/// </summary>
/// <remarks>
/// Templates are authored in C# but committed as <c>.xlsx</c>, for two reasons: the suite should
/// exercise a template that has been through the file format the way a user's would, and a
/// fixture should be reviewable as code rather than as an opaque binary. Set
/// <c>XLIBUR_REPORT_REGEN=1</c> to rewrite the committed copies from their defining code.
/// </remarks>
public static class ReportResources
{
    private const string RegenerationVariable = "XLIBUR_REPORT_REGEN";

    /// <summary>Whether this run rewrites the committed templates instead of asserting against them.</summary>
    public static bool Regenerating =>
        string.Equals(Environment.GetEnvironmentVariable(RegenerationVariable), "1", StringComparison.Ordinal);

    /// <summary>Where a failing comparison leaves the workbook it actually produced.</summary>
    public static string DiagnosticsDirectory
    {
        get
        {
            var directory = Path.Combine(AppContext.BaseDirectory, "ReportDiagnostics");
            Directory.CreateDirectory(directory);
            return directory;
        }
    }

    /// <summary>The test project's directory, located by walking up from the build output.</summary>
    public static string SourceDirectory
    {
        get
        {
            var directory = new DirectoryInfo(AppContext.BaseDirectory);

            while (directory is not null)
            {
                if (directory.EnumerateFiles("XLibur.Report.Tests.csproj").Any())
                {
                    return directory.FullName;
                }

                directory = directory.Parent;
            }

            throw new InvalidOperationException(
                "Could not locate the XLibur.Report.Tests project directory from " + AppContext.BaseDirectory);
        }
    }

    /// <summary>Opens a committed template by fixture name.</summary>
    public static Stream OpenTemplate(string name)
    {
        var resourceName = $"XLibur.Report.Tests.Resource.Templates.{name}.xlsx";
        var stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName);

        if (stream is null)
        {
            throw new FileNotFoundException(
                $"Template '{name}' is not committed. Run the suite with {RegenerationVariable}=1 to write it.",
                resourceName);
        }

        return stream;
    }

    /// <summary>Whether a template has been committed for <paramref name="name"/>.</summary>
    public static bool TemplateExists(string name) =>
        System.Reflection.Assembly.GetExecutingAssembly()
            .GetManifestResourceNames()
            .Contains($"XLibur.Report.Tests.Resource.Templates.{name}.xlsx", StringComparer.Ordinal);

    /// <summary>Writes <paramref name="workbook"/> over the committed template for <paramref name="name"/>.</summary>
    public static void WriteTemplate(string name, IXLWorkbook workbook)
    {
        var directory = Path.Combine(SourceDirectory, "Resource", "Templates");
        Directory.CreateDirectory(directory);
        workbook.SaveAs(Path.Combine(directory, name + ".xlsx"));
    }

    /// <summary>Saves <paramref name="workbook"/> into the diagnostics directory and returns its path.</summary>
    public static string WriteDiagnostic(string name, IXLWorkbook workbook)
    {
        var path = Path.Combine(DiagnosticsDirectory, name + ".xlsx");
        workbook.SaveAs(path);
        return path;
    }
}
