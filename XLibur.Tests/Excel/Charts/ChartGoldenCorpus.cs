using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Charts;

/// <summary>
/// Captures the raw chart-part XML of a workbook built in memory, which is the gate every task of
/// spec 22 is measured with: the chart tree is reorganised without the bytes it writes changing.
/// </summary>
internal static class ChartGoldenCorpus
{
    /// <summary>
    /// Saves a workbook built by <paramref name="build"/> and returns the raw XML of its first chart
    /// part, verbatim. Byte-identity of this string across a refactor is what proves the refactor
    /// changed no output.
    /// </summary>
    internal static string CaptureChartPartXml(Action<IXLWorksheet> build)
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Q1";
            ws.Cell("A2").Value = "Q2";
            ws.Cell("B1").Value = 100;
            ws.Cell("B2").Value = 200;
            ws.Cell("C1").Value = 5;
            ws.Cell("C2").Value = 8;
            build(ws);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        return FirstChartPartXml(ms);
    }

    /// <summary>Reads the raw XML of the first chart part of an already saved workbook.</summary>
    internal static string FirstChartPartXml(Stream saved)
    {
        saved.Position = 0;
        using var doc = SpreadsheetDocument.Open(saved, false);
        var chartPart = doc.WorkbookPart!
            .WorksheetParts
            .SelectMany(p => p.DrawingsPart is null
                ? Enumerable.Empty<ChartPart>()
                : p.DrawingsPart.ChartParts)
            .First();

        using var stream = chartPart.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    /// <summary>
    /// The committed fixture of the given name, or <c>null</c> when the corpus does not hold one.
    /// </summary>
    /// <remarks>
    /// The fixtures travel as embedded resources rather than as files beside the test. A release
    /// build normalises compile-time paths to <c>/_/</c>, so nothing may resolve a fixture from
    /// where its source file was; and nothing copies <c>Golden/</c> next to the test binary.
    /// </remarks>
    internal static string? ReadGolden(string name)
    {
        var assembly = typeof(ChartGoldenCorpus).Assembly;
        var resource = ResourceNameFor(assembly, name);
        if (resource == null)
            return null;

        using var stream = assembly.GetManifestResourceStream(resource)!;
        using var reader = new StreamReader(stream);
        return Normalise(reader.ReadToEnd());
    }

    /// <summary>
    /// Whether this run is allowed to write fixtures back to the source tree, which only an
    /// explicit <c>XLIBUR_WRITE_CHART_GOLDEN=1</c> turns on.
    /// </summary>
    /// <remarks>
    /// Regenerating is deliberately not automatic. A fixture the corpus is missing has to fail the
    /// run, or CI would quietly write and then assert its own output, which gates nothing.
    /// </remarks>
    internal static bool CanWriteGolden =>
        Environment.GetEnvironmentVariable("XLIBUR_WRITE_CHART_GOLDEN") == "1";

    /// <summary>
    /// Writes a fixture into the source tree, for a developer regenerating the corpus. The written
    /// file is only picked up by the next build, because the corpus is embedded.
    /// </summary>
    internal static void WriteGolden(string name, string xml)
    {
        var directory = Path.Combine(SourceDirectory(), "Excel", "Charts", "Golden");
        Directory.CreateDirectory(directory);
        File.WriteAllText(Path.Combine(directory, name + ".xml"), Normalise(xml));
    }

    /// <summary>
    /// Line endings are not part of what the corpus pins — the chart XML it holds is a single line —
    /// but <c>.gitattributes</c> checks every text file out as CRLF, so a fixture is compared for
    /// what it says rather than for how the working copy stores it.
    /// </summary>
    internal static string Normalise(string xml) => xml.Replace("\r\n", "\n");

    private static string? ResourceNameFor(Assembly assembly, string name)
    {
        var suffix = $".Golden.{name}.xml";
        return Array.Find(assembly.GetManifestResourceNames(),
            resource => resource.EndsWith(suffix, StringComparison.Ordinal));
    }

    /// <summary>
    /// The <c>XLibur.Tests</c> project directory, found by walking up from the test binary. Used
    /// only when regenerating, so a build that cannot find it is never a test failure.
    /// </summary>
    private static string SourceDirectory()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "XLibur.Tests.csproj")))
                return directory.FullName;

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException(
            $"XLibur.Tests.csproj is not above {AppContext.BaseDirectory}, so there is nowhere to " +
            "write the golden fixtures. Regenerate them from a normal local build.");
    }
}
