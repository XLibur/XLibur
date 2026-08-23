using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
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
    /// The directory the committed golden fixtures live in, resolved from the compile-time path of
    /// this file rather than from the working directory, so that the fixtures a failing run writes
    /// land next to the test rather than under <c>bin</c>.
    /// </summary>
    internal static string GoldenDirectory([CallerFilePath] string sourceFile = "") =>
        Path.Combine(Path.GetDirectoryName(sourceFile)!, "Golden");
}
