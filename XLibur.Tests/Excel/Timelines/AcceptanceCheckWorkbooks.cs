using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Timelines;

/// <summary>
/// Writes the workbooks a human opens in Excel to settle acceptance criteria 3, 4 and 7.
/// </summary>
/// <remarks>
/// Skipped in normal runs. Nothing here asserts anything: the whole point is that the automated
/// suite cannot see the failure these files exist to catch. PRD 5's central finding was that every
/// automated gate passed on a slicer feature Excel refused to render.
/// </remarks>
public class AcceptanceCheckWorkbooks
{
    private const string OutputDirectory = @"..\..\..\..\scratchpad\ac-check-timelines";

    [Test]
    [Skip("Generator, not a gate. Run by name when check workbooks are wanted.")]
    public async Task Generate()
    {
        Directory.CreateDirectory(OutputDirectory);
        var sha = Environment.GetEnvironmentVariable("AC_SHA") ?? "unstamped";

        WriteCreatedTimeline($"ac3-created-timeline-{sha}.xlsx");
        WriteSecondTimeline($"ac4-timeline-beside-an-existing-one-{sha}.xlsx");
        WriteCascade($"ac7-pivot-deleted-timeline-cascades-{sha}.xlsx");

        await Assert.That(Directory.GetFiles(OutputDirectory, "*.xlsx").Length).IsGreaterThanOrEqualTo(3);
    }

    /// <summary>
    /// Criterion 3: a created timeline opens, is drawn where it was put, and filters. Also carries a
    /// slicer on the same pivot table so the sheet exercises both control types together — the one
    /// combination no automated test covers.
    /// </summary>
    private static void WriteCreatedTimeline(string fileName)
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet("Data");

        data.Cell("A1").Value = "Date";
        data.Cell("B1").Value = "Region";
        data.Cell("C1").Value = "Amount";

        var start = new DateTime(2024, 1, 15);
        for (var i = 0; i < 24; i++)
        {
            data.Cell(i + 2, 1).Value = start.AddDays(i * 11);
            data.Cell(i + 2, 2).Value = i % 2 == 0 ? "North" : "South";
            data.Cell(i + 2, 3).Value = 100 + (i * 7);
        }

        data.Column(1).Style.DateFormat.Format = "yyyy-mm-dd";

        var pivotSheet = wb.AddWorksheet("Pivot");
        var pivotTable = pivotSheet.PivotTables.Add("SalesPivot", pivotSheet.Cell("A3"), data.Range("A1:C25"));
        pivotTable.RowLabels.Add("Region");
        pivotTable.Values.Add("Amount");

        var timeline = pivotSheet.Timelines.Add(pivotTable, "Date");
        timeline.Caption = "Pick a period";
        timeline.Style = "TimeSlicerStyleLight2";
        timeline.Position = pivotSheet.Cell("E3");

        var slicer = pivotSheet.Slicers.Add(pivotTable, "Region");
        slicer.Position = pivotSheet.Cell("E20");

        wb.SaveAs(Path.Combine(OutputDirectory, fileName));
    }

    /// <summary>Criterion 4: a second timeline must not stop Excel drawing the first.</summary>
    private static void WriteSecondTimeline(string fileName)
    {
        using var source = TestHelper.GetStreamFromResource(
            TestHelper.GetResourcePath(@"TryToLoad\Timelines_Missing_21232.xlsx"));

        using var wb = new XLWorkbook(source);
        var pivotSheet = wb.Worksheet("Pivot");
        var added = pivotSheet.Timelines.Add(pivotSheet.PivotTables.Single(), "Date");
        added.Caption = "Added by XLibur";
        added.Position = pivotSheet.Cell("C12");

        wb.SaveAs(Path.Combine(OutputDirectory, fileName));
    }

    /// <summary>Criterion 7: deleting the pivot table leaves no orphan for Excel to repair.</summary>
    private static void WriteCascade(string fileName)
    {
        using var source = TestHelper.GetStreamFromResource(
            TestHelper.GetResourcePath(@"TryToLoad\Timelines_Missing_21232.xlsx"));

        using var wb = new XLWorkbook(source);
        var pivotSheet = wb.Worksheet("Pivot");
        pivotSheet.PivotTables.Delete(pivotSheet.PivotTables.Single().Name);

        wb.SaveAs(Path.Combine(OutputDirectory, fileName));
    }
}
