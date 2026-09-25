using System;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Timelines;

/// <summary>
/// The one control-name namespace a slicer and a timeline in the same workbook share.
/// </summary>
public class SlicerAndTimelineNameTests
{
    [Test]
    public async Task A_timeline_does_not_take_a_name_a_slicer_already_has()
    {
        using var wb = new XLWorkbook();
        var data = AddData(wb);
        var pivotSheet = wb.AddWorksheet("Pivot");
        var pivotTable = pivotSheet.PivotTables.Add("Sales", pivotSheet.Cell("A3"), data.Range("A1:C25"));

        var slicer = pivotSheet.Slicers.Add(pivotTable, "Date");
        var timeline = data.Timelines.Add(pivotTable, "Date");

        await Assert.That(slicer.Name).IsEqualTo("Date");
        await Assert.That(timeline.Name).IsEqualTo("Date 1");
    }

    [Test]
    public async Task A_slicer_does_not_take_a_name_a_timeline_already_has()
    {
        using var wb = new XLWorkbook();
        var data = AddData(wb);
        var pivotSheet = wb.AddWorksheet("Pivot");
        var pivotTable = pivotSheet.PivotTables.Add("Sales", pivotSheet.Cell("A3"), data.Range("A1:C25"));

        var timeline = pivotSheet.Timelines.Add(pivotTable, "Date");
        var slicer = data.Slicers.Add(pivotTable, "Date");

        await Assert.That(timeline.Name).IsEqualTo("Date");
        await Assert.That(slicer.Name).IsEqualTo("Date 1");
    }

    private static IXLWorksheet AddData(XLWorkbook wb)
    {
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

        return data;
    }
}
