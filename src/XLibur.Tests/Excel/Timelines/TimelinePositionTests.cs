using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Timelines;

/// <summary>
/// Where a timeline sits, which is in the sheet's drawing part rather than in the timeline's own.
/// </summary>
public class TimelinePositionTests
{
    private const string Fixture = @"TryToLoad\Timelines_Missing_21232.xlsx";

    [Test]
    public async Task A_created_timeline_lands_clear_of_the_pivot_table()
    {
        using var wb = Load();

        var pivotTable = wb.Worksheet("Pivot").PivotTables.Single();
        var timeline = wb.Worksheet("Data").Timelines.Add(pivotTable, "Date");

        // Two columns right of the pivot table's rightmost column, at its top row. The pivot table
        // occupies A3:B14, so the timeline goes to D3.
        await Assert.That(timeline.Position.Address.ToString()).IsEqualTo("D3");
    }

    [Test]
    public async Task Moving_a_created_timeline_puts_the_anchor_where_it_was_told()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var pivotTable = wb.Worksheet("Pivot").PivotTables.Single();
            var timeline = wb.Worksheet("Data").Timelines.Add(pivotTable, "Date");
            timeline.Position = wb.Worksheet("Data").Cell("F5");
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);

        await Assert.That(reloaded.Worksheet("Data").Timelines.Single().Position.Address.ToString())
            .IsEqualTo("F5");
    }

    [Test]
    public async Task Moving_a_loaded_timeline_edits_its_anchor_rather_than_replacing_it()
    {
        using var original = Resource();
        var timelinePartBefore = original.PartBytes("xl/timelines/timeline1.xml");

        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            wb.Worksheet("Pivot").Timelines.Single().Position = wb.Worksheet("Pivot").Cell("E4");
            wb.SaveAs(saved);
        }

        var drawing = saved.ReadPart("xl/drawings/drawing1.xml");

        // The corner moved: C2 (col 2, row 1) to E4 (col 4, row 3).
        await Assert.That(drawing).Contains("<xdr:col>4</xdr:col>");

        // And Excel's own wrapper survived, which is what replacing the anchor would have destroyed.
        await Assert.That(drawing).Contains("mc:AlternateContent");
        await Assert.That(drawing).Contains("mc:Fallback");

        // A position-only edit must not open the timelines part at all — PatchTimeline's early
        // return on (assigned & ~Position) == None is what this pins. Without it, a part Excel
        // authored would be re-serialised on every move, and nothing above this line would notice.
        await Assert.That(saved.PartBytes("xl/timelines/timeline1.xml")).IsEquivalentTo(timelinePartBefore, CollectionOrdering.Matching);

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("Pivot").Timelines.Single().Position.Address.ToString())
            .IsEqualTo("E4");
    }

    [Test]
    public async Task Moving_a_loaded_timeline_keeps_its_size()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            wb.Worksheet("Pivot").Timelines.Single().Position = wb.Worksheet("Pivot").Cell("E4");
            wb.SaveAs(saved);
        }

        var drawing = saved.ReadPart("xl/drawings/drawing1.xml");

        // Both corners shifted by the same delta — two columns and two rows — so the band is the
        // same size it was. The original spans col 2..8, row 1..9.
        await Assert.That(drawing).Contains("<xdr:col>4</xdr:col>");
        await Assert.That(drawing).Contains("<xdr:col>10</xdr:col>");
        await Assert.That(drawing).Contains("<xdr:row>3</xdr:row>");
        await Assert.That(drawing).Contains("<xdr:row>11</xdr:row>");
    }

    #region Helpers

    private static XLWorkbook Load() => TestHelper.LoadWorkbook(Fixture);

    private static MemoryStream Resource() => TestHelper.OpenResource(Fixture);

    #endregion
}
