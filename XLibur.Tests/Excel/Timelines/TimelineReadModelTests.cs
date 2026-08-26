using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Timelines;

/// <summary>
/// The timeline read model, read against the Excel-authored fixture the round-trip suite already
/// uses.
/// </summary>
/// <remarks>
/// <c>Resource/TryToLoad/Timelines_Missing_21232.xlsx</c> carries one timeline on the <c>Date</c>
/// field of the pivot table on sheet <c>Pivot</c>. Its names are Russian, which is a feature rather
/// than an inconvenience: a reader that assumed an English cache-name convention would fail here.
/// The timeline is unfiltered — <c>filterType="unknown"</c> and no <c>x15:selection</c> — so it
/// exercises the bounds path and pins the "no selection" case.
/// </remarks>
public class TimelineReadModelTests
{
    private const string Fixture = @"TryToLoad\Timelines_Missing_21232.xlsx";

    [Test]
    public async Task The_worksheet_owns_the_timeline_drawn_on_it()
    {
        using var wb = Load();

        await Assert.That(wb.Worksheet("Pivot").Timelines.Count).IsEqualTo(1);
        await Assert.That(wb.Worksheet("Data").Timelines.Count).IsEqualTo(0);
    }

    [Test]
    public async Task A_timeline_reports_what_the_file_says()
    {
        using var wb = Load();

        var timeline = wb.Worksheet("Pivot").Timelines.Single();

        await Assert.That(timeline.Name).IsEqualTo("Date");
        await Assert.That(timeline.Caption).IsEqualTo("Date");
        await Assert.That(timeline.SourceFieldName).IsEqualTo("Date");
        await Assert.That(timeline.Level).IsEqualTo(XLTimelineLevel.Months);

        // Absent booleans default to true, which is what Excel means by omitting them.
        await Assert.That(timeline.ShowHeader).IsTrue();
        await Assert.That(timeline.ShowSelectionLabel).IsTrue();

        // The fixture writes no style attribute at all.
        await Assert.That(timeline.Style).IsNull();
    }

    [Test]
    public async Task A_timeline_binds_to_the_pivot_table_its_cache_names()
    {
        using var wb = Load();

        var timeline = wb.Worksheet("Pivot").Timelines.Single();

        await Assert.That(timeline.PivotTables.Select(pt => pt.Name))
            .IsEquivalentTo(new[] { "СводнаяТаблица2" });
        await Assert.That(timeline.Worksheet.Name).IsEqualTo("Pivot");
    }

    [Test]
    public async Task An_unfiltered_timeline_reports_its_bounds_and_no_selection()
    {
        using var wb = Load();

        var timeline = wb.Worksheet("Pivot").Timelines.Single();

        // The bounds are the date field's range rounded outward to whole years.
        await Assert.That(timeline.BoundsStart).IsEqualTo(new DateTime(1998, 1, 1));
        await Assert.That(timeline.BoundsEnd).IsEqualTo(new DateTime(2005, 1, 1));

        await Assert.That(timeline.HasSelection).IsFalse();
        await Assert.That(timeline.SelectionStart).IsNull();
        await Assert.That(timeline.SelectionEnd).IsNull();
    }

    [Test]
    public async Task A_timeline_reports_where_it_is_drawn()
    {
        using var wb = Load();

        // The fixture anchors the frame at xdr:col 2, xdr:row 1 — zero-based, so C2.
        var timeline = wb.Worksheet("Pivot").Timelines.Single();

        await Assert.That(timeline.Position.Address.ToString()).IsEqualTo("C2");
    }

    [Test]
    public async Task A_pivot_table_views_the_timelines_that_filter_it()
    {
        using var wb = Load();

        var pivotTable = wb.Worksheet("Pivot").PivotTables.Single();

        await Assert.That(pivotTable.Timelines.Select(t => t.Name)).IsEquivalentTo(new[] { "Date" });
    }

    [Test]
    public async Task A_timeline_can_be_found_by_name()
    {
        using var wb = Load();

        var timelines = wb.Worksheet("Pivot").Timelines;

        await Assert.That(timelines.Timeline("Date").Caption).IsEqualTo("Date");
        await Assert.That(timelines.TryGetTimeline("Date", out var found)).IsTrue();
        await Assert.That(found!.Name).IsEqualTo("Date");
        await Assert.That(timelines.TryGetTimeline("Nope", out _)).IsFalse();
    }

    [Test]
    public async Task Reading_a_timeline_does_not_rewrite_its_parts()
    {
        // The regression gate for the whole read model. Timeline parts survive a round trip because
        // nothing opens them; reaching one through TimeLinePart.Timelines would attach a DOM the SDK
        // writes back over the original bytes on save, taking mc:Ignorable and every attribute
        // XLibur does not model with it. The reader streams the parts detached instead.
        using var original = Resource();
        var before = PartBytes(original, "xl/timelines/timeline1.xml");
        var beforeCache = PartBytes(original, "xl/timelineCaches/timelineCache1.xml");

        using var saved = LoadAndSave();

        await Assert.That(PartBytes(saved, "xl/timelines/timeline1.xml")).IsEquivalentTo(before, CollectionOrdering.Matching);
        await Assert.That(PartBytes(saved, "xl/timelineCaches/timelineCache1.xml")).IsEquivalentTo(beforeCache, CollectionOrdering.Matching);
    }

    [Test]
    public async Task Timelines_still_load_after_a_round_trip()
    {
        using var saved = LoadAndSave();
        saved.Position = 0;
        using var wb = new XLWorkbook(saved);

        await Assert.That(wb.Worksheet("Pivot").Timelines.Single().Level).IsEqualTo(XLTimelineLevel.Months);
    }

    #region Helpers

    /// <summary>
    /// The fixture, opened over a copy that outlives this call. The workbook reads its original
    /// stream again on save, so the stream cannot be disposed when this returns.
    /// </summary>
    private static XLWorkbook Load()
    {
        var stream = Resource();
        stream.Position = 0;
        return new XLWorkbook(stream);
    }

    private static MemoryStream Resource()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(Fixture));
        var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms;
    }

    private static MemoryStream LoadAndSave()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(Fixture));
        var ms = new MemoryStream();

        using (var wb = new XLWorkbook(stream))
            wb.SaveAs(ms);

        return ms;
    }

    private static byte[] PartBytes(MemoryStream package, string partPath)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        var entry = archive.Entries.First(e =>
            e.FullName.Equals(partPath, StringComparison.OrdinalIgnoreCase));

        using var entryStream = entry.Open();
        using var buffer = new MemoryStream();
        entryStream.CopyTo(buffer);
        return buffer.ToArray();
    }

    #endregion
}
