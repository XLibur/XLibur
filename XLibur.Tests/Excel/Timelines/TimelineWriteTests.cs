using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Timelines;

/// <summary>
/// Creating timelines, editing loaded ones, and the parts neither operation may touch.
/// </summary>
/// <remarks>
/// A created timeline needs six things or Excel offers to repair the file: the timeline definition,
/// the worksheet's <c>extLst</c> reference to it, the cache part, the workbook's <c>extLst</c>
/// registration of that cache, a <c>#N/A</c> defined name, and a drawing anchor. All six are
/// asserted here.
/// </remarks>
public class TimelineWriteTests
{
    private const string Fixture = @"TryToLoad\Timelines_Missing_21232.xlsx";

    // ── Creating ────────────────────────────────────────────────────────

    [Test]
    public async Task A_created_timeline_writes_all_six_pieces()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var pivotTable = wb.Worksheet("Pivot").PivotTables.Single();
            wb.Worksheet("Data").Timelines.Add(pivotTable, "Date");
            wb.SaveAs(saved);
        }

        var entries = EntryNames(saved);

        // 1 and 2: the timeline definition and the cache part. The fixture already owns
        // timeline1/timelineCache1, so the created pair must be new parts rather than additions to
        // the existing ones.
        await Assert.That(entries.Count(n => n.StartsWith("xl/timelines/", StringComparison.Ordinal))).IsEqualTo(2);
        await Assert.That(entries.Count(n => n.StartsWith("xl/timelineCaches/", StringComparison.Ordinal))).IsEqualTo(2);

        // 3: the worksheet's extLst reference, on the sheet the timeline is drawn on.
        await Assert.That(ReadPart(saved, "xl/worksheets/sheet2.xml")).Contains("timelineRef");

        var workbookXml = ReadPart(saved, "xl/workbook.xml");

        // 4: the workbook registration, and 5: the #N/A defined name.
        await Assert.That(workbookXml).Contains("{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}");
        await Assert.That(workbookXml).Contains("NativeTimeline_Date");
        await Assert.That(workbookXml).Contains("#N/A");

        // 6: the drawing anchor.
        await Assert.That(ReadPart(saved, "xl/drawings/drawing2.xml")).Contains("timeslicer");

        // The extras from Task 2's review: assertions confirming the created cache's <state> carries
        // all six pieces the created cache writes, not just the ones already checked above. The
        // created cache part's real name is found from the entry list rather than assumed.
        var createdCachePart = entries
            .Where(n => n.StartsWith("xl/timelineCaches/", StringComparison.Ordinal))
            .OrderBy(n => n, StringComparer.Ordinal)
            .Last();
        var cacheXml = ReadPart(saved, createdCachePart);

        await Assert.That(cacheXml).Contains("minimalRefreshVersion=\"6\"");
        await Assert.That(cacheXml).Contains("lastRefreshVersion=\"6\"");
        await Assert.That(cacheXml).Contains("filterType=\"unknown\"");
        await Assert.That(cacheXml).Contains("startDate=\"1998-01-01");
        await Assert.That(cacheXml).Contains("endDate=\"2005-01-01");
    }

    [Test]
    public async Task A_created_timeline_reloads_with_what_it_was_given()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var pivotTable = wb.Worksheet("Pivot").PivotTables.Single();
            var timeline = wb.Worksheet("Data").Timelines.Add(pivotTable, "Date");
            timeline.Caption = "Pick a period";
            timeline.Style = "TimeSlicerStyleLight2";
            timeline.Level = XLTimelineLevel.Quarters;
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var reloadedTimeline = reloaded.Worksheet("Data").Timelines.Single();

        await Assert.That(reloadedTimeline.Caption).IsEqualTo("Pick a period");
        await Assert.That(reloadedTimeline.Style).IsEqualTo("TimeSlicerStyleLight2");
        await Assert.That(reloadedTimeline.Level).IsEqualTo(XLTimelineLevel.Quarters);
        await Assert.That(reloadedTimeline.SourceFieldName).IsEqualTo("Date");
        await Assert.That(reloadedTimeline.PivotTables.Single().Name).IsEqualTo("СводнаяТаблица2");
    }

    [Test]
    public async Task A_created_timeline_takes_its_bounds_from_the_fields_dates()
    {
        using var wb = Load();

        var pivotTable = wb.Worksheet("Pivot").PivotTables.Single();
        var timeline = wb.Worksheet("Data").Timelines.Add(pivotTable, "Date");

        // The field's dates run 1998-05-19 to 2004-02-06; Excel rounds outward to whole years.
        await Assert.That(timeline.BoundsStart).IsEqualTo(new DateTime(1998, 1, 1));
        await Assert.That(timeline.BoundsEnd).IsEqualTo(new DateTime(2005, 1, 1));
        await Assert.That(timeline.HasSelection).IsFalse();
    }

    [Test]
    public async Task A_timeline_over_a_field_that_holds_no_dates_is_refused()
    {
        using var wb = Load();

        var pivotTable = wb.Worksheet("Pivot").PivotTables.Single();
        var timelines = wb.Worksheet("Data").Timelines;

        // A timeline over a text field is a repair prompt, not a degraded timeline.
        await Assert.That(() => timelines.Add(pivotTable, "Name")).Throws<ArgumentException>();
        await Assert.That(() => timelines.Add(pivotTable, "NoSuchField")).Throws<ArgumentException>();
    }

    // ── The lesson from PRD 5 defect 4 ──────────────────────────────────

    [Test]
    public async Task Adding_a_timeline_beside_an_existing_one_leaves_that_ones_part_untouched()
    {
        // The guard PRD 5's slicer tests were missing. Three byte-equality assertions passed
        // throughout a feature that did not work, because each covered only a sheet where nothing
        // had been added. This adds a timeline to the sheet that already has one.
        using var original = Resource();
        var before = PartBytes(original, "xl/timelines/timeline1.xml");

        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var pivotTable = wb.Worksheet("Pivot").PivotTables.Single();
            wb.Worksheet("Pivot").Timelines.Add(pivotTable, "Date");
            wb.SaveAs(saved);
        }

        await Assert.That(PartBytes(saved, "xl/timelines/timeline1.xml")).IsEquivalentTo(before);
    }

    [Test]
    public async Task A_created_timeline_gets_a_part_of_its_own()
    {
        // Every timelines part Excel writes holds exactly one x15:timeline. Appending into the
        // sheet's existing part instead is what broke slicers: it opens a part Excel authored and
        // hands the SDK the job of serialising it again.
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var pivotTable = wb.Worksheet("Pivot").PivotTables.Single();
            wb.Worksheet("Pivot").Timelines.Add(pivotTable, "Date");
            wb.SaveAs(saved);
        }

        var parts = EntryNames(saved)
            .Where(n => n.StartsWith("xl/timelines/", StringComparison.Ordinal))
            .ToList();

        await Assert.That(parts.Count).IsEqualTo(2);

        foreach (var part in parts)
        {
            var xml = ReadPart(saved, part);
            await Assert.That(CountOccurrences(xml, "<x15:timeline ") + CountOccurrences(xml, "<timeline "))
                .IsEqualTo(1)
                .Because($"{part} must hold exactly one timeline.");
        }
    }

    // ── Editing a loaded timeline ───────────────────────────────────────

    [Test]
    public async Task Editing_a_loaded_timeline_keeps_everything_else_in_its_part()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var timeline = wb.Worksheet("Pivot").Timelines.Single();
            timeline.Caption = "Pick a period";
            timeline.Level = XLTimelineLevel.Quarters;
            wb.SaveAs(saved);
        }

        var xml = ReadPart(saved, "xl/timelines/timeline1.xml");

        await Assert.That(xml).Contains("caption=\"Pick a period\"");
        await Assert.That(xml).Contains("level=\"1\"");

        // Everything XLibur does not model survived, which is the whole point of patching rather
        // than regenerating. selectionLevel and scrollPosition are attributes no XLibur API produces.
        await Assert.That(xml).Contains("selectionLevel=\"2\"");
        await Assert.That(xml).Contains("scrollPosition=\"2004-06-07T00:00:00\"");
        await Assert.That(xml).Contains("mc:Ignorable");
    }

    [Test]
    public async Task A_timeline_nobody_touched_is_not_written_to()
    {
        // Loading a workbook and saving it after an unrelated edit must not open the timeline part.
        using var original = Resource();
        var before = PartBytes(original, "xl/timelines/timeline1.xml");

        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            wb.Worksheet("Data").Cell("Z99").Value = "unrelated";
            wb.SaveAs(saved);
        }

        await Assert.That(PartBytes(saved, "xl/timelines/timeline1.xml")).IsEquivalentTo(before);
    }

    // ── Schema ──────────────────────────────────────────────────────────

    [Test]
    public async Task A_package_with_a_created_timeline_is_schema_valid()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var pivotTable = wb.Worksheet("Pivot").PivotTables.Single();
            wb.Worksheet("Data").Timelines.Add(pivotTable, "Date");
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var doc = SpreadsheetDocument.Open(saved, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2013)
            .Validate(doc)
            .Select(error => $"{error.Path?.XPath}: {error.Description}")
            .ToList();

        await Assert.That(errors).IsEmpty();
    }

    #region Helpers

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

    private static string ReadPart(MemoryStream package, string partPath) =>
        Encoding.UTF8.GetString(PartBytes(package, partPath));

    private static System.Collections.Generic.List<string> EntryNames(MemoryStream package)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        return archive.Entries.Select(e => e.FullName).ToList();
    }

    private static int CountOccurrences(string haystack, string needle)
    {
        var count = 0;
        var index = 0;
        while ((index = haystack.IndexOf(needle, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += needle.Length;
        }

        return count;
    }

    #endregion
}
