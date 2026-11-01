using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Report.Tests.Rewriting;

/// <summary>
/// Pins what XLibur does to charts and pictures when rows are inserted, and what survives a save.
/// </summary>
/// <remarks>
/// Reference rewriting is only worth the code it takes for the cases the core library does not
/// already handle. These tests establish which those are — the spec's risk section calls picture
/// behaviour on row insert unverified, and it turns out charts have a sharper problem than
/// anything the spec anticipated.
/// </remarks>
public class DrawingMechanicsCharacterizationTests
{
    /// <summary>
    /// A one-pixel PNG, written out rather than committed: the tests care where a picture is
    /// anchored, never what it shows, and a literal keeps the fixture readable.
    /// </summary>
    private const string OnePixelPng =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

    private static MemoryStream Image() => new(System.Convert.FromBase64String(OnePixelPng));

    private static IXLWorksheet DataSheet(XLWorkbook workbook)
    {
        var sheet = workbook.AddWorksheet("Data");
        sheet.Cell("A1").Value = "Q1";
        sheet.Cell("A2").Value = "Q2";
        sheet.Cell("B1").Value = 100;
        sheet.Cell("B2").Value = 200;
        return sheet;
    }

    private static MemoryStream Save(XLWorkbook workbook)
    {
        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// A picture anchored to a cell moves down when rows are inserted above it, because its anchor
    /// is held as a live range and every live range is shifted. So the rewriter has nothing to do
    /// for pictures below an expanding range.
    /// </summary>
    [Test]
    public async Task InsertingRowsMovesAPictureAnchoredBelow()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("S");
        sheet.Cell("A1").Value = "top";

        using var image = Image();
        var picture = sheet.AddPicture(image, "Dot").MoveTo(sheet.Cell("C10"));

        sheet.Row(2).InsertRowsBelow(5);

        await Assert.That(picture.TopLeftCell.Address.RowNumber).IsEqualTo(15);
    }

    /// <summary>The shift is a full-row one: a partial-range insert leaves the anchor alone.</summary>
    [Test]
    public async Task APartialRangeInsertDoesNotMoveAPictureOutsideItsColumns()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("S");

        using var image = Image();
        var picture = sheet.AddPicture(image, "Dot").MoveTo(sheet.Cell("F10"));

        sheet.Range("A1:C5").InsertRowsAbove(3);

        await Assert.That(picture.TopLeftCell.Address.RowNumber).IsEqualTo(10);
    }

    /// <summary>A moved anchor is what gets written, so the shift survives the file format.</summary>
    [Test]
    public async Task AMovedPictureAnchorSurvivesASaveAndReload()
    {
        using var stream = new MemoryStream();

        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.AddWorksheet("S");
            using var image = Image();
            sheet.AddPicture(image, "Dot").MoveTo(sheet.Cell("C10"));
            sheet.Row(2).InsertRowsBelow(5);
            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        using var reloaded = new XLWorkbook(stream);
        await Assert.That(reloaded.Worksheet("S").Pictures.Single().TopLeftCell.Address.RowNumber).IsEqualTo(15);
    }

    /// <summary>
    /// A chart's series references are plain strings that nothing shifts: after rows are inserted
    /// into the range a series points at, the series still points where it did. This is the gap the
    /// rewriter exists to close.
    /// </summary>
    [Test]
    public async Task InsertingRowsDoesNotWidenAChartsSeriesReferences()
    {
        using var workbook = new XLWorkbook();
        var sheet = DataSheet(workbook);
        var chart = sheet.Charts.Add(XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");

        sheet.Row(1).InsertRowsBelow(5);

        await Assert.That(chart.Series.Single().ValueReferences).IsEqualTo("Data!$B$1:$B$2");
    }

    /// <summary>
    /// A chart built in memory writes whatever references it holds, so setting them works for a
    /// chart the report itself created.
    /// </summary>
    [Test]
    public async Task ANewChartsEditedReferencesAreWritten()
    {
        using var stream = new MemoryStream();

        using (var workbook = new XLWorkbook())
        {
            var sheet = DataSheet(workbook);
            var chart = sheet.Charts.Add(XLChartType.ColumnClustered);
            chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            chart.Series.Single().ValueReferences = "Data!$B$1:$B$9";
            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        using var reloaded = new XLWorkbook(stream);
        await Assert.That(reloaded.Worksheet("Data").Charts.Single().Series.Single().ValueReferences)
            .IsEqualTo("Data!$B$1:$B$9");
    }

    /// <summary>
    /// The case that matters: a chart loaded from a template. A report template is a file, so every
    /// chart the rewriter touches is a loaded one.
    /// </summary>
    /// <remarks>
    /// This did not work when the rewriter was written. The save path routes loaded charts to
    /// <c>ChartPatcher</c>, which rewrote only formatting, and the reference setters raised no flag
    /// it looked at, so an edited reference was silently dropped — the spec's assumption that
    /// setting a reference "marks the chart edited and the existing patch-on-save path persists it"
    /// was not true of the code. The core now tracks reference assignment the same way it tracks
    /// the formatting properties, and the patcher rewrites <c>c:f</c> and drops the stale cache.
    /// </remarks>
    [Test]
    public async Task ALoadedChartsEditedReferencesAreWritten()
    {
        using var reloaded = ReloadAfterEditing(series =>
        {
            series.ValueReferences = "Data!$B$1:$B$9";
            series.CategoryReferences = "Data!$A$1:$A$9";
        });

        var series = reloaded.Worksheet("Data").Charts.Single().Series.Single();
        await Assert.That(series.ValueReferences).IsEqualTo("Data!$B$1:$B$9");
        await Assert.That(series.CategoryReferences).IsEqualTo("Data!$A$1:$A$9");
    }

    /// <summary>
    /// A loaded chart nobody edited keeps its references, which is the guarantee the patcher's
    /// assignment tracking exists to protect.
    /// </summary>
    [Test]
    public async Task ALoadedChartNobodyEditedKeepsItsReferences()
    {
        using var reloaded = ReloadAfterEditing(_ => { });

        var series = reloaded.Worksheet("Data").Charts.Single().Series.Single();
        await Assert.That(series.ValueReferences).IsEqualTo("Data!$B$1:$B$2");
        await Assert.That(series.CategoryReferences).IsEqualTo("Data!$A$1:$A$2");
    }

    /// <summary>
    /// Writes a chart, loads it back, lets <paramref name="edit"/> change its only series, then
    /// saves and loads again — the round trip a report template makes.
    /// </summary>
    private static XLWorkbook ReloadAfterEditing(System.Action<IXLChartSeries> edit)
    {
        using var original = new MemoryStream();

        using (var workbook = new XLWorkbook())
        {
            var sheet = DataSheet(workbook);
            var chart = sheet.Charts.Add(XLChartType.ColumnClustered);
            chart.Series.Add("Sales", "Data!$B$1:$B$2", "Data!$A$1:$A$2");
            workbook.SaveAs(original);
        }

        original.Position = 0;
        var roundTripped = new MemoryStream();

        using (var loaded = new XLWorkbook(original))
        {
            edit(loaded.Worksheet("Data").Charts.Single().Series.Single());
            loaded.SaveAs(roundTripped);
        }

        roundTripped.Position = 0;
        return new XLWorkbook(roundTripped);
    }
}
