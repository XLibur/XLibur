using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// Two defects found while reviewing spec 28, fixed together because both are load/save
/// bookkeeping around styles rather than decoding: D14, a <c>&lt;col&gt;</c> range ending at the
/// last column losing its per-column flags, and D12, a differential-format reuse map built by
/// iterating a collection emptied on the previous line.
/// </summary>
internal class ColumnRangeAndDxfMapTests
{
    /// <summary>
    /// D14. Excel writes <c>&lt;col min="2" max="16384" hidden="1"/&gt;</c> when a user hides from
    /// column B rightwards. The reader used to treat every range ending at the last column as the
    /// sheet default and skip it, so <c>hidden</c>, <c>collapsed</c> and <c>outlineLevel</c> were
    /// silently dropped while the width survived.
    /// </summary>
    [Test]
    public async Task A_column_range_ending_at_the_last_column_keeps_its_per_column_flags()
    {
        using var ms = BuildSheetWithColumns(new Column
        {
            Min = 2U,
            Max = (uint)XLHelper.MaxColumnNumber,
            Hidden = true,
            OutlineLevel = 1,
            Width = 20D,
            CustomWidth = true,
        });

        using var wb = new XLWorkbook(ms);
        var ws = wb.Worksheet("S");

        await Assert.That(ws.Column(2).IsHidden).IsTrue();
        await Assert.That(ws.Column(2).OutlineLevel).IsEqualTo(1);

        // A column inside the range, not just its first: the range really was expanded.
        await Assert.That(ws.Column(500).IsHidden).IsTrue();
        await Assert.That(ws.Column(500).OutlineLevel).IsEqualTo(1);

        // Column A is outside the range and keeps the defaults.
        await Assert.That(ws.Column(1).IsHidden).IsFalse();
        await Assert.That(ws.Column(1).OutlineLevel).IsEqualTo(0);

        // The width the range states still becomes the sheet default, as before.
        await Assert.That(ws.ColumnWidth).IsEqualTo(20D - XLConstants.ColumnWidthOffset);
    }

    /// <summary>
    /// The other half of D14's fix: a range ending at the last column that states only width and
    /// style is still treated as the sheet default and is <b>not</b> expanded. That matters for
    /// more than tidiness — expanding it would materialise an <c>XLColumn</c> for all 16,384
    /// columns on every load of an ordinary file, which is why the skip existed. This is the shape
    /// <c>ColumnWriter.WritePostColumns</c> emits, so it is the common case.
    /// </summary>
    [Test]
    public async Task A_plain_default_range_is_still_treated_as_the_sheet_default()
    {
        using var ms = BuildSheetWithColumns(new Column
        {
            Min = 2U,
            Max = (uint)XLHelper.MaxColumnNumber,
            Width = 20D,
            CustomWidth = true,
        });

        using var wb = new XLWorkbook(ms);
        var ws = wb.Worksheet("S");

        await Assert.That(ws.ColumnWidth).IsEqualTo(20D - XLConstants.ColumnWidthOffset);

        // The range was not expanded: nothing is materialised before anything asks for a column.
        // Checked before the reads below, which materialise one column each by themselves.
        await Assert.That(((XLWorksheet)ws).Internals.ColumnsCollection.Count).IsEqualTo(0);

        // And nothing was hidden or grouped.
        await Assert.That(ws.Column(500).IsHidden).IsFalse();
    }

    /// <summary>
    /// The expensive path is bounded. Expanding a last-column range is unavoidable when it carries
    /// per-column flags — XLibur's model has nowhere else to put them — so this pins that the cost
    /// stays in the "noticeable but fine" range rather than becoming a load-time cliff. Generous
    /// bound on purpose: it is a smoke alarm, not a benchmark.
    /// </summary>
    [Test]
    public async Task Expanding_a_flag_bearing_last_column_range_stays_bounded()
    {
        using var ms = BuildSheetWithColumns(new Column
        {
            Min = 1U,
            Max = (uint)XLHelper.MaxColumnNumber,
            Hidden = true,
        });

        var sw = Stopwatch.StartNew();
        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheet("S");
            await Assert.That(ws.Column(16384).IsHidden).IsTrue();
        }

        sw.Stop();
        await Assert.That(sw.Elapsed.TotalSeconds).IsLessThan(10D)
            .Because($"expanding all {XLHelper.MaxColumnNumber} columns took {sw.Elapsed}");
    }

    /// <summary>
    /// D12. <c>AddDifferentialFormats</c> rebuilds <c>&lt;dxfs&gt;</c> from the live workbook, so
    /// a dxf nothing references any longer must not survive the save, and the count must not drift
    /// across repeated round trips. Before the fix a dead helper claimed to build a reuse map from
    /// the dxfs already in the file while iterating a collection cleared on the line above; this
    /// pins the behaviour that was actually running, so removing the helper is visibly a no-op.
    /// </summary>
    [Test]
    public async Task Removing_a_conditional_format_drops_its_dxf_and_the_count_stays_stable()
    {
        using var first = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Range("A1:A5").AddConditionalFormat().WhenGreaterThan(5).Fill
                .SetBackgroundColor(XLColor.Red);
            ws.Range("B1:B5").AddConditionalFormat().WhenLessThan(2).Fill
                .SetBackgroundColor(XLColor.Blue);
            wb.SaveAs(first);
        }

        await Assert.That(CountDxfs(first.ToArray())).IsEqualTo(2);

        // Drop one conditional format; its dxf must go with it.
        using var second = new MemoryStream();
        using (var wb = new XLWorkbook(new MemoryStream(first.ToArray(), writable: false)))
        {
            var ws = wb.Worksheet("Sheet1");
            ws.ConditionalFormats.Remove(cf => cf.Range.RangeAddress.ToStringRelative() == "B1:B5");
            wb.SaveAs(second);
        }

        await Assert.That(CountDxfs(second.ToArray())).IsEqualTo(1);

        // And a save that changes nothing leaves the count alone.
        var bytes = second.ToArray();
        for (var i = 0; i < 3; i++)
        {
            using var input = new MemoryStream(bytes, writable: false);
            using var wb = new XLWorkbook(input);
            using var output = new MemoryStream();
            wb.SaveAs(output);
            bytes = output.ToArray();
            await Assert.That(CountDxfs(bytes)).IsEqualTo(1);
        }
    }

    private static MemoryStream BuildSheetWithColumns(params Column[] cols)
    {
        var ms = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new Worksheet(
                new DocumentFormat.OpenXml.Spreadsheet.Columns(cols), new SheetData());
            wbPart.Workbook.AppendChild(new Sheets(
                new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = 1U, Name = "S" }));
            wbPart.Workbook.Save();
        }

        ms.Position = 0;
        return ms;
    }

    private static int CountDxfs(byte[] bytes)
    {
        using var input = new MemoryStream(bytes, writable: false);
        using var doc = SpreadsheetDocument.Open(input, false);
        var dxfs = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.DifferentialFormats;
        return dxfs?.ChildElements.Count ?? 0;
    }
}
