using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TUnit.Assertions.Enums;
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
    /// A flag-bearing range that spans the whole sheet is expanded onto every column — which is
    /// what makes it the expensive path, and is unavoidable, since XLibur's model has nowhere but
    /// a column object to record that a column is hidden.
    /// <para>
    /// Asserted as a materialised count rather than as elapsed time: the count is the thing that
    /// actually determines the cost and it is deterministic, whereas a wall-clock bound would be a
    /// flaky test under CI contention. A throughput budget, if one is ever wanted, belongs in the
    /// benchmark project.
    /// </para>
    /// </summary>
    [Test]
    public async Task A_flag_bearing_range_spanning_the_sheet_is_expanded_onto_every_column()
    {
        using var ms = BuildSheetWithColumns(new Column
        {
            Min = 1U,
            Max = (uint)XLHelper.MaxColumnNumber,
            Hidden = true,
        });

        using var wb = new XLWorkbook(ms);
        var ws = wb.Worksheet("S");

        await Assert.That(ws.Column(XLHelper.MaxColumnNumber).IsHidden).IsTrue();
        await Assert.That(((XLWorksheet)ws).Internals.ColumnsCollection.Count)
            .IsEqualTo(XLHelper.MaxColumnNumber);
    }

    /// <summary>
    /// A colour filter is the one criterion whose XML holds an index into <c>&lt;dxfs&gt;</c>, and
    /// <c>AddDifferentialFormats</c> rebuilds that collection on every save — conditional-format
    /// dxfs first, colour-filter dxfs last — so a loaded index is not the index it will be written
    /// at. The criteria of an unchanged column are otherwise written back verbatim, which used to
    /// carry the stale index through with them.
    /// <para>
    /// The fixture is built so the two swap places: dxf 0 is the filter's red fill and dxf 1 the
    /// conditional format's green one, and the rebuild reverses that. Before the fix the filter
    /// kept <c>dxfId="0"</c> and so pointed at the conditional format's green fill — the filter's
    /// colour silently changed on a load and save that touched nothing.
    /// </para>
    /// </summary>
    [Test]
    public async Task A_loaded_colour_filter_is_repointed_at_its_own_dxf_after_the_rebuild()
    {
        var bytes = BuildWorkbookWithColourFilterAndConditionalFormat();

        var (dxfsBefore, filterBefore, ruleBefore) = ReadDxfReferences(bytes);
        await Assert.That(dxfsBefore).IsEquivalentTo(new[] { "FFFF0000", "FF00FF00" },
            CollectionOrdering.Matching);
        await Assert.That(filterBefore).IsEqualTo(0U);
        await Assert.That(ruleBefore).IsEqualTo(1U);

        using var output = new MemoryStream();
        using (var wb = new XLWorkbook(new MemoryStream(bytes, writable: false)))
            wb.SaveAs(output);

        var (dxfsAfter, filterAfter, ruleAfter) = ReadDxfReferences(output.ToArray());

        // The rebuild writes the conditional format's dxf first, so the two have swapped.
        await Assert.That(dxfsAfter).IsEquivalentTo(new[] { "FF00FF00", "FFFF0000" },
            CollectionOrdering.Matching);

        // Both references follow the move: each still points at the colour it started with.
        await Assert.That(filterAfter).IsEqualTo(1U)
            .Because("the colour filter's red fill moved to index 1");
        await Assert.That(ruleAfter).IsEqualTo(0U)
            .Because("the conditional format's green fill moved to index 0");
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

    /// <summary>
    /// A workbook whose <c>&lt;dxfs&gt;</c> are ordered the opposite way from how XLibur rebuilds
    /// them: index 0 is the colour filter's red fill, index 1 the conditional format's green one.
    /// </summary>
    private static byte[] BuildWorkbookWithColourFilterAndConditionalFormat()
    {
        using var ms = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();

            var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet(
                new DocumentFormat.OpenXml.Spreadsheet.Fonts(new Font()) { Count = 1U },
                new DocumentFormat.OpenXml.Spreadsheet.Fills(
                    new Fill(new PatternFill { PatternType = PatternValues.None }),
                    new Fill(new PatternFill { PatternType = PatternValues.Gray125 })) { Count = 2U },
                new DocumentFormat.OpenXml.Spreadsheet.Borders(new Border()) { Count = 1U },
                new CellFormats(new CellFormat()) { Count = 1U },
                new DifferentialFormats(
                    ColourDxf("FFFF0000"),
                    ColourDxf("FF00FF00")) { Count = 2U });
            stylesPart.Stylesheet.Save();

            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new Worksheet(
                new SheetData(
                    new Row(new Cell
                    {
                        CellReference = "A1", DataType = CellValues.String,
                        CellValue = new CellValue("h"),
                    }) { RowIndex = 1U },
                    new Row(new Cell
                    {
                        CellReference = "A2", DataType = CellValues.Number,
                        CellValue = new CellValue("7"),
                    }) { RowIndex = 2U }),
                new AutoFilter(
                    new FilterColumn(new ColorFilter { FormatId = 0U, CellColor = true })
                    { ColumnId = 0U })
                { Reference = "A1:A2" },
                new ConditionalFormatting(
                    new ConditionalFormattingRule(new Formula("5"))
                    {
                        Type = ConditionalFormatValues.CellIs,
                        Operator = ConditionalFormattingOperatorValues.GreaterThan,
                        FormatId = 1U,
                        Priority = 1,
                    })
                {
                    SequenceOfReferences =
                        new ListValue<DocumentFormat.OpenXml.StringValue> { InnerText = "A1:A5" },
                });
            wsPart.Worksheet.Save();

            wbPart.Workbook.AppendChild(new Sheets(
                new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = 1U, Name = "S" }));
            wbPart.Workbook.Save();
        }

        return ms.ToArray();
    }

    private static DifferentialFormat ColourDxf(string argb)
    {
        return new DifferentialFormat(new Fill(new PatternFill
        {
            BackgroundColor = new BackgroundColor { Rgb = argb },
        }));
    }

    /// <summary>
    /// The dxf background colours in order, plus the indices the colour filter and the
    /// conditional-format rule point at.
    /// </summary>
    private static (string[] Dxfs, uint? ColorFilterId, uint? RuleId) ReadDxfReferences(byte[] bytes)
    {
        using var input = new MemoryStream(bytes, writable: false);
        using var doc = SpreadsheetDocument.Open(input, false);

        var dxfs = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.DifferentialFormats;
        var colours = dxfs is null
            ? []
            : dxfs.Elements<DifferentialFormat>()
                .Select(d => d.Fill?.PatternFill?.BackgroundColor?.Rgb?.Value ?? "?")
                .ToArray();

        var ws = doc.WorkbookPart.WorksheetParts.Single().Worksheet;
        return (
            colours,
            ws!.Descendants<ColorFilter>().FirstOrDefault()?.FormatId?.Value,
            ws.Descendants<ConditionalFormattingRule>().FirstOrDefault()?.FormatId?.Value);
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
