using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Ranges;

/// <summary>
/// Every sheet-scoped feature that must survive a structural edit, in one workbook, edited once.
/// Spec 33 moves the dispatch for all of them behind <c>ISheetListener</c>; this is what proves no
/// feature was dropped on the way.
/// <para>
/// The first four tests build the same sheet twice — once as written, once mirrored across the
/// diagonal — and assert that every feature's address moved by exactly the edit's shift. The
/// mirrored sheet under a column edit is the transposition of the plain sheet under the matching
/// row edit, so one set of expectations covers both axes and all four operations the port has a
/// method for.
/// </para>
/// <para>
/// The four <c>*_does_not_move_yet</c> tests assert the <b>current, wrong</b> behaviour. They exist
/// so that spec 33 tasks 4 and 5 must change them explicitly. Do not "fix" them here.
/// </para>
/// </summary>
public class SheetListenerCharacterizationTests
{
    /// <summary>
    /// Populates one sheet with every feature that reacts to a structural edit today. When
    /// <paramref name="transposed"/> every address is mirrored across the diagonal, so the column
    /// twin exercises the same layout on the other axis.
    /// </summary>
    private static void Populate(XLWorksheet ws, bool transposed)
    {
        IXLRange R(string a1) => ws.Range(Oriented(a1, transposed))!;
        IXLCell C(string a1) => ws.Cell(Oriented(a1, transposed))!;

        R("B10:D10").Merge();                                                      // merged ranges
        R("B12:D14").AddConditionalFormat().WhenGreaterThan(5)
            .Fill.SetBackgroundColor(XLColor.Red);                                 // conditional formats
        R("B16:B17").CreateDataValidation().WholeNumber.Between(1, 10);            // DV sqref
        R("B18:B19").CreateDataValidation().List(R("F20:F22"));                    // DV criteria formula
        ws.Workbook.DefinedNames.Add("Block", R("B22:C24"));                       // defined names
        if (transposed)
            ws.PageSetup.AddVerticalPageBreak(30);                                 // page breaks
        else
            ws.PageSetup.AddHorizontalPageBreak(30);
        C("B26").SetValue("x").SetHyperlink(new XLHyperlink("https://example.invalid/"));
        C("B28").FormulaA1 = "=" + Oriented("B10", transposed) + "+1";             // calc engine
        R("B30:C31").CreateTable();                                                // tables
        R("B33:D34").SetAutoFilter();                                              // autofilter
        ws.SparklineGroups.Add(Oriented("B36", transposed),
                               Oriented("F36:H36", transposed));                   // sparklines
    }

    [Test]
    public async Task A_row_insert_moves_every_sheet_feature()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("S");
        Populate(ws, transposed: false);

        ws.Row(1).InsertRowsAbove(3);

        await AssertEveryFeatureMovedBy(wb, ws, transposed: false, shift: 3);
    }

    [Test]
    public async Task A_column_insert_moves_every_sheet_feature()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("S");
        Populate(ws, transposed: true);

        ws.Column(1).InsertColumnsBefore(3);

        await AssertEveryFeatureMovedBy(wb, ws, transposed: true, shift: 3);
    }

    [Test]
    public async Task A_row_delete_moves_every_sheet_feature()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("S");
        Populate(ws, transposed: false);

        ws.Rows(1, 3).Delete();

        await AssertEveryFeatureMovedBy(wb, ws, transposed: false, shift: -3);
    }

    [Test]
    public async Task A_column_delete_moves_every_sheet_feature()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("S");
        Populate(ws, transposed: true);

        ws.Columns(1, 3).Delete();

        await AssertEveryFeatureMovedBy(wb, ws, transposed: true, shift: -3);
    }

    /// <summary>
    /// Every address <see cref="Populate"/> laid down, moved by <paramref name="shift"/> lines along
    /// the edited axis. Every one of these passed on the unmodified tree — this is a reading, not a
    /// prediction.
    /// </summary>
    private static async Task AssertEveryFeatureMovedBy(XLWorkbook wb, XLWorksheet ws, bool transposed, int shift)
    {
        string Moved(string a1) => Oriented(Shifted(a1, shift), transposed);

        await Assert.That(ws.MergedRanges.Single().RangeAddress.ToString())
            .IsEqualTo(Moved("B10:D10"));
        await Assert.That(ws.ConditionalFormats.Single().Ranges.Single().RangeAddress.ToString())
            .IsEqualTo(Moved("B12:D14"));
        await Assert.That(ws.DataValidations.First().Ranges.Single().RangeAddress.ToString())
            .IsEqualTo(Moved("B16:B17"));
        await Assert.That(ws.DataValidations.Skip(1).First().Ranges.Single().RangeAddress.ToString())
            .IsEqualTo(Moved("B18:B19"));
        await Assert.That(ws.DataValidations.Skip(1).First().MinValue)
            .IsEqualTo("S!" + Absolute(Moved("F20:F22")));
        await Assert.That(wb.DefinedNames.First().RefersTo)
            .IsEqualTo("S!" + Absolute(Moved("B22:C24")));
        await Assert.That(transposed ? ws.PageSetup.ColumnBreaks.Single() : ws.PageSetup.RowBreaks.Single())
            .IsEqualTo(30 + shift);
        await Assert.That(ws.Cell(Moved("B26"))!.HasHyperlink).IsTrue();
        await Assert.That(ws.Cell(Moved("B28"))!.FormulaA1)
            .IsEqualTo(Moved("B10") + "+1");
        await Assert.That(ws.Tables.Cast<IXLTable>().Single().RangeAddress.ToString())
            .IsEqualTo(Moved("B30:C31"));
        await Assert.That(ws.AutoFilter.Range.RangeAddress.ToString())
            .IsEqualTo(Moved("B33:D34"));
        await Assert.That(ws.SparklineGroups.SelectMany(g => g).Single().Location.Address.ToString())
            .IsEqualTo(Moved("B36"));
    }

    /// <summary>
    /// A chart anchor is a pair of raw <c>int</c>s on <c>XLDrawingPosition</c>, and until spec 33
    /// task 4 nothing notified it: a chart anchored at row 10 stayed at row 10 when three rows were
    /// inserted above it. <c>DrawingAnchorListener</c> now moves it.
    /// <para>
    /// This test was <c>A_chart_anchor_does_not_move_yet</c> and asserted 10 and 20 — the wrong
    /// answer, on purpose, so that task 4 had to change it and could not change it silently.
    /// </para>
    /// </summary>
    [Test]
    public async Task A_chart_anchor_moves()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = "Q1";
        ws.Cell("B1").Value = 100;
        var chart = ws.Charts.Add(XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$1", "Data!$A$1:$A$1");
        chart.Position.SetColumn(3).SetRow(10);
        chart.SecondPosition.SetColumn(10).SetRow(20);

        ws.Row(1).InsertRowsAbove(3);

        await Assert.That(chart.Position.Row).IsEqualTo(13);
        await Assert.That(chart.SecondPosition.Row).IsEqualTo(23);

        // The columns are untouched: a row insert moves neither.
        await Assert.That(chart.Position.Column).IsEqualTo(3);
        await Assert.That(chart.SecondPosition.Column).IsEqualTo(10);
    }

    /// <summary>
    /// An absolutely anchored chart is pinned in EMU with no cell reference
    /// (<c>xdr:absoluteAnchor</c>, ECMA-376 §20.5.2.1), so the grid cannot move it.
    /// </summary>
    [Test]
    public async Task An_absolutely_anchored_chart_does_not_move()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = "Q1";
        ws.Cell("B1").Value = 100;
        var chart = ws.Charts.Add(XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$1", "Data!$A$1:$A$1");
        chart.Anchor = XLDrawingAnchor.Absolute;
        chart.Position.SetColumn(3).SetRow(10);

        ws.Row(1).InsertRowsAbove(3);

        await Assert.That(chart.Position.Row).IsEqualTo(10);
    }

    /// <summary>
    /// A chart below the edit, and a chart the edit spans, are the two cases the transform must tell
    /// apart: the first moves whole, the second grows because its first corner is above the insert
    /// and its second is below it.
    /// </summary>
    [Test]
    public async Task An_insert_inside_a_chart_anchor_grows_it()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = "Q1";
        ws.Cell("B1").Value = 100;
        var chart = ws.Charts.Add(XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$1", "Data!$A$1:$A$1");
        chart.Position.SetColumn(2).SetRow(3);          // xdr 0-based: cell C4
        chart.SecondPosition.SetColumn(9).SetRow(19);   // xdr 0-based: cell J20

        ws.Row(10).InsertRowsAbove(3);

        await Assert.That(chart.Position.Row).IsEqualTo(3);
        await Assert.That(chart.SecondPosition.Row).IsEqualTo(22);
    }

    /// <summary>
    /// Deleting every row the anchor sits on clamps it to the deletion point rather than deleting
    /// the chart — the behaviour the picture, the one anchor that already moved, has always had.
    /// </summary>
    [Test]
    public async Task A_delete_covering_a_chart_anchor_clamps_it()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = "Q1";
        ws.Cell("B1").Value = 100;
        var chart = ws.Charts.Add(XLChartType.ColumnClustered);
        chart.Series.Add("Sales", "Data!$B$1:$B$1", "Data!$A$1:$A$1");
        chart.Position.SetColumn(2).SetRow(3);          // xdr 0-based: cell C4
        chart.SecondPosition.SetColumn(9).SetRow(19);   // xdr 0-based: cell J20

        ws.Rows(1, 25).Delete();

        await Assert.That(chart.Position.Row).IsEqualTo(0);        // cell row 1
        await Assert.That(chart.SecondPosition.Row).IsEqualTo(0);  // cell row 1
        await Assert.That(ws.Charts.Count).IsEqualTo(1);
    }

    /// <summary>
    /// The note case used to be half right, and that was the defect. The note itself lives in the
    /// misc slice and moved with the cells; its callout box is an <c>XLDrawingPosition</c> nobody
    /// notified, so the box stayed pinned three rows above where the note had gone —
    /// <c>XLComment</c>'s own remarks documented the mechanism. Both halves now move together.
    /// <para>
    /// This test was <c>A_note_anchor_does_not_move_yet</c> and asserted <c>Position.Row == 9</c> —
    /// the wrong answer, on purpose. Both halves are still asserted, so fixing one and not the other
    /// cannot pass.
    /// </para>
    /// </summary>
    [Test]
    public async Task A_note_anchor_moves()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.Cell("C10").CreateComment().AddText("note");

        ws.Row(1).InsertRowsAbove(3);

        // The note moved with the misc slice. This half was always right.
        await Assert.That(ws.Cell("C10").HasComment).IsFalse();
        await Assert.That(ws.Cell("C13").HasComment).IsTrue();

        // And now its callout box moved with it: one row above the note, where it started.
        await Assert.That(ws.Cell("C13").GetComment().Position.Row).IsEqualTo(12);
        await Assert.That(ws.Cell("C13").GetComment().Position.Column).IsEqualTo(4);
    }

    /// <summary>
    /// The boundary cases, where the edit lands on the note's own cell rather than clear above it.
    /// A callout sits one row above its cell, so it straddles such an edit differently from the cell
    /// and a transform applied to the anchor gets these wrong while getting the common case right —
    /// which is why the anchor takes the <em>cell's</em> displacement instead. Each case asserts the
    /// callout is still exactly one row above the note, wherever the note ended up.
    /// </summary>
    [Test]
    [Arguments("insert clear above the note", 1, 3, 13)]
    [Arguments("insert on the note's own row", 10, 3, 13)]
    [Arguments("insert one row above the note", 9, 3, 13)]
    [Arguments("delete clear above the note", 1, -3, 7)]
    [Arguments("delete ending just above the note", 7, -3, 7)]
    public async Task A_note_keeps_its_callout_one_row_above_it(
        string label, int at, int shift, int expectedNoteRow)
    {
        _ = label;
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.Cell("A10")!.CreateComment().AddText("n");

        if (shift > 0)
            ws.Row(at).InsertRowsAbove(shift);
        else
            ws.Rows(at, at - shift - 1).Delete();

        await Assert.That(ws.Cell(expectedNoteRow, 1)!.HasComment).IsTrue();
        await Assert.That(ws.Cell(expectedNoteRow, 1)!.GetComment().Position.Row)
            .IsEqualTo(expectedNoteRow - 1);
    }

    /// <summary>
    /// A note outside the edited columns does not move, so neither does its callout.
    /// </summary>
    [Test]
    public async Task A_partial_insert_outside_the_notes_column_moves_neither_half()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.Cell("D10")!.CreateComment().AddText("n");

        ws.Range("A1:A5")!.InsertRowsAbove(3);

        await Assert.That(ws.Cell("D10")!.HasComment).IsTrue();
        await Assert.That(ws.Cell("D10")!.GetComment().Position.Row).IsEqualTo(9);
    }

    /// <summary>
    /// A note's callout sits <em>above</em> its cell, so it runs out of grid before the cell does: a
    /// note on <c>A6</c> anchors its box on row 5, and deleting rows 1:5 lands the cell on row 1
    /// while the box would want row 0. Row 0 is not a cell — <c>XLDrawingPosition</c> accepts it, but
    /// the VML writer indexes <c>Worksheet.Row(Position.Row)</c> and throws, so the workbook could
    /// not be saved at all. The box is floored at the first line, which is the same concession
    /// <c>XLComment.Initialize</c> already makes for a note created on row 1.
    /// </summary>
    [Test]
    public async Task A_notes_callout_never_lands_off_the_top_of_the_sheet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.Cell("A6")!.CreateComment().AddText("n");
        await Assert.That(ws.Cell("A6")!.GetComment().Position.Row).IsEqualTo(5);

        ws.Rows(1, 5).Delete();

        await Assert.That(ws.Cell("A1")!.HasComment).IsTrue();
        await Assert.That(ws.Cell("A1")!.GetComment().Position.Row).IsEqualTo(1);

        // The assertion that actually bit: the anchor is only wrong if it reaches the file.
        using var ms = new System.IO.MemoryStream();
        await Assert.That(() => wb.SaveAs(ms, validate: true)).ThrowsNothing();
    }

    /// <summary>
    /// The same floor on the column axis. A callout sits one column to the <em>right</em> of its
    /// cell, so it cannot run off the left edge the way it runs off the top — but a caller can move
    /// it anywhere, and the floor has to hold for that too.
    /// </summary>
    [Test]
    public async Task A_notes_callout_never_lands_off_the_left_of_the_sheet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.Cell("F2")!.CreateComment().AddText("n");
        ws.Cell("F2")!.GetComment().Position.SetColumn(2);

        ws.Columns(1, 5).Delete();

        await Assert.That(ws.Cell("A2")!.HasComment).IsTrue();
        await Assert.That(ws.Cell("A2")!.GetComment().Position.Column).IsEqualTo(1);

        using var ms = new System.IO.MemoryStream();
        await Assert.That(() => wb.SaveAs(ms, validate: true)).ThrowsNothing();
    }

    /// <summary>
    /// A pivot table's report filters sit above its area, and <c>TargetCell</c> is the area's first
    /// point shifted up by that many rows. A delete clamps the area onto the deletion point, so the
    /// area has to stop far enough down the sheet for the filters to still fit — otherwise the
    /// target lands above row 1, and <c>Point</c> stores its row in an unsigned field, so it wraps
    /// rather than throwing and hands the caller <c>#REF!</c>.
    /// </summary>
    [Test]
    public async Task A_pivot_area_leaves_room_for_its_report_filters()
    {
        using var wb = new XLWorkbook();
        var source = wb.AddWorksheet("Source");
        source.Cell("A1").Value = "Name";
        source.Cell("A2").Value = "Alice";
        source.Cell("B1").Value = "Amount";
        source.Cell("B2").Value = 1;

        var ws = (XLWorksheet)wb.AddWorksheet("Pivot");
        var pt = (XLPivotTable)ws.PivotTables.Add("pt", ws.Cell("D3")!, source.Range("A1:B2")!);
        pt.ReportFilters.Add("Name");
        pt.Values.Add("Amount");

        // One report filter plus the gap puts the area two rows below the target.
        await Assert.That(pt.TargetCell.Address.ToString()).IsEqualTo("D3");

        ws.Rows(1, 4).Delete();

        await Assert.That(pt.Area.FirstPoint.Row).IsEqualTo(3);
        await Assert.That(pt.TargetCell.Address.ToString()).IsEqualTo("D1");
    }

    /// <summary>
    /// The scrollable pane's anchor cell moves with the split it belongs to. <c>SheetViewWriter</c>
    /// writes it verbatim as <c>pane/@topLeftCell</c> when it is set, so leaving it behind would
    /// write an anchor sitting inside the frozen band.
    /// </summary>
    [Test]
    public async Task The_pane_anchor_moves_with_its_split()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.SheetView.FreezeRows(5);
        ws.SheetView.PaneTopLeftCellAddress = ws.Cell("A6")!.Address;

        ws.Row(5).InsertRowsAbove(3);

        await Assert.That(ws.SheetView.SplitRow).IsEqualTo(8);
        await Assert.That(ws.SheetView.PaneTopLeftCellAddress!.ToString()).IsEqualTo("A9");
    }

    /// <summary>
    /// <c>SplitRow</c> and <c>SplitColumn</c> are raw <c>int</c>s on <c>XLSheetView</c> and until
    /// spec 33 task 5 nothing notified them, so inserting rows inside the frozen region left the
    /// freeze splitting the wrong line.
    /// <para>
    /// This test was <c>Freeze_panes_do_not_move_yet</c> and asserted 5 and 4 — the wrong answer, on
    /// purpose.
    /// </para>
    /// </summary>
    [Test]
    public async Task Freeze_panes_move()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.SheetView.Freeze(5, 4);

        ws.Row(1).InsertRowsAbove(3);
        ws.Column(1).InsertColumnsBefore(2);

        await Assert.That(ws.SheetView.SplitRow).IsEqualTo(8);
        await Assert.That(ws.SheetView.SplitColumn).IsEqualTo(6);
    }

    /// <summary>
    /// The three cases the count transform has to tell apart, plus the one where the pane goes away:
    /// an edit below the split leaves it alone, an edit inside it moves it, and deleting every line
    /// above it leaves a count of zero — which <c>SheetViewWriter</c> writes as no pane at all.
    /// </summary>
    [Test]
    [Arguments("insert below the split", 10, 3, 5)]
    [Arguments("insert inside the split", 1, 3, 8)]
    [Arguments("insert on the split line", 5, 3, 8)]
    [Arguments("delete inside the split", 2, -2, 3)]
    [Arguments("delete everything above the split", 1, -5, 0)]
    [Arguments("delete more than the split", 1, -10, 0)]
    public async Task A_row_edit_moves_the_freeze_only_when_it_is_inside_it(
        string label, int at, int shift, int expectedSplitRow)
    {
        _ = label;
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.SheetView.FreezeRows(5);

        if (shift > 0)
            ws.Row(at).InsertRowsAbove(shift);
        else
            ws.Rows(at, at - shift - 1).Delete();

        await Assert.That(ws.SheetView.SplitRow).IsEqualTo(expectedSplitRow);
    }

    /// <summary>
    /// Inserting cells into part of a row and shifting them down is not a row insert, and does not
    /// move a freeze: the frozen band spans every column, so only an edit that spans every column
    /// can grow it.
    /// </summary>
    [Test]
    public async Task A_partial_insert_does_not_move_the_freeze()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.SheetView.Freeze(5, 4);

        ws.Range("B1:B5")!.InsertRowsAbove(3);

        await Assert.That(ws.SheetView.SplitRow).IsEqualTo(5);
        await Assert.That(ws.SheetView.SplitColumn).IsEqualTo(4);
    }

    /// <summary>
    /// <c>XLPivotTable.Area</c> is a raw <c>Area</c> and until spec 33 task 5 nothing notified it,
    /// so a pivot table anchored at <c>D10</c> stayed at <c>D10</c> while the cells Excel had
    /// rendered it into moved.
    /// <para>
    /// This test was <c>A_pivot_area_does_not_move_yet</c> and asserted row 10 — the wrong answer,
    /// on purpose.
    /// </para>
    /// </summary>
    [Test]
    public async Task A_pivot_area_moves()
    {
        using var wb = new XLWorkbook();
        var source = wb.AddWorksheet("Source");
        source.Cell("A1").Value = "Name";
        source.Cell("A2").Value = "Alice";
        source.Cell("B1").Value = "Amount";
        source.Cell("B2").Value = 1;

        var ws = (XLWorksheet)wb.AddWorksheet("Pivot");
        var pt = (XLPivotTable)ws.PivotTables.Add("pt", ws.Cell("D10")!, source.Range("A1:B2")!);
        pt.RowLabels.Add("Name");
        pt.Values.Add("Amount");

        ws.Row(1).InsertRowsAbove(3);

        await Assert.That(pt.Area.FirstPoint.Row).IsEqualTo(13);
        await Assert.That(pt.Area.FirstPoint.Column).IsEqualTo(4);

        // TargetCell derives from Area, so moving the area moves the table.
        await Assert.That(pt.TargetCell.Address.ToString()).IsEqualTo("D13");
    }

    /// <summary>The address moved by <paramref name="shift"/> rows, before any transposition.</summary>
    private static string Shifted(string a1, int shift)
        => Map(a1, (letters, row) => letters + (row + shift));

    /// <summary>The address as the sheet holds it: mirrored across the diagonal when transposed.</summary>
    private static string Oriented(string a1, bool transposed)
        => transposed
            ? Map(a1, (letters, row) => XLHelper.GetColumnLetterFromNumber(row) +
                                        XLHelper.GetColumnNumberFromLetter(letters))
            : a1;

    /// <summary>The address with every row and column fixed, as a defined name stores it.</summary>
    private static string Absolute(string a1)
        => Map(a1, (letters, row) => "$" + letters + "$" + row);

    /// <summary>Applies <paramref name="rewrite"/> to each A1 reference of an address or range address.</summary>
    private static string Map(string a1, System.Func<string, int, string> rewrite)
        => string.Join(":", a1.Split(':').Select(part =>
        {
            var letters = new string(part.TakeWhile(char.IsLetter).ToArray());
            var row = int.Parse(new string(part.SkipWhile(char.IsLetter).ToArray()));
            return rewrite(letters, row);
        }));
}
