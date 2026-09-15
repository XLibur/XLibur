using System;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Styles;

/// <summary>
/// A style written through a range must reach its cells whether or not the garbage collector ran
/// between two statements (#505).
/// </summary>
/// <remarks>
/// The worksheet keeps its ranges in a weak repository, so a range the caller does not hold can be
/// collected between two statements and rebuilt by the next. A rebuilt range starts from its parent's
/// style, or the worksheet's, rather than from its cells', so a setter that compared against the
/// range's own record could skip a write, or hold back a border colour, that the cells needed.
/// <para>
/// Each style statement builds its range inside a non-inlined method, so no stack slot of the test
/// method keeps the range alive, and <see cref="Collect"/> between them forces what an ordinary run
/// leaves to chance.
/// </para>
/// </remarks>
public class RangeStyleAfterCollectionTests
{
    private static readonly XLColor ThemeColor = XLColor.FromTheme(XLThemeColor.Accent1, 0.5);

    private static void Collect()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }

    /// <summary>
    /// The statements of <c>BorderTests.SetInsideBorderPreservesOutsideBorders</c>, which failed at
    /// random in full runs when a collection fell between them.
    /// </summary>
    [Test]
    public async Task Outside_border_colour_on_cells_is_kept_when_a_collection_runs_after_the_style()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        OnCells(ws, "B2:C2", s => s.Border.SetOutsideBorder(XLBorderStyleValues.Thin));
        Collect();
        OnCells(ws, "B2:C2", s => s.Border.SetOutsideBorderColor(ThemeColor));

        foreach (var address in new[] { "B2", "C2" })
        {
            var border = ws.Cell(address).Style.Border;
            await Assert.That(border.LeftBorderColor).IsEqualTo(ThemeColor);
            await Assert.That(border.RightBorderColor).IsEqualTo(ThemeColor);
            await Assert.That(border.TopBorderColor).IsEqualTo(ThemeColor);
            await Assert.That(border.BottomBorderColor).IsEqualTo(ThemeColor);
        }
    }

    [Test]
    public async Task Outside_border_colour_on_a_range_is_kept_when_a_collection_runs_after_the_style()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        OnRange(ws, "B2:D4", s => s.Border.SetOutsideBorder(XLBorderStyleValues.Thin));
        Collect();
        OnRange(ws, "B2:D4", s => s.Border.SetOutsideBorderColor(XLColor.Red));

        await Assert.That(ws.Cell("B2").Style.Border.LeftBorderColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Cell("B2").Style.Border.TopBorderColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Cell("D4").Style.Border.RightBorderColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Cell("D4").Style.Border.BottomBorderColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Cell("C3").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.None);
    }

    [Test]
    public async Task Edge_colour_on_a_range_is_kept_when_a_collection_runs_after_the_edge_style()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        OnRange(ws, "A1:B2", s => s.Border.LeftBorder = XLBorderStyleValues.Thin);
        Collect();
        OnRange(ws, "A1:B2", s => s.Border.LeftBorderColor = XLColor.Red);

        await Assert.That(ws.Cell("A1").Style.Border.LeftBorderColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Cell("B2").Style.Border.LeftBorderColor).IsEqualTo(XLColor.Red);
    }

    [Test]
    public async Task Removing_an_outside_border_from_a_range_works_when_a_collection_runs_first()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        OnRange(ws, "B2:C3", s => s.Border.SetOutsideBorder(XLBorderStyleValues.Thin));
        Collect();
        OnRange(ws, "B2:C3", s => s.Border.SetOutsideBorder(XLBorderStyleValues.None));

        await Assert.That(ws.Cell("B2").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("B2").Style.Border.TopBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("C3").Style.Border.RightBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("C3").Style.Border.BottomBorder).IsEqualTo(XLBorderStyleValues.None);
    }

    [Test]
    public async Task Clearing_diagonal_up_on_a_range_works_when_a_collection_runs_first()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        OnRange(ws, "A1:B2", s => s.Border.DiagonalUp = true);
        Collect();
        OnRange(ws, "A1:B2", s => s.Border.DiagonalUp = false);

        await Assert.That(ws.Cell("A1").Style.Border.DiagonalUp).IsFalse();
        await Assert.That(ws.Cell("B2").Style.Border.DiagonalUp).IsFalse();
    }

    [Test]
    public async Task Clearing_bold_on_a_range_works_when_a_collection_runs_first()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        OnRange(ws, "A1:B2", s => s.Font.Bold = true);
        Collect();
        OnRange(ws, "A1:B2", s => s.Font.Bold = false);

        await Assert.That(ws.Cell("A1").Style.Font.Bold).IsFalse();
        await Assert.That(ws.Cell("B2").Style.Font.Bold).IsFalse();
    }

    [Test]
    public async Task Clearing_wrap_text_on_a_range_works_when_a_collection_runs_first()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        OnRange(ws, "A1:B2", s => s.Alignment.WrapText = true);
        Collect();
        OnRange(ws, "A1:B2", s => s.Alignment.WrapText = false);

        await Assert.That(ws.Cell("A1").Style.Alignment.WrapText).IsFalse();
        await Assert.That(ws.Cell("B2").Style.Alignment.WrapText).IsFalse();
    }

    [Test]
    public async Task Locking_a_range_again_works_when_a_collection_runs_first()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        OnRange(ws, "A1:B2", s => s.Protection.Locked = false);
        Collect();
        OnRange(ws, "A1:B2", s => s.Protection.Locked = true);

        await Assert.That(ws.Cell("A1").Style.Protection.Locked).IsTrue();
        await Assert.That(ws.Cell("B2").Style.Protection.Locked).IsTrue();
    }

    [Test]
    public async Task Clearing_a_fill_pattern_on_a_range_works_when_a_collection_runs_first()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        OnRange(ws, "A1:B2", s => s.Fill.BackgroundColor = XLColor.Red);
        Collect();
        OnRange(ws, "A1:B2", s => s.Fill.PatternType = XLFillPatternValues.None);

        await Assert.That(ws.Cell("A1").Style.Fill.PatternType).IsEqualTo(XLFillPatternValues.None);
        await Assert.That(ws.Cell("B2").Style.Fill.PatternType).IsEqualTo(XLFillPatternValues.None);
    }

    /// <summary>
    /// Setting a background colour gives an unpatterned fill the solid pattern. Whether a fill is
    /// unpatterned is a question about each cell, not about the range's record of itself.
    /// </summary>
    [Test]
    public async Task Background_colour_on_a_range_keeps_the_cells_pattern_when_a_collection_runs_first()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        OnRange(ws, "A1:B2", s => s.Fill.PatternType = XLFillPatternValues.DarkGray);
        Collect();
        OnRange(ws, "A1:B2", s => s.Fill.BackgroundColor = XLColor.Red);

        await Assert.That(ws.Cell("A1").Style.Fill.PatternType).IsEqualTo(XLFillPatternValues.DarkGray);
        await Assert.That(ws.Cell("A1").Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Cell("B2").Style.Fill.PatternType).IsEqualTo(XLFillPatternValues.DarkGray);
    }

    /// <summary>
    /// A colour given to an unstyled edge before its style, through one range facade, is still held
    /// and applied when the style arrives. The change for #505 writes colours through on ranges and
    /// must not lose this.
    /// </summary>
    [Test]
    public async Task Edge_colour_set_before_the_style_on_one_range_facade_is_applied_with_the_style()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var border = ws.Range("A1:B2").Style.Border;
        border.TopBorderColor = XLColor.Red;
        border.TopBorder = XLBorderStyleValues.Thin;

        await Assert.That(ws.Cell("A1").Style.Border.TopBorder).IsEqualTo(XLBorderStyleValues.Thin);
        await Assert.That(ws.Cell("A1").Style.Border.TopBorderColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Cell("B2").Style.Border.TopBorderColor).IsEqualTo(XLColor.Red);
    }

    /// <summary>
    /// A range held by the caller keeps its own record of its style, and a cell edited directly
    /// afterwards leaves that record behind. Setting the value again through the range must still
    /// reach that cell.
    /// </summary>
    [Test]
    public async Task Setting_a_range_value_again_reaches_a_cell_edited_directly_in_between()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range("A1:B2");

        range.Style.Font.Bold = true;
        range.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
        ws.Cell("A1").Style.Font.Bold = false;
        ws.Cell("A1").Style.Border.LeftBorder = XLBorderStyleValues.None;

        range.Style.Font.Bold = true;
        range.Style.Border.LeftBorder = XLBorderStyleValues.Thin;

        await Assert.That(ws.Cell("A1").Style.Font.Bold).IsTrue();
        await Assert.That(ws.Cell("A1").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thin);
    }

    /// <summary>
    /// A colour set before the edge's style, through one range facade, reaches every cell with
    /// the style - including a cell whose edge had no style when the colour was set, although the
    /// range's own record said the edge was styled. On a cell the same two statements give red.
    /// </summary>
    [Test]
    public async Task Edge_colour_then_style_through_a_kept_range_reaches_a_cell_whose_edge_had_no_style()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range("A1:B2");
        range.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
        ws.Cell("A1").Style.Border.LeftBorder = XLBorderStyleValues.None;

        var border = range.Style.Border;
        border.LeftBorderColor = XLColor.Red;
        border.LeftBorder = XLBorderStyleValues.Thin;

        await Assert.That(ws.Cell("A1").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thin);
        await Assert.That(ws.Cell("A1").Style.Border.LeftBorderColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Cell("B2").Style.Border.LeftBorderColor).IsEqualTo(XLColor.Red);
    }

    /// <summary>
    /// The same, through a range built fresh from a parent whose edge is styled: the new range's
    /// record comes from the parent, not from its cells.
    /// </summary>
    [Test]
    public async Task Edge_colour_then_style_through_a_range_built_from_a_styled_parent_reaches_every_cell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var parent = ws.Range("A1:B2");
        parent.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
        ws.Cell("A1").Style.Border.LeftBorder = XLBorderStyleValues.None;

        var border = parent.Range("A1:A2").Style.Border;
        border.LeftBorderColor = XLColor.Red;
        border.LeftBorder = XLBorderStyleValues.Thin;

        await Assert.That(ws.Cell("A1").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thin);
        await Assert.That(ws.Cell("A1").Style.Border.LeftBorderColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Cell("A2").Style.Border.LeftBorderColor).IsEqualTo(XLColor.Red);
    }

    /// <summary>
    /// An indent needs left, right or distributed alignment. Through a range the question is
    /// asked of each cell: a cell centred directly since the range last set the indent is
    /// left-aligned again rather than left centred with an indent.
    /// </summary>
    [Test]
    public async Task Indent_through_a_kept_range_left_aligns_a_cell_centred_directly_in_between()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range("A1:B2");
        range.Style.Alignment.Indent = 2;
        ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        range.Style.Alignment.Indent = 2;

        await Assert.That(ws.Cell("A1").Style.Alignment.Horizontal).IsEqualTo(XLAlignmentHorizontalValues.Left);
        await Assert.That(ws.Cell("A1").Style.Alignment.Indent).IsEqualTo(2);
        await Assert.That(ws.Cell("B2").Style.Alignment.Horizontal).IsEqualTo(XLAlignmentHorizontalValues.Left);
        await Assert.That(ws.Cell("B2").Style.Alignment.Indent).IsEqualTo(2);
    }

    /// <summary>
    /// Through a range an indent left-aligns each cell that cannot take one, instead of throwing
    /// because the range's own record is centred. Nothing is validated against the record, so the
    /// write cannot stop part-way.
    /// </summary>
    [Test]
    public async Task Indent_through_a_centred_range_left_aligns_its_cells_instead_of_throwing()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range("A1:B2");
        range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        range.Style.Alignment.Indent = 2;

        await Assert.That(ws.Cell("A1").Style.Alignment.Horizontal).IsEqualTo(XLAlignmentHorizontalValues.Left);
        await Assert.That(ws.Cell("A1").Style.Alignment.Indent).IsEqualTo(2);
        await Assert.That(ws.Cell("B2").Style.Alignment.Horizontal).IsEqualTo(XLAlignmentHorizontalValues.Left);
    }

    /// <summary>A single cell keeps its behaviour: a centred cell refuses an indent, unchanged.</summary>
    [Test]
    public async Task Indent_on_a_centred_cell_still_throws_and_changes_nothing()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cell = ws.Cell("A1");
        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        await Assert.That(() => { cell.Style.Alignment.Indent = 2; }).Throws<ArgumentException>();

        await Assert.That(cell.Style.Alignment.Horizontal).IsEqualTo(XLAlignmentHorizontalValues.Center);
        await Assert.That(cell.Style.Alignment.Indent).IsEqualTo(0);
    }

    /// <summary>
    /// A worksheet still skips a value its own style already holds: no collection can replace a
    /// worksheet's style, so the skip is kept there for what it saves. The skipped set must leave
    /// the cells as the first set did, and a different value must still reach them.
    /// </summary>
    [Test]
    public async Task Setting_a_worksheet_font_value_again_is_still_correct_when_a_collection_runs_first()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("B2").Value = 2;

        OnWorksheet(ws, s => s.Font.Bold = true);
        Collect();
        OnWorksheet(ws, s => s.Font.Bold = true);

        await Assert.That(ws.Cell("A1").Style.Font.Bold).IsTrue();
        await Assert.That(ws.Cell("B2").Style.Font.Bold).IsTrue();

        Collect();
        OnWorksheet(ws, s => s.Font.Bold = false);

        await Assert.That(ws.Cell("A1").Style.Font.Bold).IsFalse();
        await Assert.That(ws.Cell("B2").Style.Font.Bold).IsFalse();
    }

    [Test]
    public async Task Setting_a_worksheet_edge_style_and_colour_again_is_still_correct_when_a_collection_runs_first()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;

        OnWorksheet(ws, s => s.Border.LeftBorder = XLBorderStyleValues.Thin);
        Collect();
        OnWorksheet(ws, s => s.Border.LeftBorder = XLBorderStyleValues.Thin);
        OnWorksheet(ws, s => s.Border.LeftBorderColor = XLColor.Red);
        Collect();
        OnWorksheet(ws, s => s.Border.LeftBorderColor = XLColor.Red);

        await Assert.That(ws.Cell("A1").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thin);
        await Assert.That(ws.Cell("A1").Style.Border.LeftBorderColor).IsEqualTo(XLColor.Red);

        Collect();
        OnWorksheet(ws, s => s.Border.LeftBorder = XLBorderStyleValues.None);

        await Assert.That(ws.Cell("A1").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.None);
    }

    /// <summary>
    /// A colour every cell took at once is not held for the next style as well: it would paint
    /// over a colour a cell was given directly in between, which the cell's own edit should keep.
    /// </summary>
    [Test]
    public async Task Edge_colour_every_cell_took_does_not_repaint_a_cell_coloured_directly_in_between()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range("A1:B2");
        range.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
        range.Style.Border.LeftBorderColor = XLColor.Red;
        ws.Cell("A1").Style.Border.LeftBorderColor = XLColor.Blue;

        range.Style.Border.LeftBorder = XLBorderStyleValues.Thick;

        await Assert.That(ws.Cell("A1").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(ws.Cell("A1").Style.Border.LeftBorderColor).IsEqualTo(XLColor.Blue);
        await Assert.That(ws.Cell("B2").Style.Border.LeftBorderColor).IsEqualTo(XLColor.Red);
    }

    /// <summary>
    /// The same through <see cref="IXLBorder.OutsideBorderColor"/>, which writes through the
    /// range's edge sub-ranges and so never touches the range's own facade: a colour the range
    /// set earlier, and every cell took, must not come back with the next style.
    /// </summary>
    [Test]
    public async Task Edge_colour_every_cell_took_does_not_come_back_after_an_outside_border_colour()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range("A1:B2");
        range.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
        range.Style.Border.LeftBorderColor = XLColor.Red;
        range.Style.Border.OutsideBorderColor = XLColor.Blue;

        range.Style.Border.LeftBorder = XLBorderStyleValues.Thin;

        await Assert.That(ws.Cell("A1").Style.Border.LeftBorderColor).IsEqualTo(XLColor.Blue);
        await Assert.That(ws.Cell("A2").Style.Border.LeftBorderColor).IsEqualTo(XLColor.Blue);
        await Assert.That(ws.Cell("B1").Style.Border.LeftBorderColor).IsEqualTo(XLColor.Red);
    }

    [Test]
    public async Task Worksheet_edge_colour_every_cell_took_does_not_repaint_a_cell_coloured_directly_in_between()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("B2").Value = 2;
        ws.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
        ws.Style.Border.LeftBorderColor = XLColor.Red;
        ws.Cell("A1").Style.Border.LeftBorderColor = XLColor.Blue;

        ws.Style.Border.LeftBorder = XLBorderStyleValues.Thick;

        await Assert.That(ws.Cell("A1").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(ws.Cell("A1").Style.Border.LeftBorderColor).IsEqualTo(XLColor.Blue);
        await Assert.That(ws.Cell("B2").Style.Border.LeftBorderColor).IsEqualTo(XLColor.Red);
    }

    /// <summary>
    /// A known, accepted limit. When some cell could not take the colour - its edge had no style -
    /// the range holds the colour for the next style, and that style then paints it on every cell,
    /// including one given its own colour directly in between. Telling the two apart would need a
    /// record of each cell's colour, which the range does not keep (#505 review).
    /// </summary>
    [Test]
    public async Task Pending_colour_still_wins_over_a_cell_coloured_directly_when_another_cell_dropped_it()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range("A1:B2");
        range.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
        ws.Cell("A1").Style.Border.LeftBorder = XLBorderStyleValues.None;

        var border = range.Style.Border;
        border.LeftBorderColor = XLColor.Red;
        ws.Cell("B2").Style.Border.LeftBorderColor = XLColor.Blue;
        border.LeftBorder = XLBorderStyleValues.Thin;

        await Assert.That(ws.Cell("A1").Style.Border.LeftBorderColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Cell("B2").Style.Border.LeftBorderColor).IsEqualTo(XLColor.Red);
    }

    /// <summary>
    /// A colour set on an edge with no border is remembered only by the style object it was set
    /// through. Holding one range across both statements, a cell whose edge had no border takes the
    /// colour with the style. Two separate <c>ws.Range(...)</c> calls usually return that same
    /// cached object, but not after a garbage collection, and then the cell takes the style alone,
    /// in black - which is why the colour and the style should go through one object.
    /// </summary>
    [Test]
    public async Task Edge_colour_then_style_through_one_held_range_reaches_a_cell_whose_edge_had_no_border()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("A1:C3").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
        ws.Cell("A1").Style.Border.LeftBorder = XLBorderStyleValues.None;

        var range = ws.Range("A1:B2");
        range.Style.Border.LeftBorderColor = XLColor.Red;
        range.Style.Border.LeftBorder = XLBorderStyleValues.Thin;

        await Assert.That(ws.Cell("A1").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thin);
        await Assert.That(ws.Cell("A1").Style.Border.LeftBorderColor).IsEqualTo(XLColor.Red);
    }

    /// <summary>
    /// A worksheet skips an indent its own style already holds, as it skips every other value it
    /// holds, so a cell given its own indent directly keeps it - and the sheet is not walked.
    /// </summary>
    [Test]
    public async Task Setting_a_worksheet_indent_it_already_holds_leaves_a_cell_indented_directly_alone()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("B2").Value = 2;
        ws.Style.Alignment.Indent = 2;
        ws.Cell("A1").Style.Alignment.Indent = 1;

        ws.Style.Alignment.Indent = 2;

        await Assert.That(ws.Cell("A1").Style.Alignment.Indent).IsEqualTo(1);
        await Assert.That(ws.Cell("B2").Style.Alignment.Indent).IsEqualTo(2);
    }

    /// <summary>
    /// Which containers keep the unchanged-value skip, and which decide from their own key. A
    /// selection of cells does neither; everything that holds a style of its own for its lifetime
    /// keeps the skip. The skip changes cost, not the cells, so only this can see a container lose
    /// it - a worksheet that did cost ~50 ms and 30 MB per redundant set on a 100K-cell sheet.
    /// </summary>
    [Test]
    public async Task Which_containers_skip_an_unchanged_value_and_which_decide_from_their_own_key()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        var data = wb.AddWorksheet("Data");
        var dataRange = data.Cell("A1").InsertData(new object[] { ("Name", "Price"), ("Cake", 9) });
        var pt = dataRange!.CreatePivotTable(ws.Cell("F1"), "pivot table");
        pt.RowLabels.Add("Name");
        pt.Values.Add("Price");

        var cell = (XLStyle)ws.Cell("A1").Style;
        var sheet = (XLStyle)ws.Style;
        var row = (XLStyle)ws.Row(1).Style;
        var column = (XLStyle)ws.Column(1).Style;
        var conditionalFormat = (XLStyle)ws.Range("A1:B2").AddConditionalFormat().WhenEquals(1);
        var pivotArea = (XLStyle)pt.StyleFormats.RowGrandTotalFormats.ForElement(XLPivotStyleFormatElement.All).Style;
        var detached = (XLStyle)XLWorkbook.DefaultStyle;

        var range = (XLStyle)ws.Range("A1:B2").Style;
        var rangeRow = (XLStyle)ws.Range("A1:B2").FirstRow()!.Style;
        var cells = (XLStyle)ws.Cells("A1:B2").Style;
        var ranges = (XLStyle)ws.Ranges("A1:B2,D4").Style;
        var rows = (XLStyle)ws.Rows("1:2").Style;
        var columns = (XLStyle)ws.Columns("A:B").Style;

        await Assert.That(cell.SkipsUnchangedValues).IsTrue();
        await Assert.That(sheet.SkipsUnchangedValues).IsTrue();
        await Assert.That(row.SkipsUnchangedValues).IsTrue();
        await Assert.That(column.SkipsUnchangedValues).IsTrue();
        await Assert.That(conditionalFormat.SkipsUnchangedValues).IsTrue();
        await Assert.That(pivotArea.SkipsUnchangedValues).IsTrue();
        await Assert.That(detached.SkipsUnchangedValues).IsTrue();

        await Assert.That(range.SkipsUnchangedValues).IsFalse();
        await Assert.That(rangeRow.SkipsUnchangedValues).IsFalse();
        await Assert.That(cells.SkipsUnchangedValues).IsFalse();
        await Assert.That(ranges.SkipsUnchangedValues).IsFalse();
        await Assert.That(rows.SkipsUnchangedValues).IsFalse();
        await Assert.That(columns.SkipsUnchangedValues).IsFalse();

        await Assert.That(cell.IsWholeStyle).IsTrue();
        await Assert.That(conditionalFormat.IsWholeStyle).IsTrue();
        await Assert.That(pivotArea.IsWholeStyle).IsTrue();
        await Assert.That(detached.IsWholeStyle).IsTrue();

        await Assert.That(sheet.IsWholeStyle).IsFalse();
        await Assert.That(row.IsWholeStyle).IsFalse();
        await Assert.That(column.IsWholeStyle).IsFalse();
        await Assert.That(range.IsWholeStyle).IsFalse();
        await Assert.That(cells.IsWholeStyle).IsFalse();
    }

    /// <summary>Apply <paramref name="change"/> to the worksheet's style in a frame of its own.</summary>
    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void OnWorksheet(IXLWorksheet ws, Action<IXLStyle> change) => change(ws.Style);

    /// <summary>Build the range and apply <paramref name="change"/> in a frame of its own.</summary>
    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void OnRange(IXLWorksheet ws, string address, Action<IXLStyle> change)
        => change(ws.Range(address).Style);

    /// <summary>Build the cells and apply <paramref name="change"/> in a frame of its own.</summary>
    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void OnCells(IXLWorksheet ws, string address, Action<IXLStyle> change)
        => change(ws.Cells(address).Style);
}
