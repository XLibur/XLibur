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
    /// A chart anchor is a pair of raw <c>int</c>s on <see cref="XLDrawingPosition"/> and nothing
    /// notifies it, so a chart anchored below an insert stays where it was. Spec 33 task 4 fixes
    /// this and re-points this test.
    /// </summary>
    [Test]
    public async Task A_chart_anchor_does_not_move_yet()
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

        // WRONG on purpose. Correct is 13 / 23 once spec 33 task 4 lands.
        await Assert.That(chart.Position.Row).IsEqualTo(10);
        await Assert.That(chart.SecondPosition.Row).IsEqualTo(20);
    }

    /// <summary>
    /// The note case is half right, and that is the defect. The note itself lives in the misc slice
    /// and moves with the cells; its callout box is an <see cref="XLDrawingPosition"/> nobody
    /// notifies, so the box stays pinned three rows above where the note now is.
    /// <c>XLComment.cs</c> already documents the mechanism. Spec 33 task 4 fixes this and re-points
    /// this test — both halves, so that fixing one and not the other cannot pass.
    /// </summary>
    [Test]
    public async Task A_note_anchor_does_not_move_yet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.Cell("C10").CreateComment().AddText("note");

        ws.Row(1).InsertRowsAbove(3);

        // The note moved with the misc slice. This half is right and stays right.
        await Assert.That(ws.Cell("C10").HasComment).IsFalse();
        await Assert.That(ws.Cell("C13").HasComment).IsTrue();

        // WRONG on purpose. Correct is 12 once spec 33 task 4 lands.
        await Assert.That(ws.Cell("C13").GetComment().Position.Row).IsEqualTo(9);
    }

    /// <summary>
    /// <c>SplitRow</c> and <c>SplitColumn</c> are raw <c>int</c>s on <see cref="XLSheetView"/> and
    /// nothing notifies them. Spec 33 task 5 fixes this and re-points this test.
    /// </summary>
    [Test]
    public async Task Freeze_panes_do_not_move_yet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.SheetView.Freeze(5, 4);

        ws.Row(1).InsertRowsAbove(3);
        ws.Column(1).InsertColumnsBefore(2);

        // WRONG on purpose. Correct is 8 / 6 once spec 33 task 5 lands.
        await Assert.That(ws.SheetView.SplitRow).IsEqualTo(5);
        await Assert.That(ws.SheetView.SplitColumn).IsEqualTo(4);
    }

    /// <summary>
    /// <c>XLPivotTable.Area</c> is a raw <see cref="XLibur.Excel.Coordinates.Area"/> and nothing
    /// notifies it. Spec 33 task 5 fixes this and re-points this test.
    /// </summary>
    [Test]
    public async Task A_pivot_area_does_not_move_yet()
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

        // WRONG on purpose. Correct is 13 once spec 33 task 5 lands.
        await Assert.That(pt.Area.FirstPoint.Row).IsEqualTo(10);
        await Assert.That(pt.Area.FirstPoint.Column).IsEqualTo(4);
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
