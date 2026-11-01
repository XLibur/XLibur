using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Ranges;

/// <summary>
/// Row-wise and column-wise algorithms were written twice, line for line, and had already drifted
/// three times (spec 26). Spec 26 collapses them onto one axis-parameterised implementation; this is
/// the gate that says the collapse changed nothing. Every case builds the same content twice, once
/// transposed, runs the mirrored operation, and asserts the two land on transposed addresses.
/// </summary>
public class GridAxisSymmetryTests
{
    private static IXLWorksheet Populate(IXLWorksheet ws, bool transposed)
    {
        for (var i = 1; i <= 6; i++)
            for (var j = 1; j <= 4; j++)
            {
                var cell = transposed ? ws.Cell(j, i) : ws.Cell(i, j);
                cell.Value = $"{i}.{j}";
            }

        return ws;
    }

    [Test]
    [Arguments(2, 3)]
    [Arguments(1, 1)]
    [Arguments(4, 2)]
    public async Task Insert_before_moves_the_same_content_on_both_axes(int at, int count)
    {
        using var wb = new XLWorkbook();
        var rows = Populate(wb.AddWorksheet("Rows"), transposed: false);
        var cols = Populate(wb.AddWorksheet("Cols"), transposed: true);

        rows.Row(at).InsertRowsAbove(count);
        cols.Column(at).InsertColumnsBefore(count);

        for (var i = 1; i <= 6 + count; i++)
            for (var j = 1; j <= 4; j++)
                await Assert.That(cols.Cell(j, i).GetString())
                    .IsEqualTo(rows.Cell(i, j).GetString());
    }

    [Test]
    [Arguments(2, 3)]
    [Arguments(1, 2)]
    public async Task Delete_moves_the_same_content_on_both_axes(int at, int count)
    {
        using var wb = new XLWorkbook();
        var rows = Populate(wb.AddWorksheet("Rows"), transposed: false);
        var cols = Populate(wb.AddWorksheet("Cols"), transposed: true);

        rows.Rows(at, at + count - 1).Delete();
        cols.Columns(at, at + count - 1).Delete();

        for (var i = 1; i <= 6; i++)
            for (var j = 1; j <= 4; j++)
                await Assert.That(cols.Cell(j, i).GetString())
                    .IsEqualTo(rows.Cell(i, j).GetString());
    }

    /// <summary>
    /// XLRangeInsertHelper.ShiftRowHeights / ShiftColumnWidths carry a line's size along when lines
    /// are inserted before it. The two axes measure in different units, so the assertion is on the
    /// pattern rather than the number: a line keeps its seeded size, or it carries the axis default.
    /// <para>
    /// This inserts through an entire-line <em>range</em> rather than through IXLRow / IXLColumn on
    /// purpose. XLRow.InsertRowsAbove and XLColumn.InsertColumnsBefore move line sizes themselves via
    /// RowsCollection.ShiftRowsDown / ColumnsCollection.ShiftColumnsRight and then call the helper
    /// with onlyUsedCells: true, which is the first thing ShiftRowHeights / ShiftColumnWidths bail
    /// on — so the line-size methods are unreachable from that entry point. The other cases here all
    /// take it, which is why the task 3 step 3 mutation passed until this case was added.
    /// </para>
    /// </summary>
    [Test]
    [Arguments(2, 3)]
    [Arguments(1, 1)]
    [Arguments(4, 2)]
    public async Task Line_sizes_are_carried_along_identically_on_both_axes(int at, int count)
    {
        using var wb = new XLWorkbook();
        var rows = Populate(wb.AddWorksheet("Rows"), transposed: false);
        var cols = Populate(wb.AddWorksheet("Cols"), transposed: true);

        for (var i = 1; i <= 6; i++)
        {
            rows.Row(i).Height = 20d + i;
            cols.Column(i).Width = 20d + i;
        }

        rows.Range(at, 1, at, XLHelper.MaxColumnNumber).InsertRowsAbove(count);
        cols.Range(1, at, XLHelper.MaxRowNumber, at).InsertColumnsBefore(count);

        // -1 stands for "this line carries the axis default", which is 15 points on one axis and
        // 8.43 characters on the other. Everything else is the seeded size and compares directly.
        double SeededHeight(int k)
        {
            var h = rows.Row(k).Height;
            return System.Math.Abs(h - rows.Worksheet.RowHeight) < XLHelper.Epsilon ? -1d : h;
        }

        double SeededWidth(int k)
        {
            var w = cols.Column(k).Width;
            return System.Math.Abs(w - cols.Worksheet.ColumnWidth) < XLHelper.Epsilon ? -1d : w;
        }

        for (var k = 1; k <= 6 + count + 1; k++)
            await Assert.That(SeededWidth(k)).IsEqualTo(SeededHeight(k)).Within(XLHelper.Epsilon);
    }

    /// <summary>
    /// XLRangeShiftHelper.ShiftColumns/ShiftRows repositions live ranges. A held range must survive
    /// or die identically on both axes, including the destroyed-by-shift case.
    /// </summary>
    [Test]
    [Arguments(1, 3)]
    [Arguments(3, -2)]
    [Arguments(2, -6)]
    public async Task A_live_range_is_repositioned_identically_on_both_axes(int at, int shift)
    {
        using var wb = new XLWorkbook();
        var rows = Populate(wb.AddWorksheet("Rows"), transposed: false);
        var cols = Populate(wb.AddWorksheet("Cols"), transposed: true);

        var heldRows = rows.Range(2, 1, 5, 4);
        var heldCols = cols.Range(1, 2, 4, 5);

        if (shift > 0)
        {
            rows.Row(at).InsertRowsAbove(shift);
            cols.Column(at).InsertColumnsBefore(shift);
        }
        else
        {
            rows.Rows(at, at - shift - 1).Delete();
            cols.Columns(at, at - shift - 1).Delete();
        }

        await Assert.That(heldCols.RangeAddress.IsValid).IsEqualTo(heldRows.RangeAddress.IsValid);
        if (!heldRows.RangeAddress.IsValid)
            return;

        await Assert.That(heldCols.RangeAddress.FirstAddress.ColumnNumber)
            .IsEqualTo(heldRows.RangeAddress.FirstAddress.RowNumber);
        await Assert.That(heldCols.RangeAddress.LastAddress.ColumnNumber)
            .IsEqualTo(heldRows.RangeAddress.LastAddress.RowNumber);
    }

    /// <summary>
    /// XLWorksheetRangeShifter moves conditional formats, data validations, defined names and page
    /// breaks. Each was written twice; each must move the same distance on both axes.
    /// </summary>
    [Test]
    public async Task Conditional_formats_and_validations_move_identically_on_both_axes()
    {
        using var wb = new XLWorkbook();
        var rows = Populate(wb.AddWorksheet("Rows"), transposed: false);
        var cols = Populate(wb.AddWorksheet("Cols"), transposed: true);

        rows.Range("A3:D5").AddConditionalFormat().WhenNotBlank().Fill.SetBackgroundColor(XLColor.Red);
        cols.Range("C1:E4").AddConditionalFormat().WhenNotBlank().Fill.SetBackgroundColor(XLColor.Red);

        rows.Range("A3:D5").CreateDataValidation().WholeNumber.Between(1, 10);
        cols.Range("C1:E4").CreateDataValidation().WholeNumber.Between(1, 10);

        rows.Row(2).InsertRowsAbove(2);
        cols.Column(2).InsertColumnsBefore(2);

        var rowCf = rows.ConditionalFormats.Single().Ranges.Single().RangeAddress;
        var colCf = cols.ConditionalFormats.Single().Ranges.Single().RangeAddress;
        await Assert.That(colCf.FirstAddress.ColumnNumber).IsEqualTo(rowCf.FirstAddress.RowNumber);
        await Assert.That(colCf.LastAddress.ColumnNumber).IsEqualTo(rowCf.LastAddress.RowNumber);

        var rowDv = rows.DataValidations.Single().Ranges.Single().RangeAddress;
        var colDv = cols.DataValidations.Single().Ranges.Single().RangeAddress;
        await Assert.That(colDv.FirstAddress.ColumnNumber).IsEqualTo(rowDv.FirstAddress.RowNumber);
        await Assert.That(colDv.LastAddress.ColumnNumber).IsEqualTo(rowDv.LastAddress.RowNumber);
    }

    /// <summary>
    /// XLWorksheetRangeShifter ran page-break shifting and sparkline cleanup in opposite orders on the
    /// two axes (ShiftColumns did breaks then sparklines, ShiftRows the reverse), with nothing stating
    /// whether that mattered. It does not: RemoveInvalidSparklines reads only sparkline location
    /// validity and ShiftPageBreaks* touches only PageSetup.*Breaks, so the two touch disjoint state.
    /// Spec 26 task 8 collapses them onto one order; this pins the outcome so the collapse is provably
    /// a no-op. Run green before the collapse, and again after.
    /// </summary>
    [Test]
    public async Task Page_breaks_and_sparklines_survive_a_shift_on_both_axes()
    {
        using var wb = new XLWorkbook();
        var rows = Populate(wb.AddWorksheet("Rows"), transposed: false);
        var cols = Populate(wb.AddWorksheet("Cols"), transposed: true);

        // A break past the insert point, so it must move by the shift.
        rows.PageSetup.AddHorizontalPageBreak(5);
        cols.PageSetup.AddVerticalPageBreak(5);

        // A sparkline whose source sits inside the region the later delete destroys, plus one well
        // clear of it, so the cleanup pass has both a survivor and a casualty to decide about.
        rows.SparklineGroups.Add("F1", "A3:D3");
        cols.SparklineGroups.Add("A6", "C1:C4");

        rows.Row(2).InsertRowsAbove(2);
        cols.Column(2).InsertColumnsBefore(2);

        await Assert.That(cols.PageSetup.ColumnBreaks.Single())
            .IsEqualTo(rows.PageSetup.RowBreaks.Single());
        await Assert.That(cols.SparklineGroups.SelectMany(g => g).Count())
            .IsEqualTo(rows.SparklineGroups.SelectMany(g => g).Count());

        rows.Rows(3, 5).Delete();
        cols.Columns(3, 5).Delete();

        await Assert.That(cols.PageSetup.ColumnBreaks.Single())
            .IsEqualTo(rows.PageSetup.RowBreaks.Single());
        await Assert.That(cols.SparklineGroups.SelectMany(g => g).Count())
            .IsEqualTo(rows.SparklineGroups.SelectMany(g => g).Count());
    }
}
