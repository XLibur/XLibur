using System.Linq;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.NamedRanges;

public class NamedRangeDeletionTests
{
    private static string RefersTo(XLWorkbook wb, string name)
    {
        return wb.DefinedNames.First(dn => dn.Name == name).RefersTo;
    }

    /// <summary>
    /// Regression for issue #2866. Deleting the top row of a named range must remove that row and shift the
    /// survivors up (A3:A4 -> A3:A3), matching Excel. Previously ClosedXML shifted both endpoints upward,
    /// expanding the range to A2:A3 and including a row that was never part of it.
    /// </summary>
    [Test]
    public async Task DeletingTopRowOfNamedRange_ShrinksAndShifts()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A3").Value = "deleted";
        ws.Cell("A4").Value = "survivor";
        ws.Range("A3:A4").AddToNamed("TopDelete", XLScope.Workbook);

        ws.Row(3).Delete();

        await Assert.That(RefersTo(wb, "TopDelete")).IsEqualTo("Sheet1!$A$3:$A$3");
    }

    /// <summary>
    /// Deleting several rows that overlap the top boundary clamps the first row to the deletion start and
    /// shifts the surviving bottom up: A3:A5 with rows 2:3 deleted becomes A2:A3.
    /// </summary>
    [Test]
    public async Task DeletingRowsOverlappingTopBoundary_ClampsFirstRow()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A3").Value = "top";
        ws.Cell("A4").Value = "mid";
        ws.Cell("A5").Value = "bottom";
        ws.Range("A3:A5").AddToNamed("OverlapTop", XLScope.Workbook);

        ws.Rows(2, 3).Delete();

        await Assert.That(RefersTo(wb, "OverlapTop")).IsEqualTo("Sheet1!$A$2:$A$3");
    }

    /// <summary>
    /// Deleting a row inside the range (not on the top boundary) shrinks it from within, leaving the top
    /// fixed: A3:A5 with row 4 deleted becomes A3:A4. This case was already correct and must not regress.
    /// </summary>
    [Test]
    public async Task DeletingMiddleRowOfNamedRange_ShrinksFromWithin()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A3").Value = "top";
        ws.Cell("A4").Value = "mid";
        ws.Cell("A5").Value = "bottom";
        ws.Range("A3:A5").AddToNamed("MidDelete", XLScope.Workbook);

        ws.Row(4).Delete();

        await Assert.That(RefersTo(wb, "MidDelete")).IsEqualTo("Sheet1!$A$3:$A$4");
    }

    /// <summary>
    /// Deleting a row entirely above the range shifts the whole range up without shrinking:
    /// A3:A4 with row 1 deleted becomes A2:A3. This case was already correct and must not regress.
    /// </summary>
    [Test]
    public async Task DeletingRowAboveNamedRange_ShiftsWholeRangeUp()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A3").Value = "a";
        ws.Cell("A4").Value = "b";
        ws.Range("A3:A4").AddToNamed("AboveDelete", XLScope.Workbook);

        ws.Row(1).Delete();

        await Assert.That(RefersTo(wb, "AboveDelete")).IsEqualTo("Sheet1!$A$2:$A$3");
    }

    /// <summary>
    /// The top-boundary shrink also applies across multiple columns: B3:D4 with row 3 deleted becomes B3:D3.
    /// </summary>
    [Test]
    public async Task DeletingTopRowOfMultiColumnNamedRange_ShrinksAndShifts()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Range("B3:D4").Value = "x";
        ws.Range("B3:D4").AddToNamed("Block", XLScope.Workbook);

        ws.Row(3).Delete();

        await Assert.That(RefersTo(wb, "Block")).IsEqualTo("Sheet1!$B$3:$D$3");
    }

    /// <summary>
    /// Deleting every row a named range covers invalidates it (ClosedXML/ClosedXML#880). Excel replaces the
    /// address with #REF! and keeps the sheet prefix; previously the shifted endpoints went negative and
    /// were clamped back to row 1, leaving the name pointing at a surviving row it never covered.
    /// </summary>
    [Test]
    public async Task DeletingAllRowsOfNamedRange_BecomesRefError()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Range("A1:B2").AddToNamed("Gone", XLScope.Workbook);

        ws.Rows(1, 5).Delete();

        await Assert.That(RefersTo(wb, "Gone")).IsEqualTo("Sheet1!#REF!");
        await Assert.That(wb.DefinedNames.ValidNamedRanges()).IsEmpty();
    }

    /// <summary>
    /// The deletion need not extend past the range: deleting exactly the rows it covers also leaves #REF!.
    /// </summary>
    [Test]
    public async Task DeletingExactlyTheRowsOfNamedRange_BecomesRefError()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Range("A3:A4").AddToNamed("Exact", XLScope.Workbook);

        ws.Rows(3, 4).Delete();

        await Assert.That(RefersTo(wb, "Exact")).IsEqualTo("Sheet1!#REF!");
    }

    /// <summary>
    /// The same for a single cell: deleting the row it sits on leaves #REF!, not a neighbouring cell.
    /// </summary>
    [Test]
    public async Task DeletingRowOfSingleCellNamedRange_BecomesRefError()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A3").AddToNamed("Cell", XLScope.Workbook);

        ws.Row(3).Delete();

        await Assert.That(RefersTo(wb, "Cell")).IsEqualTo("Sheet1!#REF!");
    }

    /// <summary>
    /// A row-only reference is invalidated the same way: rows 3:4 with those rows deleted become #REF!.
    /// </summary>
    [Test]
    public async Task DeletingAllRowsOfRowOnlyNamedRange_BecomesRefError()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        wb.DefinedNames.Add("Rows", "Sheet1!$3:$4");

        ws.Rows(1, 5).Delete();

        await Assert.That(RefersTo(wb, "Rows")).IsEqualTo("Sheet1!#REF!");
    }

    /// <summary>
    /// A cell formula goes through the same shifter, so deleting every row it reads leaves #REF! rather
    /// than silently repointing the formula at whatever moved up into those rows.
    /// </summary>
    [Test]
    public async Task DeletingAllRowsReferencedByFormula_BecomesRefError()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Range("A1:A2").Value = 1;
        ws.Cell("D10").FormulaA1 = "=SUM(A1:A2)";

        ws.Rows(1, 5).Delete();

        await Assert.That(ws.Cell("D5").FormulaA1).IsEqualTo("SUM(#REF!)");
    }

    /// <summary>
    /// Row-only counterpart of DeletingExactlyTheRowsOfNamedRange_BecomesRefError. A multi-row delete is
    /// applied one row at a time, so the reference has to shrink correctly on the way down for the last
    /// pass to see it as fully deleted.
    /// </summary>
    [Test]
    public async Task DeletingExactlyTheRowsOfRowOnlyNamedRange_BecomesRefError()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        wb.DefinedNames.Add("Rows", "Sheet1!$3:$4");

        ws.Rows(3, 4).Delete();

        await Assert.That(RefersTo(wb, "Rows")).IsEqualTo("Sheet1!#REF!");
    }

    /// <summary>
    /// Row-only counterpart of the #2866 top-boundary shrink: rows 3:5 with row 3 deleted becomes 3:4, not
    /// 2:4. The row-only branch shifted both endpoints regardless of where the deletion started.
    /// </summary>
    [Test]
    public async Task DeletingTopRowOfRowOnlyNamedRange_ShrinksAndShifts()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        wb.DefinedNames.Add("Rows", "Sheet1!$3:$5");

        ws.Row(3).Delete();

        await Assert.That(RefersTo(wb, "Rows")).IsEqualTo("Sheet1!$3:$4");
    }

    /// <summary>
    /// Deleting a row inside a row-only reference shrinks it from within, leaving the top fixed.
    /// </summary>
    [Test]
    public async Task DeletingMiddleRowOfRowOnlyNamedRange_ShrinksFromWithin()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        wb.DefinedNames.Add("Rows", "Sheet1!$3:$5");

        ws.Row(4).Delete();

        await Assert.That(RefersTo(wb, "Rows")).IsEqualTo("Sheet1!$3:$4");
    }

    /// <summary>
    /// A deletion that starts above a row-only reference and ends inside it keeps the surviving top row
    /// where it is: rows 2:5 with rows 3:4 deleted becomes 2:3.
    /// </summary>
    [Test]
    public async Task DeletingRowsInsideRowOnlyNamedRange_KeepsSurvivingTop()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        wb.DefinedNames.Add("Rows", "Sheet1!$2:$5");

        ws.Rows(3, 4).Delete();

        await Assert.That(RefersTo(wb, "Rows")).IsEqualTo("Sheet1!$2:$3");
    }

    /// <summary>
    /// A row-only reference entirely below the deletion shifts up whole: rows 3:4 with row 1 deleted
    /// becomes 2:3. This case was already correct and must not regress.
    /// </summary>
    [Test]
    public async Task DeletingRowAboveRowOnlyNamedRange_ShiftsWholeRangeUp()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        wb.DefinedNames.Add("Rows", "Sheet1!$3:$4");

        ws.Row(1).Delete();

        await Assert.That(RefersTo(wb, "Rows")).IsEqualTo("Sheet1!$2:$3");
    }

    /// <summary>
    /// A deletion overshooting the top of a row-only reference leaves only what survived, moved up:
    /// rows 3:6 with rows 1:5 deleted becomes 1:1. This case was already correct and must not regress.
    /// </summary>
    [Test]
    public async Task DeletingRowsOvershootingRowOnlyNamedRange_KeepsSurvivor()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        wb.DefinedNames.Add("Rows", "Sheet1!$3:$6");

        ws.Rows(1, 5).Delete();

        await Assert.That(RefersTo(wb, "Rows")).IsEqualTo("Sheet1!$1:$1");
    }

    /// <summary>
    /// Column-only reference, exact cover: columns C:D with those columns deleted become #REF!.
    /// </summary>
    [Test]
    public async Task DeletingExactlyTheColumnsOfColumnOnlyNamedRange_BecomesRefError()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        wb.DefinedNames.Add("Cols", "Sheet1!$C:$D");

        ws.Columns(3, 4).Delete();

        await Assert.That(RefersTo(wb, "Cols")).IsEqualTo("Sheet1!#REF!");
    }

    /// <summary>
    /// Column-only counterpart of the left-boundary shrink: columns C:E with column C deleted become C:D.
    /// </summary>
    [Test]
    public async Task DeletingLeftColumnOfColumnOnlyNamedRange_ShrinksAndShifts()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        wb.DefinedNames.Add("Cols", "Sheet1!$C:$E");

        ws.Column(3).Delete();

        await Assert.That(RefersTo(wb, "Cols")).IsEqualTo("Sheet1!$C:$D");
    }

    /// <summary>
    /// Column-only counterpart of the surviving-left case: columns B:E with columns C:D deleted become B:C.
    /// </summary>
    [Test]
    public async Task DeletingColumnsInsideColumnOnlyNamedRange_KeepsSurvivingLeft()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        wb.DefinedNames.Add("Cols", "Sheet1!$B:$E");

        ws.Columns(3, 4).Delete();

        await Assert.That(RefersTo(wb, "Cols")).IsEqualTo("Sheet1!$B:$C");
    }

    /// <summary>
    /// A column-only reference entirely to the right of the deletion shifts left whole: columns C:D with
    /// column A deleted become B:C. This case was already correct and must not regress.
    /// </summary>
    [Test]
    public async Task DeletingColumnLeftOfColumnOnlyNamedRange_ShiftsWholeRangeLeft()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        wb.DefinedNames.Add("Cols", "Sheet1!$C:$D");

        ws.Column(1).Delete();

        await Assert.That(RefersTo(wb, "Cols")).IsEqualTo("Sheet1!$B:$C");
    }

    /// <summary>
    /// Column counterpart of the invalidation: deleting every column a named range covers leaves #REF!
    /// instead of clamping the endpoints back to column A.
    /// </summary>
    [Test]
    public async Task DeletingAllColumnsOfNamedRange_BecomesRefError()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Range("A1:B2").AddToNamed("GoneCols", XLScope.Workbook);

        ws.Columns(1, 5).Delete();

        await Assert.That(RefersTo(wb, "GoneCols")).IsEqualTo("Sheet1!#REF!");
        await Assert.That(wb.DefinedNames.ValidNamedRanges()).IsEmpty();
    }

    /// <summary>
    /// Column counterpart of the top-row bug: deleting the left column of a named range removes it and
    /// shifts survivors left (C1:D1 -> C1:C1) rather than expanding to B1:C1.
    /// </summary>
    [Test]
    public async Task DeletingLeftColumnOfNamedRange_ShrinksAndShifts()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("C1").Value = "deleted";
        ws.Cell("D1").Value = "survivor";
        ws.Range("C1:D1").AddToNamed("LeftDelete", XLScope.Workbook);

        ws.Column(3).Delete();

        await Assert.That(RefersTo(wb, "LeftDelete")).IsEqualTo("Sheet1!$C$1:$C$1");
    }

    /// <summary>
    /// Deleting a column entirely to the left of the range shifts the whole range left without shrinking:
    /// C1:D1 with column A deleted becomes B1:C1. Guards against over-clamping the column path.
    /// </summary>
    [Test]
    public async Task DeletingColumnLeftOfNamedRange_ShiftsWholeRangeLeft()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("C1").Value = "a";
        ws.Cell("D1").Value = "b";
        ws.Range("C1:D1").AddToNamed("ColAbove", XLScope.Workbook);

        ws.Column(1).Delete();

        await Assert.That(RefersTo(wb, "ColAbove")).IsEqualTo("Sheet1!$B$1:$C$1");
    }
}
