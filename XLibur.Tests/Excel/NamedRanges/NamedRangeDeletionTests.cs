using System.Linq;
using XLibur.Excel;
using NUnit.Framework;

namespace XLibur.Tests.Excel.NamedRanges;

[TestFixture]
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
    public void DeletingTopRowOfNamedRange_ShrinksAndShifts()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A3").Value = "deleted";
        ws.Cell("A4").Value = "survivor";
        ws.Range("A3:A4").AddToNamed("TopDelete", XLScope.Workbook);

        ws.Row(3).Delete();

        Assert.That(RefersTo(wb, "TopDelete"), Is.EqualTo("Sheet1!$A$3:$A$3"));
    }

    /// <summary>
    /// Deleting several rows that overlap the top boundary clamps the first row to the deletion start and
    /// shifts the surviving bottom up: A3:A5 with rows 2:3 deleted becomes A2:A3.
    /// </summary>
    [Test]
    public void DeletingRowsOverlappingTopBoundary_ClampsFirstRow()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A3").Value = "top";
        ws.Cell("A4").Value = "mid";
        ws.Cell("A5").Value = "bottom";
        ws.Range("A3:A5").AddToNamed("OverlapTop", XLScope.Workbook);

        ws.Rows(2, 3).Delete();

        Assert.That(RefersTo(wb, "OverlapTop"), Is.EqualTo("Sheet1!$A$2:$A$3"));
    }

    /// <summary>
    /// Deleting a row inside the range (not on the top boundary) shrinks it from within, leaving the top
    /// fixed: A3:A5 with row 4 deleted becomes A3:A4. This case was already correct and must not regress.
    /// </summary>
    [Test]
    public void DeletingMiddleRowOfNamedRange_ShrinksFromWithin()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A3").Value = "top";
        ws.Cell("A4").Value = "mid";
        ws.Cell("A5").Value = "bottom";
        ws.Range("A3:A5").AddToNamed("MidDelete", XLScope.Workbook);

        ws.Row(4).Delete();

        Assert.That(RefersTo(wb, "MidDelete"), Is.EqualTo("Sheet1!$A$3:$A$4"));
    }

    /// <summary>
    /// Deleting a row entirely above the range shifts the whole range up without shrinking:
    /// A3:A4 with row 1 deleted becomes A2:A3. This case was already correct and must not regress.
    /// </summary>
    [Test]
    public void DeletingRowAboveNamedRange_ShiftsWholeRangeUp()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A3").Value = "a";
        ws.Cell("A4").Value = "b";
        ws.Range("A3:A4").AddToNamed("AboveDelete", XLScope.Workbook);

        ws.Row(1).Delete();

        Assert.That(RefersTo(wb, "AboveDelete"), Is.EqualTo("Sheet1!$A$2:$A$3"));
    }

    /// <summary>
    /// The top-boundary shrink also applies across multiple columns: B3:D4 with row 3 deleted becomes B3:D3.
    /// </summary>
    [Test]
    public void DeletingTopRowOfMultiColumnNamedRange_ShrinksAndShifts()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Range("B3:D4").Value = "x";
        ws.Range("B3:D4").AddToNamed("Block", XLScope.Workbook);

        ws.Row(3).Delete();

        Assert.That(RefersTo(wb, "Block"), Is.EqualTo("Sheet1!$B$3:$D$3"));
    }

    /// <summary>
    /// Column counterpart of the top-row bug: deleting the left column of a named range removes it and
    /// shifts survivors left (C1:D1 -> C1:C1) rather than expanding to B1:C1.
    /// </summary>
    [Test]
    public void DeletingLeftColumnOfNamedRange_ShrinksAndShifts()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("C1").Value = "deleted";
        ws.Cell("D1").Value = "survivor";
        ws.Range("C1:D1").AddToNamed("LeftDelete", XLScope.Workbook);

        ws.Column(3).Delete();

        Assert.That(RefersTo(wb, "LeftDelete"), Is.EqualTo("Sheet1!$C$1:$C$1"));
    }

    /// <summary>
    /// Deleting a column entirely to the left of the range shifts the whole range left without shrinking:
    /// C1:D1 with column A deleted becomes B1:C1. Guards against over-clamping the column path.
    /// </summary>
    [Test]
    public void DeletingColumnLeftOfNamedRange_ShiftsWholeRangeLeft()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("C1").Value = "a";
        ws.Cell("D1").Value = "b";
        ws.Range("C1:D1").AddToNamed("ColAbove", XLScope.Workbook);

        ws.Column(1).Delete();

        Assert.That(RefersTo(wb, "ColAbove"), Is.EqualTo("Sheet1!$B$1:$C$1"));
    }
}
