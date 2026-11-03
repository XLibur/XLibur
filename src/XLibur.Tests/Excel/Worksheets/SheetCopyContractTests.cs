using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Worksheets;

/// <summary>
/// The parts of a sheet that <see cref="IXLWorksheet.CopyTo(string)"/> does <em>not</em> carry to the
/// copy. What a copy does carry is pinned part by part in <see cref="XLWorksheetTests"/>
/// (<c>CopyWorksheetPreserves*</c>) and the <c>SheetCopy*</c> fixtures.
/// </summary>
/// <remarks>
/// These tests pin today's behaviour, not a decision that it is right. Excel's Move or Copy does
/// copy charts and slicers. When a copy learns to carry one of them, flip its test here in the same
/// change, so the gap closes on purpose rather than by accident.
/// </remarks>
public class SheetCopyContractTests
{
    [Test]
    public async Task A_copy_does_not_carry_charts()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").InsertData(new[] { 1, 2, 3 });
        ws.Charts.Add(XLChartType.ColumnClustered);

        var copy = ws.CopyTo("Sheet2");

        await Assert.That(ws.Charts.Count()).IsEqualTo(1);
        await Assert.That(copy.Charts.Count()).IsEqualTo(0);
    }

    [Test]
    public async Task A_copy_carries_a_table_but_not_its_slicer()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "Region";
        ws.Cell("A2").Value = "North";
        ws.Cell("A3").Value = "South";
        var table = ws.Range("A1:A3").CreateTable();
        ws.Slicers.Add(table, "Region");

        var copy = ws.CopyTo("Sheet2");

        await Assert.That(copy.Tables.Count()).IsEqualTo(1);
        await Assert.That(ws.Slicers.Count()).IsEqualTo(1);
        await Assert.That(copy.Slicers.Count()).IsEqualTo(0);
    }

    [Test]
    public async Task A_copy_is_not_the_active_tab()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.SetTabActive();

        var copy = ws.CopyTo("Sheet2");

        await Assert.That(ws.TabActive).IsTrue();
        await Assert.That(copy.TabActive).IsFalse();
    }
}
