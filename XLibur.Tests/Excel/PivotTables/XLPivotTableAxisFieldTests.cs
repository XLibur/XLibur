using System;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.PivotTables;
/// <summary>
/// Tests methods of interface <see cref="IXLPivotField"/> implemented through <see cref="XLPivotTableAxisField"/>.
/// </summary>
internal class XLPivotTableAxisFieldTests
{
    [Test]
    public async Task CustomName_can_be_changed()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Cell("A1").InsertData(new object[]
        {
            ("ID", "Color", "Count"),
            (1, "Blue", 10),
        });
        var pt = ws.PivotTables.Add("pt", ws.Cell("E1"), range!);
        var colorField = pt.RowLabels.Add("Color");

        colorField.SetCustomName("Changed color");

        await Assert.That(pt.RowLabels.Get(0).CustomName).IsEqualTo("Changed color");
    }

    [Test]
    public async Task CustomName_throws_exception_when_name_is_already_used()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Cell("A1").InsertData(new object[]
        {
            ("ID", "Color", "Count"),
            (1, "Blue", 10),
        });
        var pt = ws.PivotTables.Add("pt", ws.Cell("E1"), range!);
        var idField = pt.RowLabels.Add("ID", "Custom ID");
        var colorField = pt.RowLabels.Add("Color");

        var ex1 = await Assert.That(() => idField.SetCustomName("Color")).Throws<ArgumentException>();
        await Assert.That(ex1.Message).IsEqualTo("Custom name 'Color' is already used by another field.");
        var ex2 = await Assert.That(() => colorField.SetCustomName("Custom ID")).Throws<ArgumentException>();
        await Assert.That(ex2.Message).IsEqualTo("Custom name 'Custom ID' is already used by another field.");
    }
}
