using System;
using System.Linq;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Sparklines;

public class SparklineShiftTests
{
    [Test]
    public async Task SparklineAreShiftedOnColumnInsert()
    {
        await AssertSparklinePosition("D2", ws => ws.Column("C").InsertColumnsAfter(2), "F2");
    }

    [Test]
    public async Task SparklineAreShiftedOnColumnDelete()
    {
        await AssertSparklinePosition("F2", ws => ws.Column("C").Delete(), "E2");
    }

    [Test]
    public async Task SparklineColumnShiftedOutOfSheetAreRemoved()
    {
        await AssertSparklinePosition("XFD1", ws => ws.Column("C").InsertColumnsAfter(1), null);
    }

    [Test]
    public async Task SparklineAreShiftedOnRowInsert()
    {
        await AssertSparklinePosition("B3", ws => ws.Row(2).InsertRowsBelow(3), "B6");
    }

    [Test]
    public async Task SparklineAreShiftedOnRowDelete()
    {
        await AssertSparklinePosition("F8", ws => ws.Rows(4, 6).Delete(), "F5");
    }

    [Test]
    public async Task SparklineRowShiftedOutOfSheetAreRemoved()
    {
        await AssertSparklinePosition($"A{XLHelper.MaxRowNumber}", ws => ws.Row(2).InsertRowsBelow(1), null);
    }

    private static async Task AssertSparklinePosition(string sparklineAddress, Action<IXLWorksheet> insertAction, string expectedAddress)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B2").Value = 1;
        ws.Cell("C2").Value = 2;
        var sparklineGroup = ws.SparklineGroups.Add(sparklineAddress, "B2:C2");
        insertAction(ws);
        await Assert.That(sparklineGroup.SingleOrDefault()?.Location.Address.ToString()).IsEqualTo(expectedAddress);
        if (expectedAddress is null)
            await Assert.That(sparklineGroup).IsEmpty();
    }
}
