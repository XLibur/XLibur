using XLibur.Excel;
using XLibur.Excel.Patterns;
using XLibur.Excel.Ranges.Index;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel.Coordinates;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Ranges;

public class RangeIndexTest
{
    private const int TestCount = 10000;

    [Test]
    public async Task FindExistingMatches()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
        var index = FillIndexWithTestData(ws);

        for (var i = 1; i <= TestCount; i++)
        {
            for (var j = 2; j <= 4; j++)
            {
                var address = new XLAddress(ws, i * 2, j, false, false);
                await Assert.That(index.Contains(in address)).IsTrue();
            }
        }
    }

    [Test]
    public async Task FindNonExistingMatches()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
        var index = FillIndexWithTestData(ws);

        for (var i = 1; i <= TestCount; i++)
        {
            var address = new XLAddress(ws, i * 2 + 1, 3, false, false);
            await Assert.That(index.Contains(in address)).IsFalse();
        }
    }

    [Test]
    public async Task FindExistingIntersections()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
        var index = FillIndexWithTestData(ws);

        for (var i = 1; i <= TestCount; i++)
        {
            var rangeAddress = new XLRangeAddress(
                new XLAddress(ws, i * 2, 1 + i % 4, false, false),
                new XLAddress(ws, i * 2 + 1, 8 - i % 3, false, false));

            await Assert.That(index.Intersects(in rangeAddress)).IsTrue();
        }

        for (var i = 2; i < 4; i++)
        {
            var columnAddress = XLRangeAddress.EntireColumn(ws, i);
            await Assert.That(index.Intersects(in columnAddress)).IsTrue();
        }
    }

    [Test]
    public async Task FindNonExistingIntersections()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
        var index = FillIndexWithTestData(ws);

        for (var i = 1; i <= TestCount; i++)
        {
            var rangeAddress = new XLRangeAddress(
                new XLAddress(ws, i * 2 + 1, 1 + i % 4, false, false),
                new XLAddress(ws, i * 2 + 1, 8 - i % 3, false, false));

            await Assert.That(index.Intersects(in rangeAddress)).IsFalse();
        }

        var columnAddress = XLRangeAddress.EntireColumn(ws, 1);
        await Assert.That(index.Intersects(in columnAddress)).IsFalse();
        columnAddress = XLRangeAddress.EntireColumn(ws, 5);
        await Assert.That(index.Intersects(in columnAddress)).IsFalse();
    }

    [Test]
    public async Task FindMatchAfterColumnShifting()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
        var index = FillIndexWithTestData(ws);

        ws.Column(1).InsertColumnsBefore(1000);

        var address = new XLAddress(ws, 102, 1004, false, false);

        await Assert.That(index.Contains(in address)).IsTrue();
    }

    [Test]
    public async Task FindIntersectionsAfterColumnShifting()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
        var index = FillIndexWithTestData(ws);

        ws.Column(3).InsertColumnsBefore(2);

        var rangeAddress = new XLRangeAddress(ws, "F102:E103");

        await Assert.That(index.Intersects(in rangeAddress)).IsTrue();
    }

    [Test]
    public async Task FindMatchAfterRowShifting()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
        var index = FillIndexWithTestData(ws);

        ws.Row(10).InsertRowsBelow(3);

        var address = new XLAddress(ws, 103, 4, false, false);

        await Assert.That(index.Contains(in address)).IsTrue();
    }

    [Test]
    public async Task FindIntersectionsAfterRowShifting()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
        var index = FillIndexWithTestData(ws);

        ws.Row(10).InsertRowsBelow(3);

        var rangeAddress = new XLRangeAddress(ws, "C103:E103");

        await Assert.That(index.Intersects(in rangeAddress)).IsTrue();
    }

    [Test]
    public async Task CreateQuadTree()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
        var quadTree = new Quadrant();
        var range = ws.Range("BT76:CA87");

        quadTree.Add(range);

        var level0 = quadTree;
        await Assert.That(level0.MinimumColumn).IsEqualTo(1);
        await Assert.That(level0.MaximumColumn).IsEqualTo(XLHelper.MaxColumnNumber);
        await Assert.That(level0.MinimumRow).IsEqualTo(1);
        await Assert.That(level0.MaximumRow).IsEqualTo(XLHelper.MaxRowNumber);
        await Assert.That(level0.Ranges).IsNull();
        await Assert.That(level0.Children.Count).IsEqualTo(128);
        await Assert.That(level0.Children.All(child => child.Level == 1)).IsTrue();
        await Assert.That(level0.Children.Count(child =>
            child.MinimumColumn == 1 &&
            child.MaximumColumn == 8192 &&
            child.X == 0)).IsEqualTo(64);
        await Assert.That(level0.Children.Count(child =>
            child.MinimumColumn == 8193 &&
            child.MaximumColumn == 16384 &&
            child.X == 1)).IsEqualTo(64);
        await Assert.That(level0.Children.Count(child =>
            child.MinimumRow == 1 &&
            child.MaximumRow == 8192 &&
            child.Y == 0)).IsEqualTo(2);
        await Assert.That(level0.Children.Count(child =>
            child.MinimumRow == 16385 &&
            child.MaximumRow == 24576 &&
            child.Y == 2)).IsEqualTo(2);

        await Assert.That(level0.Children[0].Children.Any()).IsTrue();
        await Assert.That(level0.Children.Skip(1).All(child => child.Children == null)).IsTrue();

        var level8 = level0
            .Children[0] // 1
            .Children[0] // 2
            .Children[0] // 3
            .Children[0] // 4
            .Children[0] // 5
            .Children[0] // 6
            .Children[0] // 7
            .Children[^1]; // 8

        await Assert.That(level8.MinimumColumn).IsEqualTo(65);
        await Assert.That(level8.MinimumRow).IsEqualTo(65);
        await Assert.That(level8.MaximumColumn).IsEqualTo(128);
        await Assert.That(level8.MaximumRow).IsEqualTo(128);

        var level9 = level8.Children[0];
        await Assert.That(level9.Ranges).IsNotNull();
        await Assert.That(level9.Ranges.Single()).IsEqualTo(range);
    }

    [Test]
    public async Task XLRangesCountChangesCorrectly()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
        var range1 = ws.Range("A1:B2");
        var range2 = ws.Range("A2:B3");
        var range3 = ws.Range("A1:B2"); // same as range1

        var ranges = new XLRanges { range1 };
        await Assert.That(ranges.Count).IsEqualTo(1);
        ranges.Add(range2);
        await Assert.That(ranges.Count).IsEqualTo(2);
        ranges.Add(range3);
        await Assert.That(ranges.Count).IsEqualTo(2);

        // Add many entries to activate QuadTree
        for (var i = 1; i <= TestCount; i++)
        {
            ranges.Add(ws.Range(i * 2, 2, i * 2, 4));
        }

        await Assert.That(ranges.Count).IsEqualTo(2 + TestCount);

        for (var i = 1; i <= TestCount; i++)
        {
            ranges.Remove(ws.Range(i * 2, 2, i * 2, 4));
        }

        await Assert.That(ranges.Count).IsEqualTo(2);

        ranges.Remove(range3);
        await Assert.That(ranges.Count).IsEqualTo(1);
        ranges.Remove(range2);
        await Assert.That(ranges.Count).IsEqualTo(0);
        ranges.Remove(range1);
        await Assert.That(ranges.Count).IsEqualTo(0);
    }

    private static XLRangeIndex<IXLRangeBase> CreateRangeIndex(IXLWorksheet worksheet)
    {
        return new XLRangeIndex<IXLRangeBase>((XLWorksheet)worksheet);
    }

    private static XLRangeIndex<IXLRangeBase> FillIndexWithTestData(IXLWorksheet worksheet)
    {
        var ranges = new List<IXLRange>();
        for (var i = 1; i <= TestCount; i++)
        {
            ranges.Add(worksheet.Range(i * 2, 2, i * 2, 4));
        }

        var index = CreateRangeIndex(worksheet);
        ranges.ForEach(r => index.Add(r));
        return index;
    }
}
