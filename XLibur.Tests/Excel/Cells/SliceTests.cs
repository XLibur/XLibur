using XLibur.Excel;
using XLibur.Excel.Coordinates;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Cells;

public class SliceTests
{
    [Test]
    public async Task Stores_Values()
    {
        var slice = new Slice<int>();
        var point = new XLSheetPoint(574, 241);
        slice.Set(point, 1);
        await Assert.That(slice[point]).IsEqualTo(1);
    }

    [Test]
    public async Task Setting_Value_To_Default_Clears_Element()
    {
        var slice = new Slice<int>();
        var point = new XLSheetPoint(574, 241);
        slice.Set(point, 1);
        await Assert.That(slice.MaxRow).IsEqualTo(574);
        await Assert.That(slice.MaxColumn).IsEqualTo(241);

        slice.Set(point, 0);

        await Assert.That(slice.MaxRow).IsEqualTo(0);
        await Assert.That(slice.MaxColumn).IsEqualTo(0);
    }

    [Test]
    public async Task Keeps_Track_Of_Max_Used_Coordinates()
    {
        var slice = new Slice<int>();
        slice.Set(54, 32, 1);
        slice.Set(140, 32, 1);
        slice.Set(140, 72, 1);

        await Assert.That(slice.MaxRow).IsEqualTo(140);
        await Assert.That(slice.MaxColumn).IsEqualTo(72);

        slice.Set(140, 72, 0);

        await Assert.That(slice.MaxRow).IsEqualTo(140);
        await Assert.That(slice.MaxColumn).IsEqualTo(32);

        slice.Set(140, 32, 0);

        await Assert.That(slice.MaxRow).IsEqualTo(54);
        await Assert.That(slice.MaxColumn).IsEqualTo(32);

        slice.Set(54, 32, 0);

        await Assert.That(slice.MaxRow).IsEqualTo(0);
        await Assert.That(slice.MaxColumn).IsEqualTo(0);
    }

    [Test]
    public async Task Keeps_Track_Of_Used_Rows()
    {
        var slice = new Slice<int>();
        await Assert.That(slice.UsedRows).IsEmpty();

        slice.Set(new XLSheetPoint(1, 1), 1);
        await Assert.That(slice.UsedRows).IsEquivalentTo([1]);

        slice.Set(new XLSheetPoint(70, 1), 1);
        await Assert.That(slice.UsedRows).IsEquivalentTo([1, 70]);

        slice.Set(new XLSheetPoint(35, 1), 1);
        await Assert.That(slice.UsedRows).IsEquivalentTo([1, 35, 70]);

        slice.Set(new XLSheetPoint(35, 2), 1);
        await Assert.That(slice.UsedRows).IsEquivalentTo([1, 35, 70]);

        slice.Set(new XLSheetPoint(35, 1), 0);
        await Assert.That(slice.UsedRows).IsEquivalentTo([1, 35, 70]);

        slice.Set(new XLSheetPoint(35, 2), 0);
        await Assert.That(slice.UsedRows).IsEquivalentTo([1, 70]);

        slice.Set(new XLSheetPoint(1, 1), 0);
        await Assert.That(slice.UsedRows).IsEquivalentTo([70]);

        slice.Set(new XLSheetPoint(70, 1), 0);
        await Assert.That(slice.UsedRows).IsEmpty();
    }

    [Test]
    public async Task Keeps_Track_Of_Used_Columns()
    {
        var slice = new Slice<int>();
        await Assert.That(slice.UsedColumns).IsEmpty();

        slice.Set(new XLSheetPoint(1, 5), 1);
        await Assert.That(slice.UsedColumns).IsEquivalentTo([5]);

        slice.Set(new XLSheetPoint(1, 750), 1);
        await Assert.That(slice.UsedColumns).IsEquivalentTo([5, 750]);

        slice.Set(new XLSheetPoint(1, 90), 1);
        await Assert.That(slice.UsedColumns).IsEquivalentTo([5, 90, 750]);

        slice.Set(new XLSheetPoint(2, 5), 1);
        await Assert.That(slice.UsedColumns).IsEquivalentTo([5, 90, 750]);

        slice.Set(new XLSheetPoint(1, 5), 0);
        await Assert.That(slice.UsedColumns).IsEquivalentTo([5, 90, 750]);

        slice.Set(new XLSheetPoint(2, 5), 0);
        await Assert.That(slice.UsedColumns).IsEquivalentTo([90, 750]);

        slice.Set(new XLSheetPoint(1, 750), 0);
        await Assert.That(slice.UsedColumns).IsEquivalentTo([90]);

        slice.Set(new XLSheetPoint(1, 90), 0);
        await Assert.That(slice.UsedColumns).IsEmpty();
    }

    [Test]
    public async Task Clear_Range_Sets_Values_To_Default()
    {
        var slice = new Slice<int>();
        var outsideAddress = new XLSheetPoint(1, 1);
        slice.Set(outsideAddress, 1);
        var firstCorner = new XLSheetPoint(50, 20);
        slice.Set(firstCorner, 1);
        var insideAddress = new XLSheetPoint(55, 22);
        slice.Set(insideAddress, 1);
        var lastCorner = new XLSheetPoint(60, 30);
        slice.Set(lastCorner, 1);

        slice.Clear(new XLSheetRange(firstCorner, lastCorner));
        await Assert.That(slice[outsideAddress]).IsEqualTo(1);
        await Assert.That(slice[firstCorner]).IsEqualTo(0);
        await Assert.That(slice[insideAddress]).IsEqualTo(0);
        await Assert.That(slice[lastCorner]).IsEqualTo(0);
    }

    [Test]
    public async Task InsertAreaAndShiftDown_Moves_Area_Cells_Down_And_Purges_Values_Outside_Worksheet()
    {
        var slice = new Slice<int>();
        slice.Set(1, 1, 1);
        slice.Set(3, 1, 2);
        var purgedAddress = new XLSheetPoint(XLHelper.MaxRowNumber, 2);
        slice.Set(purgedAddress, 3);

        var outsideAddress = new XLSheetPoint(1, 3);
        slice.Set(outsideAddress, 4);

        slice.InsertAreaAndShiftDown(new XLSheetRange(new XLSheetPoint(1, 1), new XLSheetPoint(2, 2)));

        await Assert.That(slice[3, 1]).IsEqualTo(1);
        await Assert.That(slice[5, 1]).IsEqualTo(2);
        await Assert.That(slice[XLHelper.MaxRowNumber, 2]).IsEqualTo(0);
        await Assert.That(slice[outsideAddress]).IsEqualTo(4);
    }

    [Test]
    public async Task InsertAreaAndShiftRight_Moves_Area_Cells_Down_And_Purges_Values_Outside_Worksheet()
    {
        var slice = new Slice<int>();
        slice.Set(1, 1, 1);
        slice.Set(1, 3, 2);
        var purgedAddress = new XLSheetPoint(2, XLHelper.MaxColumnNumber);
        slice.Set(purgedAddress, 3);

        var outsideAddress = new XLSheetPoint(3, 1);
        slice.Set(outsideAddress, 4);

        slice.InsertAreaAndShiftRight(new XLSheetRange(new XLSheetPoint(1, 1), new XLSheetPoint(2, 2)));

        await Assert.That(slice[1, 3]).IsEqualTo(1);
        await Assert.That(slice[1, 5]).IsEqualTo(2);
        await Assert.That(slice[purgedAddress]).IsEqualTo(0);
        await Assert.That(slice[outsideAddress]).IsEqualTo(4);
    }

    [Test]
    public async Task DeleteAreaAndShiftUp_Moves_Area_Cells_Up()
    {
        var slice = new Slice<int>();
        var aboveAddress = new XLSheetPoint(1, 3);
        slice.Set(aboveAddress, 1);
        var firstCorner = new XLSheetPoint(2, 2);
        slice.Set(firstCorner, 2);
        var secondCorner = new XLSheetPoint(4, 5);
        slice.Set(secondCorner, 3);
        var rightAddress = new XLSheetPoint(3, 6);
        slice.Set(rightAddress, 4);
        var belowAddress = new XLSheetPoint(5, 3);
        slice.Set(belowAddress, 5);
        var leftAddress = new XLSheetPoint(3, 1);
        slice.Set(leftAddress, 6);

        var deleteArea = new XLSheetRange(firstCorner, secondCorner);
        slice.DeleteAreaAndShiftUp(deleteArea);
        await Assert.That(slice[firstCorner]).IsEqualTo(0);
        await Assert.That(slice[secondCorner]).IsEqualTo(0);
        await Assert.That(slice[belowAddress.Row - deleteArea.Height, belowAddress.Column]).IsEqualTo(5);
        await Assert.That(slice[aboveAddress]).IsEqualTo(1);
        await Assert.That(slice[rightAddress]).IsEqualTo(4);
        await Assert.That(slice[leftAddress]).IsEqualTo(6);
    }

    [Test]
    public async Task DeleteAreaAndShiftLeft_Moves_Area_Cells_Left()
    {
        var slice = new Slice<int>();
        var leftAddress = new XLSheetPoint(3, 1);
        slice.Set(leftAddress, 1);
        var firstCorner = new XLSheetPoint(2, 2);
        slice.Set(firstCorner, 2);
        var secondCorner = new XLSheetPoint(5, 4);
        slice.Set(secondCorner, 3);
        var belowAddress = new XLSheetPoint(6, 3);
        slice.Set(belowAddress, 4);
        var rightAddress = new XLSheetPoint(3, 5);
        slice.Set(rightAddress, 5);
        var aboveAddress = new XLSheetPoint(1, 3);
        slice.Set(aboveAddress, 6);

        var deleteArea = new XLSheetRange(firstCorner, secondCorner);
        slice.DeleteAreaAndShiftLeft(deleteArea);
        await Assert.That(slice[firstCorner]).IsEqualTo(0);
        await Assert.That(slice[secondCorner]).IsEqualTo(0);
        await Assert.That(slice[rightAddress.Row, rightAddress.Column - deleteArea.Width]).IsEqualTo(5);
        await Assert.That(slice[leftAddress]).IsEqualTo(1);
        await Assert.That(slice[belowAddress]).IsEqualTo(4);
        await Assert.That(slice[aboveAddress]).IsEqualTo(6);
    }
}
