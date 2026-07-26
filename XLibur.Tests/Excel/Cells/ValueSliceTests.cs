using XLibur.Excel;
using XLibur.Excel.Coordinates;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Cells;

public class ValueSliceTests
{
    [Test]
    public async Task Deleting_worksheet_dereferences_all_texts_in_its_value_slice()
    {
        using var wb = new XLWorkbook();
        var sst = wb.SharedStringTable;
        var keptWs = wb.AddWorksheet();
        var removedWs = wb.AddWorksheet();
        keptWs.Cell("A1").Value = "Double referenced text";
        removedWs.Cell("A1").Value = "Double referenced text";
        removedWs.Cell("B1").Value = "Single referenced text";

        await Assert.That(sst.Count).IsEqualTo(2);

        wb.Worksheets.Delete(removedWs.Name);

        await Assert.That(sst.Count).IsEqualTo(1);
        await Assert.That(keptWs.Cell(1, 1).Value).IsEqualTo("Double referenced text");
    }

    [Test]
    public async Task Clear_dereferences_texts_in_the_range()
    {
        using var wb = new XLWorkbook();
        var sst = wb.SharedStringTable;
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "Double referenced text";
        ws.Cell("B2").Value = "Double referenced text";
        ws.Cell("C2").Value = "Single referenced text";

        await Assert.That(sst.Count).IsEqualTo(2);
        ((XLWorksheet)ws).Internals.CellsCollection.ValueSlice.Clear(new Area(2, 2, 2, 3));
        await Assert.That(sst.Count).IsEqualTo(1);
        await Assert.That(ws.Cell("A1").Value).IsEqualTo("Double referenced text");
    }

    [Test]
    public async Task DeleteAreaAndShiftLeft_dereferences_all_texts_deleted_area()
    {
        using var wb = new XLWorkbook();
        var sst = wb.SharedStringTable;
        var ws = wb.AddWorksheet();
        ws.Cell("B2").Value = "Deleted Single Reference"; // id 0
        ws.Cell("C1").Value = "Kept Single Reference"; // id 1
        ws.Cell("B1").Value = "Kept Double Reference"; // id 2
        ws.Cell("C3").Value = "Kept Double Reference"; // id 2

        ((XLWorksheet)ws).Internals.CellsCollection.ValueSlice.DeleteAreaAndShiftLeft(new Area(2, 2, 3, 3));

        await Assert.That(sst.Count).IsEqualTo(2);
        await Assert.That(sst[1]).IsEqualTo("Kept Single Reference");
        await Assert.That(sst[2]).IsEqualTo("Kept Double Reference");
    }

    [Test]
    public async Task DeleteAreaAndShiftUp_dereferences_all_texts_deleted_area()
    {
        using var wb = new XLWorkbook();
        var sst = wb.SharedStringTable;
        var ws = wb.AddWorksheet();
        ws.Cell("B2").Value = "Deleted Single Reference"; // id 0
        ws.Cell("A3").Value = "Kept Single Reference"; // id 1
        ws.Cell("A2").Value = "Kept Double Reference"; // id 2
        ws.Cell("C3").Value = "Kept Double Reference"; // id 2

        ((XLWorksheet)ws).Internals.CellsCollection.ValueSlice.DeleteAreaAndShiftLeft(new Area(2, 2, 3, 3));

        await Assert.That(sst.Count).IsEqualTo(2);
        await Assert.That(sst[1]).IsEqualTo("Kept Single Reference");
        await Assert.That(sst[2]).IsEqualTo("Kept Double Reference");
    }

    [Test]
    public async Task InsertAreaAndShiftDown_dereferences_all_texts_in_pushed_out_range()
    {
        using var wb = new XLWorkbook();
        var sst = wb.SharedStringTable;
        var ws = wb.AddWorksheet();

        ws.Cell("B2").Value = "Kept Single Reference"; // id 0
        ws.Cell("C1048576").Value = "Deleted Single Reference"; // id 1
        ws.Cell("B1048574").Value = "Kept Double Reference"; // id 2
        ws.Cell("B1048576").Value = "Kept Double Reference"; // id 2
        ((XLWorksheet)ws).Internals.CellsCollection.ValueSlice.InsertAreaAndShiftDown(new Area(3, 2, 4, 3));

        await Assert.That(sst.Count).IsEqualTo(2);
        await Assert.That(sst[0]).IsEqualTo("Kept Single Reference");
        await Assert.That(sst[2]).IsEqualTo("Kept Double Reference");
    }

    [Test]
    public async Task InsertAreaAndShiftRight_dereferences_all_texts_in_pushed_out_range()
    {
        using var wb = new XLWorkbook();
        var sst = wb.SharedStringTable;
        var ws = wb.AddWorksheet();

        ws.Cell("B2").Value = "Kept Single Reference"; // id 0
        ws.Cell("XFD2").Value = "Deleted Single Reference"; // id 1
        ws.Cell("XFD3").Value = "Kept Double Reference"; // id 2
        ws.Cell("XFB3").Value = "Kept Double Reference"; // id 2
        ((XLWorksheet)ws).Internals.CellsCollection.ValueSlice.InsertAreaAndShiftRight(new Area(2, 3, 3, 4));

        await Assert.That(sst.Count).IsEqualTo(2);
        await Assert.That(sst[0]).IsEqualTo("Kept Single Reference");
        await Assert.That(sst[2]).IsEqualTo("Kept Double Reference");
    }
}
