using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Ranges;

/// <summary>
/// Rows and columns are renumbered by one shared shift, <c>XLLineCollection.ShiftLines</c>
/// (#621 finding 18). Columns used to sort their keys on every insert; these pin that both axes
/// carry each materialised line to its new number and drop the ones pushed off the sheet.
/// </summary>
public class LineCollectionShiftTests
{
    [Test]
    public async Task InsertColumnsBefore_CarriesEveryMaterialisedColumnAlong()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Column(2).Width = 12;
        ws.Column(3).Width = 13;
        ws.Column(5).Width = 15;

        ws.Column(2).InsertColumnsBefore(2);

        await Assert.That(ws.Column(4).Width).IsEqualTo(12);
        await Assert.That(ws.Column(5).Width).IsEqualTo(13);
        await Assert.That(ws.Column(7).Width).IsEqualTo(15);
        await Assert.That(ws.Column(4).ColumnNumber()).IsEqualTo(4);
    }

    [Test]
    public async Task InsertRowsAbove_CarriesEveryMaterialisedRowAlong()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Row(2).Height = 22;
        ws.Row(3).Height = 23;
        ws.Row(5).Height = 25;

        ws.Row(2).InsertRowsAbove(2);

        await Assert.That(ws.Row(4).Height).IsEqualTo(22);
        await Assert.That(ws.Row(5).Height).IsEqualTo(23);
        await Assert.That(ws.Row(7).Height).IsEqualTo(25);
        await Assert.That(ws.Row(4).RowNumber()).IsEqualTo(4);
    }

    [Test]
    public async Task InsertColumnsBefore_DropsColumnsPushedPastTheLastColumn()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Column(XLHelper.MaxColumnNumber - 1).Width = 20;
        ws.Column(XLHelper.MaxColumnNumber).Width = 30;

        ws.Column(1).InsertColumnsBefore(1);

        var columns = ((XLWorksheet)ws).Internals.ColumnsCollection;
        await Assert.That(columns.Keys.Max()).IsEqualTo(XLHelper.MaxColumnNumber);
        await Assert.That(columns[XLHelper.MaxColumnNumber].Width).IsEqualTo(20);
    }
}
