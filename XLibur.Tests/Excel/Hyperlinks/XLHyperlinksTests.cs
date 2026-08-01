using System;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Hyperlinks;

public class XLHyperlinksTests
{
    [Test]
    [MethodDataSource(nameof(StructuralChangeCases))]
    public async Task Hyperlink_is_moved_on_sheet_structure_change(string hyperlinkPosition, Action<IXLWorksheet> structuralChange, string expectedPosition)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var hyperlink = new XLHyperlink("https://example.com");
        ws.Cell(hyperlinkPosition).SetHyperlink(hyperlink);

        structuralChange(ws);

        await Assert.That(ws.Cell(hyperlinkPosition).HasHyperlink).IsFalse();
        await Assert.That(hyperlink).IsSameReferenceAs(ws.Cell(expectedPosition).GetHyperlink());
    }

    public static IEnumerable<object[]> StructuralChangeCases
    {
        get
        {
            return new List<(string, Action<IXLWorksheet>, string)>
            {
                ("D5", ws => ws.Range("A5:B5").Delete(XLShiftDeletedCells.ShiftCellsLeft), "B5"),
                ("D5", ws => ws.Range("B2:D4").Delete(XLShiftDeletedCells.ShiftCellsUp), "D2"),
                ("D5", ws => ws.Column("D").InsertColumnsBefore(2), "F5"), // Insert column leftward
                ("D5", ws => ws.Row(2).InsertRowsAbove(4), "D9"), // Insert row above
            }.Select(x => new object[] { x.Item1, x.Item2, x.Item3 });
        }
    }
}
