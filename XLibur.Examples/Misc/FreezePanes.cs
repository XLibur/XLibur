using XLibur.Excel;

namespace XLibur.Examples.Misc;

public class FreezePanes : IXLExample
{
    public void Create(string filePath)
    {
        using (var wb = new XLWorkbook())
        {
            // Freeze rows and columns in one shot
            var ws1 = wb.AddWorksheet("Freeze1");
            ws1.Cell(5, 5).SetActive();
            ws1.SheetView.Freeze(3, 3);

            // You can also be more specific on what you want to freeze
            // For example:
            var ws2 = wb.AddWorksheet("FreezeRows");
            ws2.Cell(5, 5).SetActive();
            ws2.SheetView.FreezeRows(3);

            var ws3 = wb.AddWorksheet("FreezeColumns");
            ws3.Cell(5, 5).SetActive();
            ws3.SheetView.FreezeColumns(3);

            // Setting the splits without freezing gives the draggable split bar of View -> Split.
            // In that state the two values are OOXML's split positions in twentieths of a point,
            // not line counts: 900 is three default 15pt rows, 2880 three default 48pt columns.
            var wsSplit = wb.AddWorksheet("Split View");
            wsSplit.Cell(2, 2).SetActive();
            wsSplit.SheetView.SplitRow = 900;
            wsSplit.SheetView.SplitColumn = 2880;

            wb.SaveAs(filePath);
        }
    }
}
