using System.Linq;
using XLibur.Excel;

namespace XLibur.Examples.Delete;

public class DeleteRows : IXLExample
{
    public void Create(string filePath)
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Delete red rows");

        // Put a value in a few cells
        foreach (var r in Enumerable.Range(1, 5))
        {
            foreach (var c in Enumerable.Range(1, 5))
                ws.Cell(r, c).Value = $"R{r}C{c}";
        }

        var blueRow = ws.Rows(1, 2);
        var redRow = ws.Row(5);

        blueRow.Style.Fill.BackgroundColor = XLColor.Blue;

        redRow.Style.Fill.BackgroundColor = XLColor.Red;
        workbook.SaveAs(filePath);

        using var workbook2 = new XLWorkbook(filePath);
        var ws2 = workbook2.Worksheets.Worksheet("Delete red rows");

        ws2.Rows(1, 2).Delete();
        workbook2.Save();
    }
}