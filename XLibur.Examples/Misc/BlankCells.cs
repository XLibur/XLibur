using XLibur.Excel;


namespace XLibur.Examples.Misc;

public class BlankCells : IXLExample
{
    public void Create(string filePath)
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell(1, 1).Value = "X";
        ws.Cell(1, 1).Clear();
        wb.SaveAs(filePath);
    }
}
