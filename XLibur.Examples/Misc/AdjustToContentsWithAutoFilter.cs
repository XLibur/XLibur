using XLibur.Excel;

namespace XLibur.Examples.Misc;

public class AdjustToContentsWithAutoFilter : IXLExample
{
    public void Create(string filePath)
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("AutoFilter");
        ws.Cell("A1").Value = "AVeryLongColumnHeader";
        ws.Cell("A2").Value = "John";
        ws.Cell("A3").Value = "Hank";
        ws.Cell("A4").Value = "Dagny";

        ws.RangeUsed().SetAutoFilter();

        ws.Columns().AdjustToContents();

        wb.SaveAs(filePath);
    }
}
