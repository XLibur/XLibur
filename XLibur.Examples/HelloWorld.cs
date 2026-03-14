using XLibur.Excel;

namespace XLibur.Examples;

public class HelloWorld
{
    public void Create(string filePath)
    {
        var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Sample Sheet");
        worksheet.Cell("A1").Value = "Hello World!";
        workbook.SaveAs(filePath);
    }
}
