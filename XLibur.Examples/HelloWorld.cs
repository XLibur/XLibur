using XLibur.Excel;

namespace XLibur.Examples;

public static class HelloWorld
{
    public static void Create(string filePath)
    {
        var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Sample Sheet");
        worksheet.Cell("A1").Value = "Hello World!";
        workbook.SaveAs(filePath);
    }
}
