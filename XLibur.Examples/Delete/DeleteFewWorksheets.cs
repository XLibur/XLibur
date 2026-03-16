using System.IO;
using XLibur.Excel;

namespace XLibur.Examples.Delete;

public class DeleteFewWorksheets : IXLExample
{
    public void Create(string filePath)
    {
        var tempFile = ExampleHelper.GetTempFilePath(filePath);
        try
        {
            Prepare(tempFile);
            Delete(filePath, tempFile);
        }
        finally
        {
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }
        }
    }

    private static void Delete(string filePath, string tempFile)
    {
        var workbook = new XLWorkbook(tempFile);
        workbook.Worksheets.Delete("1");
        workbook.Worksheets.Delete("2");
        workbook.SaveAs(filePath);
    }

    private static void Prepare(string tempFile)
    {
        var workbook = new XLWorkbook();
        workbook.Worksheets.Add("1");
        workbook.Worksheets.Add("2");
        workbook.Worksheets.Add("3");
        workbook.Worksheets.Add("4");
        workbook.SaveAs(tempFile);
    }
}
