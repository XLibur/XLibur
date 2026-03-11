using ClosedXML.Excel;

namespace ClosedXML.Examples.Misc;

public class WorkbookProtection : IXLExample
{
    #region Methods

    // Public
    public void Create(string filePath)
    {
        using var wb = new XLWorkbook();
        wb.Worksheets.Add("Workbook Protection");
        wb.Protect("Abc@123");
        wb.SaveAs(filePath);
    }

    #endregion Methods
}
