using System;
using XLibur.Excel;

namespace XLibur.Examples.Misc;

public class WorkbookProperties : IXLExample
{
    public void Create(string filePath)
    {
        var wb = new XLWorkbook();
        wb.Worksheets.Add("Workbook Properties");

        wb.Properties.Author = "theAuthor";
        wb.Properties.Title = "theTitle";
        wb.Properties.Subject = "theSubject";
        wb.Properties.Category = "theCategory";
        wb.Properties.Keywords = "theKeywords";
        wb.Properties.Comments = "theComments";
        wb.Properties.Status = "theStatus";
        wb.Properties.LastModifiedBy = "theLastModifiedBy";
        wb.Properties.Company = "theCompany";
        wb.Properties.Manager = "theManager";

        // Creating/Using custom properties
        wb.CustomProperties.Add("theText", "XXX");
        wb.CustomProperties.Add("theDate",
            new DateTime(2011, 1, 1, 17, 0, 0,
                DateTimeKind.Utc)); // Use UTC to make sure the test can be run in any time zone
        wb.CustomProperties.Add("theNumber", 123.456);
        wb.CustomProperties.Add("theBoolean", true);

        wb.SaveAs(filePath);
    }
}
