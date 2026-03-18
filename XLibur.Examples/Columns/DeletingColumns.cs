using XLibur.Excel;

namespace XLibur.Examples.Columns;

public class DeletingColumns : IXLExample
{
    public void Create(string filePath)
    {
        var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Deleting Columns");

        ws.Row(1).InsertRowsBelow(2);

        var rng1 = ws.Range("B2:D2");
        var rng2 = ws.Range("F2:G2");
        var rng3 = ws.Range("A1:A3");
        var col1 = ws.Column(1);

        rng1.Style.Fill.BackgroundColor = XLColor.Orange;
        rng2.Style.Fill.BackgroundColor = XLColor.Blue;
        rng3.Style.Fill.BackgroundColor = XLColor.Red;
        col1.Style.Fill.BackgroundColor = XLColor.Black;

        ws.Columns("A,C,E:H").Delete();
        ws.Cell("A2").Value = "OK";
        ws.Cell("B2").Value = "OK";

        workbook.SaveAs(filePath);
    }
}
