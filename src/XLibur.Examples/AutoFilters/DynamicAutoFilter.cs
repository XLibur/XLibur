using System.IO;
using XLibur.Excel;

namespace XLibur.Examples.AutoFilters;

public class DynamicAutoFilter : IXLExample
{
    public void Create(string filePath)
    {
        var wb = new XLWorkbook();

        #region Single Column Numbers
        var singleColumnNumbers = "Single Column Numbers";
        var ws = wb.Worksheets.Add(singleColumnNumbers);

        // Add a bunch of numbers to filter
        ws.Cell("A1").SetValue("Numbers")
            .CellBelow().SetValue(2)
            .CellBelow().SetValue(3)
            .CellBelow().SetValue(3)
            .CellBelow().SetValue(5)
            .CellBelow().SetValue(1)
            .CellBelow().SetValue(4);

        // Add filters
        ws.RangeUsed()!.SetAutoFilter().Column(1).AboveAverage();

        #endregion

        #region Multi Column
        var multiColumn = "Multi Column";
        ws = wb.Worksheets.Add(multiColumn);

        ws.Cell("A1").SetValue("First")
            .CellBelow().SetValue("B")
            .CellBelow().SetValue("C")
            .CellBelow().SetValue("C")
            .CellBelow().SetValue("E")
            .CellBelow().SetValue("A")
            .CellBelow().SetValue("D");

        ws.Cell("B1").SetValue("Numbers")
            .CellBelow().SetValue(2)
            .CellBelow().SetValue(3)
            .CellBelow().SetValue(3)
            .CellBelow().SetValue(5)
            .CellBelow().SetValue(1)
            .CellBelow().SetValue(4);

        ws.Cell("C1").SetValue("Strings")
            .CellBelow().SetValue("B")
            .CellBelow().SetValue("C")
            .CellBelow().SetValue("C")
            .CellBelow().SetValue("E")
            .CellBelow().SetValue("A")
            .CellBelow().SetValue("D");

        // Add filters
        ws.RangeUsed()!.SetAutoFilter().Column(2).BelowAverage();

        #endregion

        using (var ms = new MemoryStream())
        {
            wb.SaveAs(ms);

            var workbook = new XLWorkbook(ms);

            #region Single Column Numbers
            #endregion

            #region Multi Column
            #endregion

            workbook.SaveAs(filePath);
            ms.Close();
        }
    }
}
