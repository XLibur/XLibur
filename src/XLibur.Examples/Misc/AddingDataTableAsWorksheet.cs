using System;
using System.Data;
using System.Linq;
using XLibur.Excel;

namespace XLibur.Examples.Misc;

public class AddingDataTableAsWorksheet : IXLExample
{
    public void Create(string filePath)
    {
        var wb = new XLWorkbook();

        var dataTable = GetTable("Information");

        // Add a DataTable as a worksheet
        wb.Worksheets.Add(dataTable);
        wb.Worksheets.First().Columns().AdjustToContents();

        wb.SaveAs(filePath);
    }

    private static DataTable GetTable(string tableName)
    {
        var table = new DataTable();
        table.TableName = tableName;
        table.Columns.Add("Dosage", typeof(int));
        table.Columns.Add("Drug", typeof(string));
        table.Columns.Add("Patient", typeof(string));
        table.Columns.Add("Date", typeof(DateTime));

        table.Rows.Add(25, "Indocin", "David", new DateTime(2000, 1, 1, 0, 0, 0, DateTimeKind.Unspecified));
        table.Rows.Add(50, "Enebrel", "Sam", new DateTime(2000, 1, 2, 0, 0, 0, DateTimeKind.Unspecified));
        table.Rows.Add(10, "Hydralazine", "Christoff", new DateTime(2000, 1, 3, 0, 0, 0, DateTimeKind.Unspecified));
        table.Rows.Add(21, "Combivent", "Janet", new DateTime(2000, 1, 4, 0, 0, 0, DateTimeKind.Unspecified));
        table.Rows.Add(100, "Dilantin", "Melanie", new DateTime(2000, 1, 5, 0, 0, 0, DateTimeKind.Unspecified));
        return table;
    }
}
