using System;
using XLibur.Excel;

namespace XLibur.Examples.AutoFilters;

public class DateTimeGroupAutoFilter : IXLExample
{
    public void Create(string filePath)
    {
        using var wb = new XLWorkbook();

        #region Single Column Dates

        var singleColumnDates = "Single Column Dates";
        var ws = wb.Worksheets.Add(singleColumnDates);

        // Add a bunch of dates to filter
        ws.Cell("A1").SetValue("Dates")
            .CellBelow().SetValue(new DateTime(2018, 1, 1, 0, 0, 0, DateTimeKind.Unspecified).AddDays(2))
            .CellBelow().SetValue(new DateTime(2018, 1, 1, 0, 0, 0, DateTimeKind.Unspecified).AddDays(3))
            .CellBelow().SetValue(new DateTime(2018, 1, 1, 0, 0, 0, DateTimeKind.Unspecified).AddDays(3))
            .CellBelow().SetValue(new DateTime(2018, 1, 1, 0, 0, 0, DateTimeKind.Unspecified).AddDays(5))
            .CellBelow().SetValue(new DateTime(2018, 1, 1, 0, 0, 0, DateTimeKind.Unspecified).AddDays(1))
            .CellBelow().SetValue(new DateTime(2018, 1, 1, 0, 0, 0, DateTimeKind.Unspecified).AddDays(4));

        ws.Column(1).Style.NumberFormat.Format = "d MMMM yyyy";

        // Add filters
        ws.RangeUsed().SetAutoFilter().Column(1).AddDateGroupFilter(new DateTime(2018, 1, 1, 0, 0, 0, DateTimeKind.Unspecified).AddDays(3), XLDateTimeGrouping.Day);

        // Sort the filtered list
        ws.AutoFilter.Sort();

        #endregion Single Column Dates

        ws.Columns().AdjustToContents();
        wb.SaveAs(filePath);
    }
}
