using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.PivotTables;

public class XLPivotTableFiltersTests
{
    [Test]
    [Property("Description", "https://github.com/ClosedXML/ClosedXML/issues/2486")]
    public async Task AddSelectedValue_allows_value_not_present_in_data()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var data = ws.Cell("A1").InsertData(new object[]
        {
            ("Col1", "Col2"),
            ("A", false),
            ("B", false),
        });

        var pt = ws.PivotTables.Add("pt", ws.Cell("E1"), data);
        pt.RowLabels.Add("Col1");
        var filter = pt.ReportFilters.Add("Col2");

        // true is not among the data values, but should still be allowed as a filter selection
        await Assert.That(() => filter.AddSelectedValue(true)).ThrowsNothing();
    }

    [Test]
    public async Task Adding_and_removing_filters_shifts_pivot_table_area()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var data = ws.Cell("A1").InsertData(new object[]
        {
            ("Name", "City", "Flavor", "Value"),
            ("Cake", "Tokyo", "Vanilla", 7),
        });

        var pt = ws.PivotTables.Add("pt", ws.Cell("E2"), data);

        // No filter, the table is at the original cell
        await Assert.That(((XLPivotTable)pt).Area.ToString()).IsEqualTo("E2");

        pt.ReportFilters.Add("City");

        // First filter also adds divider row between filter and the table.
        await Assert.That(((XLPivotTable)pt).Area.ToString()).IsEqualTo("E4");

        pt.ReportFilters.Add("Flavor");

        // When second filter is added, there is no need to add second divider row.
        await Assert.That(((XLPivotTable)pt).Area.ToString()).IsEqualTo("E5");

        pt.ReportFilters.Remove("City");
        await Assert.That(((XLPivotTable)pt).Area.ToString()).IsEqualTo("E4");

        pt.ReportFilters.Remove("Flavor");
        await Assert.That(((XLPivotTable)pt).Area.ToString()).IsEqualTo("E2");
    }
}
