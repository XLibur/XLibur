using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using Field = DocumentFormat.OpenXml.Spreadsheet.Field;
using PageField = DocumentFormat.OpenXml.Spreadsheet.PageField;
using PivotField = DocumentFormat.OpenXml.Spreadsheet.PivotField;

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

        var pt = ws.PivotTables.Add("pt", ws.Cell("E1"), data!);
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

        var pt = ws.PivotTables.Add("pt", ws.Cell("E2"), data!);

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

    [Test]
    [Property("Description", "#551: Remove passed the filter's position in the report filters as a field index, so the field on the rows at that index lost its axis")]
    public async Task Removing_a_report_filter_takes_its_own_field_off_the_page_axis_not_a_field_on_the_rows()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var pt = CreatePivotTable(wb);
            pt.RowLabels.Add("Name");
            pt.ReportFilters.Add("Region");
            await Assert.That(Describe(pt)).IsEqualTo(new Layout(
                "axisRow:Name,-:-,axisPage:Region", Rows: "0", Columns: "", Pages: "2", "E3", RowPageCount: 1))
                .Because("the filter's position 0 must be a field on the rows, or this proves nothing");

            pt.ReportFilters.Remove("Region");

            await Assert.That(Describe(pt)).IsEqualTo(new Layout(
                "axisRow:Name,-:-,-:-", Rows: "0", Columns: "", Pages: "", "E1", RowPageCount: 0));
            wb.SaveAs(saved);
        }

        await Assert.That(SavedLayout(saved)).IsEqualTo(new Layout(
            "axisRow:Name,-:-,-:-", Rows: "0", Columns: "", Pages: "", "E1", RowPageCount: 0));
        await Assert.That(ReloadedLabels(saved)).IsEqualTo("rows=Name filters=");
    }

    [Test]
    [Property("Description", "#551: removing the second of two report filters took the first filter's field off the page axis and left its own")]
    public async Task Removing_the_second_of_two_report_filters_leaves_the_first_on_the_page_axis()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var pt = CreatePivotTable(wb);
            pt.RowLabels.Add("Name");
            pt.ReportFilters.Add("Month");
            pt.ReportFilters.Add("Region");
            await Assert.That(Describe(pt)).IsEqualTo(new Layout(
                "axisRow:Name,axisPage:Month,axisPage:Region", Rows: "0", Columns: "", Pages: "1,2", "E4", RowPageCount: 2))
                .Because("the second filter's position 1 must be the first filter's field, or this proves nothing");

            pt.ReportFilters.Remove("Region");

            await Assert.That(Describe(pt)).IsEqualTo(new Layout(
                "axisRow:Name,axisPage:Month,-:-", Rows: "0", Columns: "", Pages: "1", "E3", RowPageCount: 1));
            wb.SaveAs(saved);
        }

        await Assert.That(SavedLayout(saved)).IsEqualTo(new Layout(
            "axisRow:Name,axisPage:Month,-:-", Rows: "0", Columns: "", Pages: "1", "E3", RowPageCount: 1));
        await Assert.That(ReloadedLabels(saved)).IsEqualTo("rows=Name filters=Month");
    }

    [Test]
    [Property("Description", "#551: a report filter at a position equal to the index of a field on the columns took that field off the columns")]
    public async Task Removing_a_report_filter_leaves_the_field_on_the_columns_whose_index_is_the_filters_position()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var pt = CreatePivotTable(wb);
            pt.ColumnLabels.Add("Month");
            pt.ReportFilters.Add("Name");
            pt.ReportFilters.Add("Region");
            await Assert.That(Describe(pt)).IsEqualTo(new Layout(
                "axisPage:Name,axisCol:Month,axisPage:Region", Rows: "", Columns: "1", Pages: "0,2", "E4", RowPageCount: 2))
                .Because("the second filter's position 1 must be the field on the columns, or this proves nothing");

            pt.ReportFilters.Remove("Region");

            await Assert.That(Describe(pt)).IsEqualTo(new Layout(
                "axisPage:Name,axisCol:Month,-:-", Rows: "", Columns: "1", Pages: "0", "E3", RowPageCount: 1));

            pt.ReportFilters.Remove("Name");

            await Assert.That(Describe(pt)).IsEqualTo(new Layout(
                "-:-,axisCol:Month,-:-", Rows: "", Columns: "1", Pages: "", "E1", RowPageCount: 0))
                .Because("a filter whose position is its own field index is taken off, as before");
            wb.SaveAs(saved);
        }

        await Assert.That(SavedLayout(saved)).IsEqualTo(new Layout(
            "-:-,axisCol:Month,-:-", Rows: "", Columns: "1", Pages: "", "E1", RowPageCount: 0));
        await Assert.That(ReloadedLabels(saved)).IsEqualTo("rows= columns=Month filters=");
    }

    private static XLPivotTable CreatePivotTable(XLWorkbook wb)
    {
        var ws = wb.AddWorksheet("Data");
        var range = ws.Cell("A1").InsertData(new object[]
        {
            ("Name", "Month", "Region"),
            ("Cake", "Jan", "North"),
            ("Pie", "Feb", "South"),
        });
        return (XLPivotTable)ws.PivotTables.Add("pt", ws.Cell("E1"), range!);
    }

    /// <summary>
    /// The layout of a pivot table. <paramref name="Fields"/> is the axis and the name of each pivot
    /// field, as the file writes them (<c>-</c> for none). The other lists are field indexes.
    /// </summary>
    private sealed record Layout(string Fields, string Rows, string Columns, string Pages, string Location, int RowPageCount);

    private static Layout Describe(XLPivotTable pt)
    {
        var fields = string.Join(",", pt.PivotFields.Select(f =>
        {
            var axis = f.Axis?.ToString() is { } a ? char.ToLowerInvariant(a[0]) + a[1..] : "-";
            return $"{axis}:{f.Name ?? "-"}";
        }));
        return new Layout(
            fields,
            string.Join(",", pt.RowAxis.Fields.Select(f => f.Value)),
            string.Join(",", pt.ColumnAxis.Fields.Select(f => f.Value)),
            string.Join(",", pt.Filters.Fields.Select(f => f.Field)),
            pt.Area.ToString(),
            pt.Filters.GetSize().Height);
    }

    private static Layout SavedLayout(Stream package)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        var definition = document.WorkbookPart!.WorksheetParts
            .SelectMany(part => part.PivotTableParts)
            .Select(part => part.PivotTableDefinition!)
            .Single();
        var fields = string.Join(",", definition.PivotFields!.Elements<PivotField>()
            .Select(f => $"{f.Axis?.InnerText ?? "-"}:{f.Name?.Value ?? "-"}"));
        return new Layout(
            fields,
            string.Join(",", definition.RowFields?.Elements<Field>().Select(f => f.Index!.Value) ?? []),
            string.Join(",", definition.ColumnFields?.Elements<Field>().Select(f => f.Index!.Value) ?? []),
            string.Join(",", definition.PageFields?.Elements<PageField>().Select(f => f.Field!.Value) ?? []),
            definition.Location!.Reference!.Value!,
            (int)(definition.Location.RowPageCount?.Value ?? 0));
    }

    private static string ReloadedLabels(Stream package)
    {
        package.Position = 0;
        using var wb = new XLWorkbook(package);
        var pt = wb.Worksheet("Data").PivotTables.Single();
        var rows = $"rows={string.Join(",", pt.RowLabels.Select(f => f.SourceName))}";
        var columns = pt.ColumnLabels.Any() ? $" columns={string.Join(",", pt.ColumnLabels.Select(f => f.SourceName))}" : "";
        return $"{rows}{columns} filters={string.Join(",", pt.ReportFilters.Select(f => f.SourceName))}";
    }
}
