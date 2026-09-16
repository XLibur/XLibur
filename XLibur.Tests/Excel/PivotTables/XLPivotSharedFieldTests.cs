using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using DataFieldElement = DocumentFormat.OpenXml.Spreadsheet.DataField;
using Field = DocumentFormat.OpenXml.Spreadsheet.Field;
using PageField = DocumentFormat.OpenXml.Spreadsheet.PageField;
using PivotField = DocumentFormat.OpenXml.Spreadsheet.PivotField;

namespace XLibur.Tests.Excel.PivotTables;

/// <summary>
/// A pivot field can be used in more than one place at once: on the rows, the columns or the
/// filters, and in the values. Taking it off one of them must leave the settings the others need.
/// </summary>
public class XLPivotSharedFieldTests
{
    [Test]
    [Property("Description", "#555: removing a value took the field off the rows, so the saved file had a row field with no axis")]
    public async Task Removing_a_value_leaves_the_field_on_the_rows()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var pt = CreatePivotTable(wb);
            pt.RowLabels.Add("Name");
            pt.Values.Add("Name", "Count");
            await Assert.That(Describe(pt)).IsEqualTo(new Layout(
                "axisRow:Name:data,-:-:-,-:-:-", Rows: "0", Pages: "", Values: "0"))
                .Because("the field must be on the rows and in the values at once, or this proves nothing");

            pt.Values.Remove("Count");

            await Assert.That(Describe(pt)).IsEqualTo(new Layout(
                "axisRow:Name:-,-:-:-,-:-:-", Rows: "0", Pages: "", Values: ""));
            wb.SaveAs(saved);
        }

        await Assert.That(SavedLayout(saved)).IsEqualTo(new Layout(
            "axisRow:Name:-,-:-:-,-:-:-", Rows: "0", Pages: "", Values: ""));
        await Assert.That(ReloadedLabels(saved)).IsEqualTo("rows=Name filters= values=");
    }

    [Test]
    [Property("Description", "#555: clearing the values took the field off the rows")]
    public async Task Clearing_the_values_leaves_the_field_on_the_rows()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var pt = CreatePivotTable(wb);
            pt.RowLabels.Add("Name");
            pt.Values.Add("Name", "Count");
            await Assert.That(Describe(pt)).IsEqualTo(new Layout(
                "axisRow:Name:data,-:-:-,-:-:-", Rows: "0", Pages: "", Values: "0"))
                .Because("the field must be on the rows and in the values at once, or this proves nothing");

            pt.Values.Clear();

            await Assert.That(Describe(pt)).IsEqualTo(new Layout(
                "axisRow:Name:-,-:-:-,-:-:-", Rows: "0", Pages: "", Values: ""));
            wb.SaveAs(saved);
        }

        await Assert.That(SavedLayout(saved)).IsEqualTo(new Layout(
            "axisRow:Name:-,-:-:-,-:-:-", Rows: "0", Pages: "", Values: ""));
        await Assert.That(ReloadedLabels(saved)).IsEqualTo("rows=Name filters= values=");
    }

    [Test]
    [Property("Description", "#555: removing a report filter cleared dataField on a field that was also a value")]
    public async Task Removing_a_report_filter_leaves_the_field_in_the_values()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var pt = CreatePivotTable(wb);
            pt.ReportFilters.Add("Name");
            pt.Values.Add("Name", "Count");
            await Assert.That(Describe(pt)).IsEqualTo(new Layout(
                "axisPage:Name:data,-:-:-,-:-:-", Rows: "", Pages: "0", Values: "0"))
                .Because("the field must be a filter and a value at once, or this proves nothing");

            pt.ReportFilters.Remove("Name");

            await Assert.That(Describe(pt)).IsEqualTo(new Layout(
                "-:-:data,-:-:-,-:-:-", Rows: "", Pages: "", Values: "0"));
            wb.SaveAs(saved);
        }

        await Assert.That(SavedLayout(saved)).IsEqualTo(new Layout(
            "-:-:data,-:-:-,-:-:-", Rows: "", Pages: "", Values: "0"));
        await Assert.That(ReloadedLabels(saved)).IsEqualTo("rows= filters= values=Count");
    }

    [Test]
    [Property("Description", "#555: clearing the report filters cleared dataField on a field that was also a value")]
    public async Task Clearing_the_report_filters_leaves_the_field_in_the_values()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var pt = CreatePivotTable(wb);
            pt.ReportFilters.Add("Name");
            pt.Values.Add("Name", "Count");
            await Assert.That(Describe(pt)).IsEqualTo(new Layout(
                "axisPage:Name:data,-:-:-,-:-:-", Rows: "", Pages: "0", Values: "0"))
                .Because("the field must be a filter and a value at once, or this proves nothing");

            pt.ReportFilters.Clear();

            await Assert.That(Describe(pt)).IsEqualTo(new Layout(
                "-:-:data,-:-:-,-:-:-", Rows: "", Pages: "", Values: "0"));
            wb.SaveAs(saved);
        }

        await Assert.That(SavedLayout(saved)).IsEqualTo(new Layout(
            "-:-:data,-:-:-,-:-:-", Rows: "", Pages: "", Values: "0"));
        await Assert.That(ReloadedLabels(saved)).IsEqualTo("rows= filters= values=Count");
    }

    [Test]
    [Property("Description", "#555: taking a field off the rows cleared dataField on a field that was also a value")]
    public async Task Removing_a_field_from_the_rows_leaves_it_in_the_values()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var pt = CreatePivotTable(wb);
            pt.RowLabels.Add("Name");
            pt.Values.Add("Name", "Count");
            await Assert.That(Describe(pt)).IsEqualTo(new Layout(
                "axisRow:Name:data,-:-:-,-:-:-", Rows: "0", Pages: "", Values: "0"))
                .Because("the field must be on the rows and in the values at once, or this proves nothing");

            pt.RowLabels.Remove("Name");

            await Assert.That(Describe(pt)).IsEqualTo(new Layout(
                "-:-:data,-:-:-,-:-:-", Rows: "", Pages: "", Values: "0"));
            wb.SaveAs(saved);
        }

        await Assert.That(SavedLayout(saved)).IsEqualTo(new Layout(
            "-:-:data,-:-:-,-:-:-", Rows: "", Pages: "", Values: "0"));
        await Assert.That(ReloadedLabels(saved)).IsEqualTo("rows= filters= values=Count");
    }

    [Test]
    [Property("Description", "#555: clearing the rows cleared dataField on a field that was also a value")]
    public async Task Clearing_the_rows_leaves_the_field_in_the_values()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var pt = CreatePivotTable(wb);
            pt.RowLabels.Add("Name");
            pt.Values.Add("Name", "Count");
            await Assert.That(Describe(pt)).IsEqualTo(new Layout(
                "axisRow:Name:data,-:-:-,-:-:-", Rows: "0", Pages: "", Values: "0"))
                .Because("the field must be on the rows and in the values at once, or this proves nothing");

            pt.RowLabels.Clear();

            await Assert.That(Describe(pt)).IsEqualTo(new Layout(
                "-:-:data,-:-:-,-:-:-", Rows: "", Pages: "", Values: "0"));
            wb.SaveAs(saved);
        }

        await Assert.That(SavedLayout(saved)).IsEqualTo(new Layout(
            "-:-:data,-:-:-,-:-:-", Rows: "", Pages: "", Values: "0"));
        await Assert.That(ReloadedLabels(saved)).IsEqualTo("rows= filters= values=Count");
    }

    [Test]
    [Property("Description", "#555: the same field can back several values, and removing one cleared the flag the others still need")]
    public async Task Removing_one_of_two_values_over_the_same_field_leaves_it_a_data_field()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var pt = CreatePivotTable(wb);
            pt.Values.Add("Sold", "Sum of Sold");
            pt.Values.Add("Sold", "Count of Sold");
            await Assert.That(Describe(pt)).IsEqualTo(new Layout(
                "-:-:-,-:-:-,-:-:data", Rows: "", Pages: "", Values: "2,2"))
                .Because("both values must be over the one field, or this proves nothing");

            pt.Values.Remove("Sum of Sold");

            await Assert.That(Describe(pt)).IsEqualTo(new Layout(
                "-:-:-,-:-:-,-:-:data", Rows: "", Pages: "", Values: "2"));
            wb.SaveAs(saved);
        }

        await Assert.That(SavedLayout(saved)).IsEqualTo(new Layout(
            "-:-:-,-:-:-,-:-:data", Rows: "", Pages: "", Values: "2"));
        await Assert.That(ReloadedLabels(saved)).IsEqualTo("rows= filters= values=Count of Sold");
    }

    [Test]
    [Property("Description", "#555: a field left in no place keeps no name, so the name can be used again")]
    public async Task Removing_the_last_value_clears_a_name_the_field_can_no_longer_use()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb);
        pt.Values.Add("Sold", "Total");

        // Excel saves a name on a pivot field that is only a value, and a load keeps it.
        pt.PivotFields[2].Name = "Total";
        await Assert.That(Describe(pt)).IsEqualTo(new Layout(
            "-:-:-,-:-:-,-:Total:data", Rows: "", Pages: "", Values: "2"))
            .Because("the field must carry a name of its own, or this proves nothing");

        pt.Values.Remove("Total");

        await Assert.That(Describe(pt)).IsEqualTo(new Layout(
            "-:-:-,-:-:-,-:-:-", Rows: "", Pages: "", Values: ""));
        await Assert.That(() => pt.RowLabels.Add("Region", "Total")).ThrowsNothing()
            .Because("the name is free once no field uses it");
    }

    private static XLPivotTable CreatePivotTable(XLWorkbook wb)
    {
        var ws = wb.AddWorksheet("Data");
        var range = ws.Cell("A1").InsertData(new object[]
        {
            ("Name", "Region", "Sold"),
            ("Cake", "North", 7),
            ("Pie", "South", 3),
        });
        return (XLPivotTable)ws.PivotTables.Add("pt", ws.Cell("E1"), range!);
    }

    /// <summary>
    /// The layout of a pivot table. <paramref name="Fields"/> is the axis, the name and the data
    /// flag of each pivot field, as the file writes them (<c>-</c> for none). The other lists are
    /// field indexes.
    /// </summary>
    private sealed record Layout(string Fields, string Rows, string Pages, string Values);

    private static Layout Describe(XLPivotTable pt)
    {
        var fields = string.Join(",", pt.PivotFields.Select(f =>
        {
            var axis = f.Axis?.ToString() is { } a ? char.ToLowerInvariant(a[0]) + a[1..] : "-";
            return $"{axis}:{f.Name ?? "-"}:{(f.DataField ? "data" : "-")}";
        }));
        // XLPivotDataFields carries two enumerable interfaces, so name the one to read.
        IReadOnlyCollection<XLPivotDataField> dataFields = pt.DataFields;
        return new Layout(
            fields,
            string.Join(",", pt.RowAxis.Fields.Select(f => f.Value)),
            string.Join(",", pt.Filters.Fields.Select(f => f.Field)),
            string.Join(",", dataFields.Select(f => f.Field)));
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
            .Select(f =>
            {
                var isData = f.DataField?.Value == true;
                return $"{f.Axis?.InnerText ?? "-"}:{f.Name?.Value ?? "-"}:{(isData ? "data" : "-")}";
            }));
        return new Layout(
            fields,
            string.Join(",", definition.RowFields?.Elements<Field>().Select(f => f.Index!.Value) ?? []),
            string.Join(",", definition.PageFields?.Elements<PageField>().Select(f => f.Field!.Value) ?? []),
            string.Join(",", definition.DataFields?.Elements<DataFieldElement>().Select(f => f.Field!.Value) ?? []));
    }

    private static string ReloadedLabels(Stream package)
    {
        package.Position = 0;
        using var wb = new XLWorkbook(package);
        var pt = wb.Worksheet("Data").PivotTables.Single();
        var rows = string.Join(",", pt.RowLabels.Select(f => f.SourceName));
        var filters = string.Join(",", pt.ReportFilters.Select(f => f.SourceName));
        var values = string.Join(",", pt.Values.Select(v => v.CustomName));
        return $"rows={rows} filters={filters} values={values}";
    }
}
