using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;
using XLibur.Excel.AutoFilters;
using XLibur.Excel.IO;

namespace XLibur.Tests.Excel.PivotTables;

/// <summary>
/// The <c>filters</c> collection holds the label, value, date and top-N filters applied to pivot
/// fields — what Excel offers from a field's dropdown. Dropping it on load silently un-filters
/// the pivot table on the next save, which changes what the workbook shows.
/// </summary>
/// <remarks>
/// Not the report-filter axis: that is <c>pageFields</c>, modelled by
/// <c>XLPivotTable.Filters</c>, and was already supported.
/// </remarks>
public class PivotFilterTests
{
    [Test]
    public async Task PivotFilters_SurviveARoundTrip()
    {
        using var input = new MemoryStream();
        CreateWorkbookWithPivotFilters(input);
        input.Position = 0;

        using var wb = new XLWorkbook(input);
        using var output = new MemoryStream();
        wb.SaveAs(output);

        var written = ReadPivotFilters(output);

        await Assert.That(written.Count).IsEqualTo(2);

        var caption = written[0];
        await Assert.That(caption.Field!.Value).IsEqualTo(0U);
        await Assert.That(caption.Id!.Value).IsEqualTo(1U);
        await Assert.That(caption.Type!.InnerText).IsEqualTo("captionEqual");
        await Assert.That(caption.StringValue1!.Value).IsEqualTo("Cookies");
        await Assert.That(caption.Name!.Value).IsEqualTo("Only cookies");

        var value = written[1];
        await Assert.That(value.Field!.Value).IsEqualTo(1U);
        await Assert.That(value.Id!.Value).IsEqualTo(2U);
        await Assert.That(value.Type!.InnerText).IsEqualTo("valueGreaterThan");
        await Assert.That(value.EvaluationOrder!.Value).IsEqualTo(-1);
    }

    /// <summary>
    /// <c>autoFilter</c> is required by the schema and carries the actual criteria, so a round
    /// trip that kept only the attributes would still lose what the filter does.
    /// </summary>
    [Test]
    public async Task PivotFilterAutoFilter_SurvivesARoundTrip()
    {
        using var input = new MemoryStream();
        CreateWorkbookWithPivotFilters(input);
        input.Position = 0;

        using var wb = new XLWorkbook(input);
        using var output = new MemoryStream();
        wb.SaveAs(output);

        var autoFilter = ReadPivotFilters(output)[0].AutoFilter;
        await Assert.That(autoFilter).IsNotNull();

        var filterColumn = autoFilter!.Elements<FilterColumn>().Single();
        await Assert.That(filterColumn.ColumnId!.Value).IsEqualTo(0U);

        var filterValue = filterColumn.Elements<Filters>().Single().Elements<Filter>().Single();
        await Assert.That(filterValue.Val!.Value).IsEqualTo("Cookies");
    }

    /// <summary>
    /// The top-N filter carries a <c>top10</c> child rather than a value list, so it exercises a
    /// different branch of the preserved sub-tree.
    /// </summary>
    [Test]
    public async Task PivotFilterTop10AutoFilter_SurvivesARoundTrip()
    {
        using var input = new MemoryStream();
        CreateWorkbookWithPivotFilters(input);
        input.Position = 0;

        using var wb = new XLWorkbook(input);
        using var output = new MemoryStream();
        wb.SaveAs(output);

        var autoFilter = ReadPivotFilters(output)[1].AutoFilter!;
        var top10 = autoFilter.Elements<FilterColumn>().Single().Elements<Top10>().Single();

        // top defaults to true, so it is written only when the filter takes the bottom instead.
        // The input spelled it out; omitting it says the same thing.
        await Assert.That(top10.Top?.Value ?? true).IsTrue();
        await Assert.That(top10.Val!.Value).IsEqualTo(3D);
    }

    [Test]
    public async Task LoadedPivotFilters_AreExposedOnThePivotTable()
    {
        using var input = new MemoryStream();
        CreateWorkbookWithPivotFilters(input);
        input.Position = 0;

        using var wb = new XLWorkbook(input);
        var pt = (XLPivotTable)wb.Worksheet("PivotSheet").PivotTables.Single();

        await Assert.That(pt.PivotFilters.Count).IsEqualTo(2);
        await Assert.That(pt.PivotFilters[0].Type).IsEqualTo("captionEqual");
        await Assert.That(pt.PivotFilters[0].StringValue1).IsEqualTo("Cookies");
        await Assert.That(pt.PivotFilters[1].EvaluationOrder).IsEqualTo(-1);

        var top10 = pt.PivotFilters[1].AutoFilter.Columns.Single().Criteria as XLTop10Criteria;
        await Assert.That(top10).IsNotNull();
        await Assert.That(top10!.Value).IsEqualTo(3D);
    }

    /// <summary>
    /// Most pivot tables have no filters, and an empty <c>filters</c> element is not valid, so
    /// nothing must be written for them.
    /// </summary>
    [Test]
    public async Task PivotTableWithoutFilters_WritesNoFiltersElement()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet("Data");
        var range = data.FirstCell().InsertData(new object[]
        {
            ("Pastry", "Sold"),
            ("Waffle", 3),
            ("Donut", 5),
        });

        var pivots = wb.AddWorksheet("Pivots");
        var pt = pivots.PivotTables.Add("pvt", pivots.Cell("A1"), range);
        pt.RowLabels.Add("Pastry");
        pt.Values.Add("Sold");

        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var definition = doc.WorkbookPart!.WorksheetParts
            .SelectMany(wsp => wsp.GetPartsOfType<PivotTablePart>())
            .Single()
            .PivotTableDefinition;

        await Assert.That(definition.Elements<PivotFilters>().Any()).IsFalse();
    }

    private static List<PivotFilter> ReadPivotFilters(MemoryStream saved) =>
        PivotFilterWorkbook.ReadFilters(saved);

    /// <summary>
    /// A minimal workbook whose pivot table carries two filters: a caption filter holding a value
    /// list, and a value filter holding a top-N.
    /// </summary>
    private static void CreateWorkbookWithPivotFilters(Stream stream) =>
        PivotFilterWorkbook.Create(stream, BuildCaptionFilter(), BuildValueFilter());

    /// <summary>
    /// A caption filter whose <c>autoFilter</c> holds an explicit value list.
    /// </summary>
    private static PivotFilter BuildCaptionFilter()
    {
        var filterColumn = new FilterColumn { ColumnId = 0U };
        var filterValues = new Filters();
        filterValues.Append(new Filter { Val = "Cookies" });
        filterColumn.Append(filterValues);

        var autoFilter = new AutoFilter { Reference = "A1:A3" };
        autoFilter.Append(filterColumn);

        var pivotFilter = new PivotFilter
        {
            Field = 0U,
            Id = 1U,
            Type = PivotFilterValues.CaptionEqual,
            StringValue1 = "Cookies",
            Name = "Only cookies",
        };
        pivotFilter.Append(autoFilter);
        return pivotFilter;
    }

    /// <summary>
    /// A value filter whose <c>autoFilter</c> holds a top-N instead, and a non-default
    /// <c>evalOrder</c>.
    /// </summary>
    private static PivotFilter BuildValueFilter()
    {
        var filterColumn = new FilterColumn { ColumnId = 0U };
        filterColumn.Append(new Top10 { Top = true, Val = 3D });

        var autoFilter = new AutoFilter { Reference = "B1:B3" };
        autoFilter.Append(filterColumn);

        var pivotFilter = new PivotFilter
        {
            Field = 1U,
            Id = 2U,
            Type = PivotFilterValues.ValueGreaterThan,
            EvaluationOrder = -1,
        };
        pivotFilter.Append(autoFilter);
        return pivotFilter;
    }
}
