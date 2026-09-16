using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;

namespace XLibur.Tests.Excel.PivotTables;

/// <summary>
/// Where the report filters sit relative to the table. The filter area is above
/// <see cref="XLPivotTable.Area"/> with a one-row gap, and <c>TargetCell</c> is the first cell of
/// that filter area, so the two properties that change how much room the filters need have to move
/// the area to keep the target cell where it is (#571).
/// </summary>
public class XLPivotTableFilterAreaTests
{
    [Test]
    [Property("Description", "#571: FilterFieldsPageWrap changed how much room the filters need without moving the table, so the filter area ran into the table")]
    public async Task Setting_the_page_wrap_on_a_table_that_has_filters_moves_the_table_down_to_fit_them()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb, "E1");
        pt.FilterAreaOrder = XLFilterAreaOrder.OverThenDown;
        AddFilters(pt, 3);

        // Three filters across one row, plus the divider row below them.
        await Assert.That(Describe(pt)).IsEqualTo("target=E1 area=E3 filters=3x1");

        pt.FilterFieldsPageWrap = 1;

        // One filter per row now: three rows of filters and the divider row below them.
        await Assert.That(Describe(pt)).IsEqualTo("target=E1 area=E5 filters=1x3");
    }

    [Test]
    [Property("Description", "#571: FilterAreaOrder changed how much room the filters need without moving the table")]
    public async Task Setting_the_filter_area_order_on_a_table_that_has_filters_moves_the_table_to_fit_them()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb, "E1");
        AddFilters(pt, 3);

        // The default order runs the filters down the sheet: three rows plus the divider row.
        await Assert.That(Describe(pt)).IsEqualTo("target=E1 area=E5 filters=1x3");

        pt.FilterAreaOrder = XLFilterAreaOrder.OverThenDown;

        // Across the sheet the same three filters need one row, so the table comes back up.
        await Assert.That(Describe(pt)).IsEqualTo("target=E1 area=E3 filters=3x1");
    }

    [Test]
    [Property("Description", "#571: a wrap that needs fewer rows than the filters used takes the table back up, keeping the target cell where it is")]
    public async Task A_page_wrap_that_needs_fewer_rows_takes_the_table_back_up()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb, "E4");
        AddFilters(pt, 3);

        await Assert.That(Describe(pt)).IsEqualTo("target=E4 area=E8 filters=1x3");

        // Two filters down the first column and the third in the next one: two rows, not three.
        pt.FilterFieldsPageWrap = 2;

        await Assert.That(Describe(pt)).IsEqualTo("target=E4 area=E7 filters=2x2");
    }

    [Test]
    [Arguments(true)]
    [Arguments(false)]
    [Property("Description", "#571: setting both properties before any filter was added was the workaround, and it must still give the layout it did")]
    public async Task The_layout_is_the_same_whether_the_properties_are_set_before_or_after_the_filters(bool beforeFilters)
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb, "E1");

        if (beforeFilters)
        {
            pt.FilterAreaOrder = XLFilterAreaOrder.OverThenDown;
            pt.FilterFieldsPageWrap = 2;
            AddFilters(pt, 3);
        }
        else
        {
            AddFilters(pt, 3);
            pt.FilterAreaOrder = XLFilterAreaOrder.OverThenDown;
            pt.FilterFieldsPageWrap = 2;
        }

        // Two filters across the first row and the third on the next, plus the divider row.
        await Assert.That(Describe(pt)).IsEqualTo("target=E1 area=E4 filters=2x2");
    }

    [Test]
    [Property("Description", "#571: setting a property twice must not move the table twice, and setting it to the value it already has must not move it at all")]
    public async Task Setting_a_property_to_the_value_it_already_has_leaves_the_table_where_it_is()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb, "E1");
        AddFilters(pt, 3);
        pt.FilterFieldsPageWrap = 2;

        await Assert.That(Describe(pt)).IsEqualTo("target=E1 area=E4 filters=2x2");

        pt.FilterFieldsPageWrap = 2;
        pt.FilterAreaOrder = XLFilterAreaOrder.DownThenOver;

        await Assert.That(Describe(pt)).IsEqualTo("target=E1 area=E4 filters=2x2");
    }

    [Test]
    [Property("Description", "#571: with no filter there is nothing to make room for, so neither property may move the table")]
    public async Task Neither_property_moves_a_table_that_has_no_filters()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb, "E4");

        pt.FilterFieldsPageWrap = 3;
        pt.FilterAreaOrder = XLFilterAreaOrder.OverThenDown;

        await Assert.That(Describe(pt)).IsEqualTo("target=E4 area=E4 filters=0x0");
    }

    [Test]
    [Property("Description", "#571: the saved file must leave room above the table's own ref for the filter area the wrap asks for")]
    public async Task A_page_wrap_set_after_the_filters_saves_a_geometry_whose_filters_fit_above_the_table()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var pt = CreatePivotTable(wb, "E2");
            pt.RowLabels.Add("Name");
            pt.Values.Add("Value");
            pt.FilterAreaOrder = XLFilterAreaOrder.OverThenDown;
            AddFilters(pt, 3);

            pt.FilterFieldsPageWrap = 1;

            await Assert.That(Describe(pt)).IsEqualTo("target=E2 area=E6 filters=1x3");
            wb.SaveAs(saved);
        }

        // Three rows of filters and the divider row above ref, so E2 to E4 hold the filters, E5 is
        // the gap and the table starts at E6. Nothing overlaps.
        await Assert.That(SavedFilterArea(saved))
            .IsEqualTo("ref=E6 rowPageCount=3 colPageCount=1 pageWrap=1");

        // The counts are not read back - the size is derived again from the wrap, the order and the
        // filter count - so a reload has to reach the same geometry.
        await Assert.That(ReloadedFilterArea(saved)).IsEqualTo("target=E2 area=E6 filters=1x3");
    }

    /// <summary>
    /// The target cell, the area and the size of the filter area in one line. The target cell reads
    /// <c>#REF!</c> when the filters need more rows than there are above the area, because
    /// <c>Point</c> stores the row 0-based in an unsigned field and a row of 0 wraps.
    /// </summary>
    private static string Describe(XLPivotTable pt)
    {
        var size = pt.Filters.GetSize();
        return $"target={pt.TargetCell.Address} area={pt.Area} filters={size.Width}x{size.Height}";
    }

    private static XLPivotTable CreatePivotTable(XLWorkbook wb, string targetCell)
    {
        var ws = wb.AddWorksheet("Data");
        var range = ws.Cell("A1").InsertData(new object[]
        {
            ("F1", "F2", "F3", "Name", "Value"),
            ("a", "b", "c", "Cake", 1),
        });

        return (XLPivotTable)ws.PivotTables.Add("pt", ws.Cell(targetCell), range!);
    }

    private static void AddFilters(XLPivotTable pt, int filterCount)
    {
        for (var i = 1; i <= filterCount; i++)
            pt.ReportFilters.Add($"F{i}");
    }

    private static string SavedFilterArea(Stream package)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        var definition = document.WorkbookPart!.WorksheetParts
            .SelectMany(part => part.PivotTableParts)
            .Select(part => part.PivotTableDefinition!)
            .Single();
        var location = definition.Location!;
        return $"ref={location.Reference!.Value} " +
               $"rowPageCount={Attribute(location, "rowPageCount")} " +
               $"colPageCount={Attribute(location, "colPageCount")} " +
               $"pageWrap={Attribute(definition, "pageWrap")}";
    }

    /// <summary>
    /// The attribute as the file holds it, or the value Excel assumes when it is absent. Read by
    /// name so that the assertion is about the file, not about how the SDK names the attribute.
    /// </summary>
    private static string Attribute(OpenXmlElement element, string localName)
    {
        return element.GetAttributes().FirstOrDefault(a => a.LocalName == localName).Value ?? "0";
    }

    private static string ReloadedFilterArea(Stream package)
    {
        package.Position = 0;
        using var wb = new XLWorkbook(package);
        return Describe((XLPivotTable)wb.Worksheet("Data").PivotTables.Single());
    }
}
