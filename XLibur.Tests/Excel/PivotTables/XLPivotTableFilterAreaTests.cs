using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using XLibur.Excel.Coordinates;

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

    [Test]
    [Property("Description", "#571 review: Clear took every report filter off without moving the table, so the target cell ended up where the table itself was")]
    public async Task Clearing_the_report_filters_takes_the_table_back_up_to_its_target_cell()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb, "E1");
        AddFilters(pt, 3);

        await Assert.That(Describe(pt)).IsEqualTo("target=E1 area=E5 filters=1x3");

        pt.ReportFilters.Clear();

        // Nothing above the table any more, so the table is back at the target cell, as it is
        // when the same three filters are removed one at a time.
        await Assert.That(Describe(pt)).IsEqualTo("target=E1 area=E1 filters=0x0");
    }

    [Test]
    [Property("Description", "#571 review: Remove shifted the area with no lower bound, so a file whose ref has no room for its filters moved the target cell off the top of the sheet")]
    public async Task Removing_a_filter_keeps_the_target_cell_on_the_sheet_when_the_area_has_no_room_for_the_filters()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb, "E1");
        pt.FilterAreaOrder = XLFilterAreaOrder.OverThenDown;
        pt.FilterFieldsPageWrap = 1;
        AddFilters(pt, 3);

        // The geometry a file saved before this fix holds: a ref with nowhere near enough room
        // above it for the filter area the wrap asks for. Set directly, because XLibur no longer
        // produces it.
        pt.Area = new Area(2, 5, 2, 5);
        await Assert.That(Describe(pt)).IsEqualTo("target=#REF! area=E2 filters=1x3")
            .Because("the area must start above the filters, or this proves nothing");

        pt.ReportFilters.Remove("F3");

        // Two filters and the divider row need three rows, so the table is held at row 4 and the
        // target cell lands on row 1 rather than off the top of the sheet.
        await Assert.That(Describe(pt)).IsEqualTo("target=E1 area=E4 filters=1x2");
    }

    [Test]
    [Property("Description", "#578: the shift clipped the area at the last row, so a table at the bottom of the sheet lost height when it gained a filter")]
    public async Task A_filter_added_to_a_table_at_the_bottom_of_the_sheet_keeps_the_table_whole()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb, "E1");

        // A ten-row table whose last row is the last row of the sheet. Set directly, because the
        // area of a table built here is one row high.
        pt.Area = new Area(XLHelper.MaxRowNumber - 9, 5, XLHelper.MaxRowNumber, 5);

        pt.ReportFilters.Add("F1");

        // The filter and its gap row need two rows and there are none below the table to move
        // into, so the table keeps all ten of its rows and takes the two rows above it instead.
        // The target cell is the cost of that: it is two rows higher than where the table was put.
        await Assert.That(Describe(pt)).IsEqualTo(
            $"target=E{XLHelper.MaxRowNumber - 11} " +
            $"area=E{XLHelper.MaxRowNumber - 9}:E{XLHelper.MaxRowNumber} filters=1x1");
    }

    [Test]
    [Property("Description", "#578: the cap must only bite when the rows really are not there, so a table with room below it moves the whole shift")]
    public async Task A_filter_added_to_a_table_with_room_below_it_moves_the_table_the_whole_shift()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb, "E1");

        // The same ten-row table, eleven rows higher: the shift has somewhere to go.
        pt.Area = new Area(XLHelper.MaxRowNumber - 20, 5, XLHelper.MaxRowNumber - 11, 5);

        pt.ReportFilters.Add("F1");

        // The filter area needs two rows and there are eleven below the table, so the table moves
        // down by the whole two and the target cell stays exactly where it was put.
        await Assert.That(Describe(pt)).IsEqualTo(
            $"target=E{XLHelper.MaxRowNumber - 20} " +
            $"area=E{XLHelper.MaxRowNumber - 18}:E{XLHelper.MaxRowNumber - 9} filters=1x1");
    }

    [Test]
    [Property("Description", "#578: a table taller than the rows left below its filter area cannot keep both, and the filters have to win")]
    public async Task A_table_too_tall_for_the_rows_below_its_filter_area_keeps_the_filters_room()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb, "E1");

        // A table that covers all but the first row of the sheet, so there is no arrangement that
        // both fits it whole and leaves two rows above it for the filter area.
        pt.Area = new Area(2, 5, XLHelper.MaxRowNumber, 5);

        pt.ReportFilters.Add("F1");

        // The filters keep their two rows and the extent gives up the single row that cannot
        // exist, because the alternative is the target cell above row 1 that #571 was about.
        await Assert.That(Describe(pt))
            .IsEqualTo($"target=E1 area=E3:E{XLHelper.MaxRowNumber} filters=1x1");
    }

    [Test]
    [Property("Description", "#578: taking the filter off again must not need the rows the table never got, so the height survives the round trip")]
    public async Task Taking_the_filter_off_a_table_at_the_bottom_of_the_sheet_keeps_the_table_whole()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb, "E1");
        pt.Area = new Area(XLHelper.MaxRowNumber - 9, 5, XLHelper.MaxRowNumber, 5);
        pt.ReportFilters.Add("F1");

        pt.ReportFilters.Remove("F1");

        // Nothing above the table any more, so the area comes back up to the target cell the added
        // filter left it with. The ten rows are still there, which is what the cap is for: the
        // table has slid two rows up the sheet, but it is the table it was.
        await Assert.That(Describe(pt)).IsEqualTo(
            $"target=E{XLHelper.MaxRowNumber - 11} " +
            $"area=E{XLHelper.MaxRowNumber - 11}:E{XLHelper.MaxRowNumber - 2} filters=0x0");
    }

    [Test]
    [Property("Description", "#578: a sheet edit really does consume the rows, so it still clips and must behave exactly as it did before #574")]
    public async Task An_insert_that_pushes_a_table_off_the_bottom_of_the_sheet_still_clips_it()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb, "E1");
        pt.Area = new Area(XLHelper.MaxRowNumber - 9, 5, XLHelper.MaxRowNumber, 5);

        wb.Worksheet("Data").Row(1000).InsertRowsAbove(5);

        // The five rows the insert put above the table are five the table's last rows fall off the
        // sheet to make room for. Unlike a filter needing room, the rows were genuinely consumed,
        // so the extent is clipped at the last row, as every other extent on the sheet is.
        await Assert.That(Describe(pt)).IsEqualTo(
            $"target=E{XLHelper.MaxRowNumber - 4} " +
            $"area=E{XLHelper.MaxRowNumber - 4}:E{XLHelper.MaxRowNumber} filters=0x0");
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
