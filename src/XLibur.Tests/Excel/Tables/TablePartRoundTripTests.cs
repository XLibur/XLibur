using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using XLibur.Excel;
using XLibur.Excel.Tables;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Tables;

public class TablePartRoundTripTests
{
    private const string Slicers = @"TryToLoad\SlicersOnPivotAndTable.xlsx";
    private const string ColumnStyle = @"TryToLoad\LO\xlsx\column-style-autofilter.xlsx";
    private const string ImageReference = @"TryToLoad\external_image_reference_2608.xlsx";
    private const string TableSourcePivots = @"TryToLoad\TemplateWithTableSourcePivotTables.xlsx";
    private const string Timelines = @"TryToLoad\Timelines_Missing_21232.xlsx";

    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    [Test]
    public async Task A_table_slicer_still_finds_its_table_after_a_save_when_the_table_id_is_not_1()
    {
        // Excel writes a table id past 1 whenever an earlier table was deleted.
        using var source = TableIdFive();

        using var saved = source.LoadAndSave();

        await Assert.That(TableXml(saved, "SalesTable").Attribute("id")?.Value).IsEqualTo("5");

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var tableSlicer = reloaded.Worksheets.SelectMany(ws => ws.Slicers)
            .Single(s => s.SourceKind == XLSlicerSourceKind.Table);
        await Assert.That(tableSlicer.Table?.Name).IsEqualTo("SalesTable");
    }

    [Test]
    public async Task A_save_keeps_the_ids_of_the_columns_in_the_order_the_file_has_them()
    {
        using var saved = TestHelper.OpenResource(ColumnStyle).LoadAndSave();

        await Assert.That(ColumnIds(TableXml(saved, "HateIndex"))).IsEquivalentTo(new[] { "1", "3", "2", "4", "5" },
            TUnit.Assertions.Enums.CollectionOrdering.Matching);
    }

    [Test]
    public async Task A_save_keeps_the_ids_of_every_loaded_table()
    {
        using var source = TestHelper.OpenResource(ColumnStyle);
        var before = TableIdsByName(source);

        using var saved = source.LoadAndSave();

        await Assert.That(TableIdsByName(saved)).IsEquivalentTo(before);
    }

    [Test]
    public async Task A_new_table_takes_an_id_no_loaded_table_has()
    {
        using var source = TableIdFive();
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(source))
        {
            var ws = wb.AddWorksheet("New");
            ws.Cell("A1").InsertTable(new[] { new { Name = "a", Value = 1 } }, "NewTable");
            wb.SaveAs(saved);
        }

        var ids = TableIdsByName(saved);
        await Assert.That(ids["SalesTable"]).IsEqualTo("5");
        await Assert.That(ids["NewTable"]).IsEqualTo("6");
    }

    [Test]
    public async Task A_new_column_takes_an_id_no_column_of_its_table_had()
    {
        using var saved = new MemoryStream();
        using (var wb = TestHelper.LoadWorkbook(ColumnStyle))
        {
            var table = wb.Table("HateIndex");
            table.Resize(table.Worksheet.Range("A1:F10"));
            wb.SaveAs(saved);
        }

        await Assert.That(ColumnIds(TableXml(saved, "HateIndex"))).IsEquivalentTo(new[] { "1", "3", "2", "4", "5", "6" },
            TUnit.Assertions.Enums.CollectionOrdering.Matching);
    }

    [Test]
    public async Task A_deleted_column_keeps_its_id_from_being_given_to_a_new_one()
    {
        using var saved = new MemoryStream();
        using (var wb = TestHelper.LoadWorkbook(ColumnStyle))
        {
            var table = wb.Table("HateIndex");
            table.Field("abs").Delete();
            table.Resize(table.Worksheet.Range("A1:E10"));
            wb.SaveAs(saved);
        }

        await Assert.That(ColumnIds(TableXml(saved, "HateIndex"))).IsEquivalentTo(new[] { "1", "3", "2", "4", "6" },
            TUnit.Assertions.Enums.CollectionOrdering.Matching);
    }

    [Test]
    public async Task A_new_table_slicer_binds_to_the_id_of_its_column_rather_than_its_position()
    {
        using var saved = new MemoryStream();
        using (var wb = TestHelper.LoadWorkbook(ColumnStyle))
        {
            var table = (XLTable)wb.Table("HateIndex");
            ((XLWorksheet)table.Worksheet).SlicersInternal.AddTableSlicer(table, "Hate Count");
            wb.SaveAs(saved);
        }

        var cache = saved.PartsUnder("xl/slicerCaches/").Select(p => saved.ReadPart(p))
            .Single(x => x.Contains("tableSlicerCache"));
        await Assert.That(cache).Contains("tableId=\"7\"");
        await Assert.That(cache).Contains("column=\"2\"");
    }

    [Test]
    [Arguments(ColumnStyle, "HateIndex", "A2:E10")]
    [Arguments(ImageReference, "Таблица1", "E2:E16")]
    public async Task A_save_keeps_the_sort_state(string fixture, string tableName, string sortRef)
    {
        using var saved = TestHelper.OpenResource(fixture).LoadAndSave();

        var table = TableXml(saved, tableName);
        await Assert.That(table.Element(Main + "sortState")?.Attribute("ref")?.Value).IsEqualTo(sortRef);
    }

    [Test]
    public async Task A_save_keeps_the_header_and_border_formats_of_a_table()
    {
        using var source = TestHelper.OpenResource(ColumnStyle);
        var before = TableXml(source, "HateIndex");

        using var saved = source.LoadAndSave();

        var after = TableXml(saved, "HateIndex");
        foreach (var attribute in new[] { "headerRowDxfId", "headerRowBorderDxfId", "tableBorderDxfId" })
        {
            var dxfId = after.Attribute(attribute)?.Value;
            await Assert.That(dxfId).IsNotNull().Because($"{attribute} must survive the save");
            await Assert.That(Dxf(saved, dxfId!)).IsEqualTo(Dxf(source, before.Attribute(attribute)!.Value))
                .Because($"{attribute} must name the format it named before the save");
        }
    }

    [Test]
    public async Task A_save_keeps_the_data_format_of_a_column_whose_cells_differ_in_style()
    {
        using var source = TestHelper.OpenResource(ImageReference);
        var before = TableXmls(source).SelectMany(t => ColumnsOf(t).Select(c => (Table: t.Attribute("displayName")!.Value, Column: c)))
            .Where(x => x.Column.Attribute("dataDxfId") is not null)
            .ToArray();

        using var saved = source.LoadAndSave();

        await Assert.That(before.Length).IsEqualTo(4);
        foreach (var (table, column) in before)
        {
            var name = column.Attribute("name")!.Value;
            var dxfId = Column(TableXml(saved, table), name).Attribute("dataDxfId")?.Value;
            await Assert.That(dxfId).IsNotNull().Because($"{table}[{name}] must keep its data format");
            await Assert.That(Dxf(saved, dxfId!)).IsEqualTo(Dxf(source, column.Attribute("dataDxfId")!.Value));
        }
    }

    [Test]
    public async Task A_column_whose_cells_share_a_style_takes_that_style_as_its_data_format()
    {
        using var source = TestHelper.OpenResource(ImageReference);
        var (tableName, columnName, loadedDxfId) = TableXmls(source)
            .SelectMany(t => ColumnsOf(t).Select(c => (t.Attribute("displayName")!.Value, c.Attribute("name")!.Value, c.Attribute("dataDxfId")?.Value)))
            .First(x => x.Item3 is not null);

        using var saved = source.LoadAndSave(wb =>
        {
            var cells = wb.Table(tableName).Field(columnName).DataCells;
            cells.Style = XLWorkbook.DefaultStyle;
            cells.Style.Fill.SetBackgroundColor(XLColor.Red);
        });

        var dxfId = Column(TableXml(saved, tableName), columnName).Attribute("dataDxfId")?.Value;
        await Assert.That(dxfId).IsNotNull();
        await Assert.That(Dxf(saved, dxfId!)).IsNotEqualTo(Dxf(source, loadedDxfId!))
            .Because("A style every data cell of the column shares is the model's format and replaces the loaded one.");
    }

    [Test]
    public async Task A_save_keeps_the_header_format_of_a_table_slicer_fixture()
    {
        using var source = TestHelper.OpenResource(Slicers);
        var before = TableXml(source, "SalesTable");

        using var saved = source.LoadAndSave();

        var dxfId = TableXml(saved, "SalesTable").Attribute("headerRowDxfId")?.Value;
        await Assert.That(dxfId).IsNotNull();
        await Assert.That(Dxf(saved, dxfId!)).IsEqualTo(Dxf(source, before.Attribute("headerRowDxfId")!.Value));
    }

    [Test]
    public async Task A_hidden_totals_row_keeps_the_totals_it_remembers()
    {
        using var source = TestHelper.OpenResource(TableSourcePivots);
        var before = TableXml(source, "PerSchemeRequiredBonds");

        using var saved = source.LoadAndSave();

        var after = TableXml(saved, "PerSchemeRequiredBonds");
        await Assert.That(after.Attribute("totalsRowShown")).IsNull()
            .Because("The totals row has been shown, which is what lets Excel turn it back on with its totals.");
        await Assert.That(Column(after, "Calculation_id").Attribute("totalsRowLabel")?.Value).IsEqualTo("Total");
        await Assert.That(Column(after, "Cash").Attribute("totalsRowFunction")?.Value).IsEqualTo("sum");

        var dxfId = Column(after, "Cash").Attribute("totalsRowDxfId")?.Value;
        await Assert.That(dxfId).IsNotNull();
        await Assert.That(Dxf(saved, dxfId!)).IsEqualTo(Dxf(source, Column(before, "Cash").Attribute("totalsRowDxfId")!.Value));
    }

    [Test]
    public async Task A_save_keeps_the_insert_row()
    {
        using var saved = TestHelper.OpenResource(TableSourcePivots).LoadAndSave();

        await Assert.That(TableXml(saved, "StressedLiabilities").Attribute("insertRow")?.Value).IsEqualTo("1");
    }

    [Test]
    public async Task A_save_keeps_the_named_cell_styles_of_a_table()
    {
        using var source = TestHelper.OpenResource(Timelines);
        var before = TableXmls(source).Single();

        using var saved = source.LoadAndSave();

        var after = TableXmls(saved).Single();
        await Assert.That(after.Attribute("headerRowCellStyle")?.Value).IsEqualTo(before.Attribute("headerRowCellStyle")!.Value);
        await Assert.That(ColumnsOf(after).Select(c => c.Attribute("dataCellStyle")?.Value))
            .IsEquivalentTo(ColumnsOf(before).Select(c => c.Attribute("dataCellStyle")?.Value));
    }

    [Test]
    [Arguments(ColumnStyle)]
    [Arguments(ImageReference)]
    [Arguments(TableSourcePivots)]
    [Arguments(Timelines)]
    [Arguments(Slicers)]
    public async Task A_saved_table_part_is_valid(string fixture)
    {
        using var saved = TestHelper.OpenResource(fixture).LoadAndSave();

        saved.Position = 0;
        using var document = SpreadsheetDocument.Open(saved, false);
        var errors = new OpenXmlValidator().Validate(document)
            .Where(e => e.Part?.Uri.OriginalString.Contains("/tables/", StringComparison.Ordinal) == true)
            .Select(e => $"{e.Part!.Uri} {e.Path?.XPath}: {e.Description}")
            .ToArray();

        await Assert.That(errors).IsEmpty();
    }

    [Test]
    public async Task A_resized_table_drops_its_sort_state()
    {
        using var saved = new MemoryStream();
        using (var wb = TestHelper.LoadWorkbook(ColumnStyle))
        {
            var table = wb.Table("HateIndex");
            table.Resize(table.Worksheet.Range("A1:E12"));
            wb.SaveAs(saved);
        }

        await Assert.That(TableXml(saved, "HateIndex").Element(Main + "sortState")).IsNull()
            .Because("A sort state over the old range names cells the table no longer has in that place.");
    }

    [Test]
    public async Task A_sort_of_the_table_drops_the_sort_state_it_replaced()
    {
        using var saved = new MemoryStream();
        using (var wb = TestHelper.LoadWorkbook(ColumnStyle))
        {
            wb.Table("HateIndex").DataRange!.Sort(1, XLSortOrder.Descending);
            wb.SaveAs(saved);
        }

        await Assert.That(TableXml(saved, "HateIndex").Element(Main + "sortState")).IsNull()
            .Because("The loaded sort state names column D, and Excel would re-sort the rows by it on Reapply.");
    }

    [Test]
    public async Task A_sort_of_a_range_overlapping_the_table_drops_its_sort_state()
    {
        using var saved = new MemoryStream();
        using (var wb = TestHelper.LoadWorkbook(ColumnStyle))
        {
            var ws = wb.Table("HateIndex").Worksheet;
            ws.Range("A5:E10").Sort(2);
            wb.SaveAs(saved);
        }

        await Assert.That(TableXml(saved, "HateIndex").Element(Main + "sortState")).IsNull();
    }

    [Test]
    public async Task A_sort_or_an_edit_that_leaves_the_table_order_alone_keeps_its_sort_state()
    {
        using var saved = new MemoryStream();
        using (var wb = TestHelper.LoadWorkbook(ColumnStyle))
        {
            var ws = wb.Table("HateIndex").Worksheet;
            ws.Cell("H1").Value = 2;
            ws.Cell("H2").Value = 1;
            ws.Range("H1:H2").Sort();
            ws.Cell("C3").Value = 999;
            wb.SaveAs(saved);
        }

        await Assert.That(TableXml(saved, "HateIndex").Element(Main + "sortState")).IsNotNull()
            .Because("Excel keeps a table's sort state when a value in it changes, and Reapply sorts again.");
    }

    [Test]
    public async Task A_renamed_column_keeps_its_id_the_table_formats_and_the_sort_state()
    {
        using var saved = new MemoryStream();
        using (var wb = TestHelper.LoadWorkbook(ColumnStyle))
        {
            wb.Table("HateIndex").Field("Hate Count").Name = "Hate";
            wb.SaveAs(saved);
        }

        var table = TableXml(saved, "HateIndex");
        await Assert.That(Column(table, "Hate").Attribute("id")?.Value).IsEqualTo("2");
        await Assert.That(table.Attribute("headerRowDxfId")).IsNotNull();
        await Assert.That(table.Element(Main + "sortState")).IsNotNull()
            .Because("A sort condition names its column by cell reference, which a rename leaves where it was.");
    }

    [Test]
    public async Task A_renamed_column_drops_the_totals_a_hidden_totals_row_remembers()
    {
        using var saved = new MemoryStream();
        using (var wb = TestHelper.LoadWorkbook(TableSourcePivots))
        {
            wb.Table("PerSchemeRequiredBonds").Field("Cash").Name = "Money";
            wb.SaveAs(saved);
        }

        var table = TableXml(saved, "PerSchemeRequiredBonds");
        await Assert.That(Column(table, "Money").Attribute("totalsRowFunction")).IsNull();
        await Assert.That(Column(table, "Calculation_id").Attribute("totalsRowLabel")).IsNull();
    }

    [Test]
    public async Task A_second_save_keeps_what_the_first_save_carried()
    {
        using var source = TestHelper.OpenResource(ColumnStyle);
        var before = TableXml(source, "HateIndex");

        using var once = source.LoadAndSave();
        using var twice = once.LoadAndSave();

        var after = TableXml(twice, "HateIndex");
        await Assert.That(after.Element(Main + "sortState")).IsNotNull();
        await Assert.That(Dxf(twice, after.Attribute("headerRowDxfId")!.Value))
            .IsEqualTo(Dxf(source, before.Attribute("headerRowDxfId")!.Value));
    }

    [Test]
    public async Task A_loaded_table_with_filter_buttons_sorts_through_its_autofilter()
    {
        using var wb = TestHelper.LoadWorkbook(Slicers);
        var table = wb.Table("SalesTable");
        var amounts = table.DataRange!.Column(3).Cells().Select(c => c.GetDouble()).ToArray();

        table.AutoFilter.Sort(3, XLSortOrder.Descending);

        await Assert.That(table.AutoFilter.IsEnabled).IsTrue();
        await Assert.That(table.DataRange!.Column(3).Cells().Select(c => c.GetDouble()))
            .IsEquivalentTo(amounts.OrderByDescending(a => a), TUnit.Assertions.Enums.CollectionOrdering.Matching);
    }

    [Test]
    public async Task A_loaded_table_without_filter_buttons_keeps_its_autofilter_off_through_a_save()
    {
        using var source = TestHelper.OpenResource(Slicers);
        source.RewritePart("xl/tables/table1.xml", xml => System.Text.RegularExpressions.Regex.Replace(xml, "<autoFilter [^>]*/>", ""));

        bool showAutoFilter = true, isEnabled = true;
        using var saved = source.LoadAndSave(wb =>
        {
            var table = wb.Table("SalesTable");
            showAutoFilter = table.ShowAutoFilter;
            isEnabled = table.AutoFilter.IsEnabled;
        });

        await Assert.That(showAutoFilter).IsFalse();
        await Assert.That(isEnabled).IsFalse();
        await Assert.That(TableXml(saved, "SalesTable").Element(Main + "autoFilter")).IsNull();
    }

    private static MemoryStream TableIdFive()
    {
        var source = TestHelper.OpenResource(Slicers);
        source.RewritePart("xl/tables/table1.xml", xml => xml.Replace(" id=\"1\" xr:uid", " id=\"5\" xr:uid"));
        source.RewritePart("xl/slicerCaches/slicerCache2.xml", xml => xml.Replace("tableId=\"1\"", "tableId=\"5\""));
        return source;
    }

    private static XElement[] TableXmls(Stream package) =>
        package.PartsUnder("xl/tables/").Where(p => p.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .Select(p => XElement.Parse(package.ReadPart(p)))
            .ToArray();

    private static XElement TableXml(Stream package, string name) =>
        TableXmls(package).Single(t => t.Attribute("name")?.Value == name);

    private static System.Collections.Generic.Dictionary<string, string> TableIdsByName(Stream package) =>
        TableXmls(package).ToDictionary(t => t.Attribute("name")!.Value, t => t.Attribute("id")!.Value);

    private static XElement[] ColumnsOf(XElement table) =>
        table.Element(Main + "tableColumns")!.Elements(Main + "tableColumn").ToArray();

    private static string[] ColumnIds(XElement table) =>
        ColumnsOf(table).Select(c => c.Attribute("id")!.Value).ToArray();

    private static XElement Column(XElement table, string name) =>
        ColumnsOf(table).Single(c => c.Attribute("name")?.Value == name);

    private static string Dxf(Stream package, string index)
    {
        var styles = XElement.Parse(package.ReadPart("xl/styles.xml"));
        return Canonical(styles.Element(Main + "dxfs")!.Elements(Main + "dxf").ElementAt(int.Parse(index)));
    }

    private static string Canonical(XElement element) =>
        element.Name.LocalName
        + "[" + string.Join(",", element.Attributes().Where(a => !a.IsNamespaceDeclaration)
            .Select(a => a.Name.LocalName + "=" + a.Value).Order()) + "]"
        + "(" + string.Concat(element.Elements().Select(Canonical)) + ")";
}
