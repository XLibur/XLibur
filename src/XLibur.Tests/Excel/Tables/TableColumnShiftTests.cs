using System.IO;
using System.Linq;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Tables;

/// <summary>
/// A worksheet column deleted or inserted inside a table changes its fields the same way with the
/// header row hidden as with it shown: the deleted column's field goes, an inserted column gets a
/// new field, and the fields of the other columns keep their objects and move with them (#1013).
/// </summary>
/// <remarks>
/// The fixture's <c>SalesTable</c> sits at <c>A1:C15</c> with columns Region, Date and Amount.
/// </remarks>
public class TableColumnShiftTests
{
    private const string Fixture = @"TryToLoad\SlicersOnPivotAndTable.xlsx";

    [Test]
    [Arguments(true)]
    [Arguments(false)]
    public async Task Deleting_the_first_column_removes_its_field_and_shifts_the_rest(bool showHeaderRow)
    {
        using var wb = TestHelper.LoadWorkbook(Fixture);
        var ws = wb.Worksheet("Data");
        var table = ws.Table("SalesTable");
        table.ShowHeaderRow = showHeaderRow;
        var date = table.Field("Date");
        var amount = table.Field("Amount");

        ws.Column("A").Delete();

        await Assert.That(FieldsOf(table)).IsEquivalentTo(new[] { "Date@0", "Amount@1" }, CollectionOrdering.Matching);
        await Assert.That(table.Field("Date")).IsSameReferenceAs(date);
        await Assert.That(table.Field("Amount")).IsSameReferenceAs(amount);
    }

    [Test]
    [Arguments(true)]
    [Arguments(false)]
    public async Task Deleting_a_middle_column_removes_its_field_and_shifts_the_rest(bool showHeaderRow)
    {
        using var wb = TestHelper.LoadWorkbook(Fixture);
        var ws = wb.Worksheet("Data");
        var table = ws.Table("SalesTable");
        table.ShowHeaderRow = showHeaderRow;
        var region = table.Field("Region");
        var amount = table.Field("Amount");

        ws.Column("B").Delete();

        await Assert.That(FieldsOf(table)).IsEquivalentTo(new[] { "Region@0", "Amount@1" }, CollectionOrdering.Matching);
        await Assert.That(table.Field("Region")).IsSameReferenceAs(region);
        await Assert.That(table.Field("Amount")).IsSameReferenceAs(amount);
    }

    [Test]
    public async Task A_headerless_table_saves_the_columns_left_after_a_delete()
    {
        using var saved = new MemoryStream();
        using (var wb = TestHelper.LoadWorkbook(Fixture))
        {
            var ws = wb.Worksheet("Data");
            ws.Table("SalesTable").ShowHeaderRow = false;
            ws.Column("A").Delete();
            wb.SaveAs(saved);
        }

        var tableXml = saved.ReadPart("xl/tables/table1.xml");
        await Assert.That(tableXml).Contains("ref=\"A2:B15\"");
        await Assert.That(tableXml).Contains("name=\"Date\"");
        await Assert.That(tableXml).Contains("name=\"Amount\"");
        await Assert.That(tableXml).DoesNotContain("name=\"Region\"");
    }

    [Test]
    public async Task Inserting_a_column_inside_a_headerless_table_adds_a_field_at_its_position()
    {
        using var wb = TestHelper.LoadWorkbook(Fixture);
        var ws = wb.Worksheet("Data");
        var table = ws.Table("SalesTable");
        table.ShowHeaderRow = false;
        var date = table.Field("Date");

        ws.Column("B").InsertColumnsBefore(1);

        var fields = table.Fields.ToList();
        await Assert.That(fields.Select(f => f.Index)).IsEquivalentTo(new[] { 0, 1, 2, 3 }, CollectionOrdering.Matching);
        await Assert.That(fields[0].Name).IsEqualTo("Region");
        await Assert.That(fields[2]).IsSameReferenceAs(date);
        await Assert.That(fields[3].Name).IsEqualTo("Amount");
        await Assert.That(new[] { "Region", "Date", "Amount" }).DoesNotContain(fields[1].Name);
    }

    [Test]
    public async Task Deleting_a_field_of_a_headerless_table_shifts_the_rest_once()
    {
        using var wb = TestHelper.LoadWorkbook(Fixture);
        var table = wb.Worksheet("Data").Table("SalesTable");
        table.ShowHeaderRow = false;

        table.Field("Region").Delete();

        await Assert.That(FieldsOf(table)).IsEquivalentTo(new[] { "Date@0", "Amount@1" }, CollectionOrdering.Matching);
    }

    [Test]
    public async Task A_kept_field_follows_its_column_as_soon_as_another_field_is_deleted()
    {
        using var wb = TestHelper.LoadWorkbook(Fixture);
        var ws = wb.Worksheet("Data");
        var table = ws.Table("SalesTable");
        var amount = table.Field("Amount");

        table.Field("Date").Delete();
        amount.Name = "Total";

        await Assert.That(ws.Cell("B1").GetText()).IsEqualTo("Total");
        await Assert.That(ws.Cell("C1").IsEmpty()).IsTrue();
    }

    [Test]
    public async Task Deleting_a_table_column_of_a_headerless_table_removes_its_field()
    {
        using var wb = TestHelper.LoadWorkbook(Fixture);
        var table = wb.Worksheet("Data").Table("SalesTable");
        table.ShowHeaderRow = false;

        table.Column(1).Delete();

        await Assert.That(FieldsOf(table)).IsEquivalentTo(new[] { "Date@0", "Amount@1" }, CollectionOrdering.Matching);
    }

    [Test]
    [Arguments(true)]
    [Arguments(false)]
    public async Task Inserting_columns_before_a_table_with_expandRange_adds_a_field_at_index_0(bool showHeaderRow)
    {
        using var wb = new XLWorkbook();
        var table = CreatePTable(wb, showHeaderRow);
        var pa = table.Field("Pa");

        ((IXLRange)table).InsertColumnsBefore(1, true);

        var fields = table.Fields.ToList();
        await Assert.That(fields.Select(f => $"{f.Name}@{f.Index}").Skip(1))
            .IsEquivalentTo(new[] { "Pa@1", "Pb@2", "Pc@3", "Pd@4" }, CollectionOrdering.Matching);
        await Assert.That(fields[0].Index).IsEqualTo(0);
        await Assert.That(new[] { "Pa", "Pb", "Pc", "Pd" }).DoesNotContain(fields[0].Name);
        await Assert.That(table.Field("Pa")).IsSameReferenceAs(pa);
        await Assert.That(table.RangeAddress.ToString()).IsEqualTo(showHeaderRow ? "C2:G5" : "C3:G5");
    }

    [Test]
    [Arguments(true)]
    [Arguments(false)]
    public async Task Inserting_columns_after_a_table_with_expandRange_adds_a_last_field(bool showHeaderRow)
    {
        using var wb = new XLWorkbook();
        var table = CreatePTable(wb, showHeaderRow);

        ((IXLRange)table).InsertColumnsAfter(1, true);

        var fields = FieldsOf(table);
        await Assert.That(fields.Take(4))
            .IsEquivalentTo(new[] { "Pa@0", "Pb@1", "Pc@2", "Pd@3" }, CollectionOrdering.Matching);
        await Assert.That(fields.Length).IsEqualTo(5);
        await Assert.That(fields[4]).EndsWith("@4");
        await Assert.That(table.RangeAddress.ToString()).IsEqualTo(showHeaderRow ? "C2:G5" : "C3:G5");
    }

    [Test]
    public async Task Inserting_rows_above_a_headerless_table_with_expandRange_keeps_its_fields()
    {
        using var wb = new XLWorkbook();
        var table = CreatePTable(wb, showHeaderRow: false);

        ((IXLRange)table).InsertRowsAbove(1, true);

        await Assert.That(FieldsOf(table))
            .IsEquivalentTo(new[] { "Pa@0", "Pb@1", "Pc@2", "Pd@3" }, CollectionOrdering.Matching);
    }

    /// <summary>A table with fields Pa to Pd at <c>C2:F5</c>, the repro of #1017.</summary>
    private static IXLTable CreatePTable(XLWorkbook wb, bool showHeaderRow)
    {
        var ws = wb.AddWorksheet();
        string[] names = ["Pa", "Pb", "Pc", "Pd"];
        for (var i = 0; i < names.Length; i++)
            ws.Cell(2, 3 + i).Value = names[i];

        for (var row = 3; row <= 5; row++)
            for (var col = 3; col <= 6; col++)
                ws.Cell(row, col).Value = row * 10 + col;

        var table = ws.Range("C2:F5").CreateTable();
        table.ShowHeaderRow = showHeaderRow;
        _ = table.Fields.ToList();
        return table;
    }

    private static string[] FieldsOf(IXLTable table) =>
        table.Fields.Select(f => $"{f.Name}@{f.Index}").ToArray();
}
