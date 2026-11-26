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

    private static string[] FieldsOf(IXLTable table) =>
        table.Fields.Select(f => $"{f.Name}@{f.Index}").ToArray();
}
