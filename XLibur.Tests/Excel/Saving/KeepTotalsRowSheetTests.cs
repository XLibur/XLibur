using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Saving;

/// <summary>
/// A sheet whose table has a totals row keeps its part, or its cells, on the first save of a loaded
/// workbook when the sheet has not changed (#702, #721).
/// </summary>
/// <remarks>
/// The writer writes a totals-row cell with no formula as any other cell, from its value, and the
/// load leaves the cell as the file has it, whatever the field's label (#742). A label or totals
/// function set through the API is written into the cell, so the save sees a change.
/// </remarks>
public class KeepTotalsRowSheetTests
{
    private const string Sheet1 = "xl/worksheets/sheet1.xml";

    [Test]
    public async Task An_unchanged_sheet_with_a_labelled_totals_row_keeps_its_part()
    {
        using var source = Source(label: "Total");

        using var saved = source.LoadAndSave();

        await Assert.That(saved.IsKeptWhole(Sheet1)).IsTrue();
        await AssertValues(saved);
        await AssertTotals(saved, "Total");
    }

    [Test]
    public async Task An_unchanged_sheet_with_an_empty_totals_row_keeps_its_part()
    {
        using var source = Source(label: null);

        using var saved = source.LoadAndSave();

        await Assert.That(saved.IsKeptWhole(Sheet1)).IsTrue();
        await AssertValues(saved);
    }

    [Test]
    public async Task Loading_a_label_the_cell_already_holds_leaves_the_cells_unchanged()
    {
        using var source = Source(label: "Total");
        source.Position = 0;

        using var wb = new XLWorkbook(source);
        var ws = (XLWorksheet)wb.Worksheet("Data");

        await Assert.That(ws.CellsChangedSinceLoad()).IsFalse();
        await Assert.That(ws.Table("Sales").Field("Name").TotalsRowLabel).IsEqualTo("Total");
        await Assert.That(ws.Table("Sales").Field("Name").TotalsRowFunction).IsEqualTo(XLTotalsRowFunction.None);
    }

    [Test]
    public async Task A_label_with_a_leading_quote_the_cell_already_holds_keeps_its_quote()
    {
        // Excel writes a label typed as ''Total so (checked over COM): the cell holds 'Total under
        // quotePrefix="1", and the table part has totalsRowLabel="'Total". The style needs no change,
        // but the label's setter would still strip the leading quote, and the save would then write
        // Total: the load must leave the cell as the file has it.
        using var source = QuotedLabelSource();
        await Assert.That(source.ReadPart(TablePart(source))).Contains("totalsRowLabel=\"'Total\"");
        source.Position = 0;

        using (var wb = new XLWorkbook(source))
        {
            var ws = wb.Worksheet("Data");
            await Assert.That(((XLWorksheet)ws).CellsChangedSinceLoad()).IsFalse();
            await Assert.That(ws.Cell("A4").GetText()).IsEqualTo("'Total");
            await Assert.That(ws.Cell("A4").Style.IncludeQuotePrefix).IsTrue();
            await Assert.That(ws.Table("Sales").Field("Name").TotalsRowLabel).IsEqualTo("'Total");
        }

        using var saved = source.LoadAndSave();

        await Assert.That(saved.IsKeptWhole(Sheet1)).IsTrue();
        await AssertTotals(saved, "'Total");
    }

    [Test]
    public async Task Changing_the_label_makes_the_save_write_the_sheet()
    {
        using var source = Source(label: "Total");

        using var saved = source.LoadAndSave(wb =>
            wb.Worksheet("Data").Table("Sales").Field("Name").TotalsRowLabel = "Sum");

        await Assert.That(saved.IsWrittenFromModel(Sheet1)).IsTrue();
        await AssertTotals(saved, "Sum");
    }

    [Test]
    public async Task Giving_a_field_a_totals_function_makes_the_save_write_the_sheet()
    {
        using var source = Source(label: "Total");

        using var saved = source.LoadAndSave(wb =>
            wb.Worksheet("Data").Table("Sales").Field("Amount").TotalsRowFunction = XLTotalsRowFunction.Sum);

        await Assert.That(saved.IsWrittenFromModel(Sheet1)).IsTrue();
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("Data").Cell("B4").FormulaA1).Contains("SUBTOTAL");
    }

    [Test]
    public async Task Hiding_the_totals_row_makes_the_save_write_the_sheet()
    {
        using var source = Source(label: "Total");

        using var saved = source.LoadAndSave(wb => wb.Worksheet("Data").Table("Sales").ShowTotalsRow = false);

        await Assert.That(saved.IsWrittenFromModel(Sheet1)).IsTrue();
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("Data").Table("Sales").ShowTotalsRow).IsFalse();
    }

    [Test]
    public async Task Removing_the_table_keeps_the_cells_with_the_label_as_a_value()
    {
        // Without the table, the writer writes the label cell as any other cell, as the file has it.
        using var source = Source(label: "Total");

        using var saved = source.LoadAndSave(wb => wb.Worksheet("Data").Tables.Remove("Sales"));

        await Assert.That(saved.HasKeptCells(Sheet1)).IsTrue();
        await AssertValues(saved);
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("Data").Tables.Count()).IsEqualTo(0);
        await Assert.That(reloaded.Worksheet("Data").Cell("A4").GetText()).IsEqualTo("Total");
    }

    [Test]
    public async Task An_unchanged_sheet_with_a_totals_cell_whose_field_has_no_label_keeps_its_part()
    {
        // The table part gives the field no label, though its totals cell holds text. Until #742 the
        // writer left such a cell out, so the save wrote the sheet; it now writes the cell's value.
        using var source = Source(label: "Total");
        source.RewritePart(TablePart(source), xml => Regex.Replace(xml, " totalsRowLabel=\"Total\"", string.Empty));

        using var saved = source.LoadAndSave();

        await Assert.That(saved.IsKeptWhole(Sheet1)).IsTrue();
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("Data").Cell("A4").GetText()).IsEqualTo("Total");
        await Assert.That(reloaded.Worksheet("Data").Table("Sales").Field("Name").TotalsRowLabel).IsNull();
    }

    [Test]
    public async Task An_unchanged_sheet_with_a_totals_cell_that_differs_from_its_label_keeps_its_part()
    {
        // Until #742 the load wrote the table's label into the cell, so the save wrote the sheet. The
        // load now leaves the cell as the file has it, as Excel shows it, and the save keeps both.
        using var source = Source(label: "Total");
        source.RewritePart(TablePart(source), xml => xml.Replace("totalsRowLabel=\"Total\"", "totalsRowLabel=\"Sum\"",
            StringComparison.Ordinal));

        using var saved = source.LoadAndSave();

        await Assert.That(saved.IsKeptWhole(Sheet1)).IsTrue();
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("Data").Cell("A4").GetText()).IsEqualTo("Total");
        await Assert.That(reloaded.Worksheet("Data").Table("Sales").Field("Name").TotalsRowLabel).IsEqualTo("Sum");
    }

    [Test]
    public async Task An_unchanged_sheet_with_an_inline_string_label_keeps_its_part()
    {
        // Until #742 the writer wrote a label as a shared string, and threw for an inline string, so
        // the save wrote the sheet. It now writes the cell's value, and keeps the file's cells. The
        // shared string the cell no longer uses goes too, and Other's text moves up into its index: a
        // text no cell uses ahead of a used one refuses the cells of every sheet.
        using var source = Source(label: "Total");
        source.RewriteSheet1(xml => Regex.Replace(xml, "<x:c r=\"A4\"([^>]*) t=\"s\"><x:v>\\d+</x:v></x:c>",
            "<x:c r=\"A4\"$1 t=\"inlineStr\"><x:is><x:t>Total</x:t></x:is></x:c>"));
        source.RewritePart("xl/sharedStrings.xml", xml => xml.Replace("<x:si><x:t>Total</x:t></x:si>", string.Empty,
            StringComparison.Ordinal));
        source.RewritePart("xl/worksheets/sheet2.xml", xml => xml.Replace("<x:v>5</x:v>", "<x:v>4</x:v>",
            StringComparison.Ordinal));
        await Assert.That(source.Sheet1Xml()).Contains("inlineStr");
        await Assert.That(source.ReadPart("xl/sharedStrings.xml")).DoesNotContain(">Total<");

        using var saved = source.LoadAndSave();

        await Assert.That(saved.IsKeptWhole(Sheet1)).IsTrue();
        await Assert.That(saved.Sheet1Xml()).Contains("inlineStr");
        await AssertTotals(saved, "Total");
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("Other").Cell("A1").GetText()).IsEqualTo("Other");
    }

    /// <summary>
    /// Sheet Data holds the table Sales over A1:B4: Name and Amount, ("a", 1.5) and ("b", 2), and a
    /// totals row whose Name cell has <paramref name="label"/>. Amount has no totals function. Data's
    /// part is marked by <see cref="KeptSheet.Mark"/>.
    /// </summary>
    private static MemoryStream Source(string? label)
    {
        var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Name";
            ws.Cell("B1").Value = "Amount";
            ws.Cell("A2").Value = "a";
            ws.Cell("B2").Value = 1.5;
            ws.Cell("A3").Value = "b";
            ws.Cell("B3").Value = 2;

            var table = ws.Range("A1:B3").CreateTable("Sales");
            table.ShowTotalsRow = true;
            if (label is not null)
                table.Field("Name").TotalsRowLabel = label;

            wb.AddWorksheet("Other").Cell("A1").Value = "Other";
            wb.SaveAs(stream);
        }

        return stream.Mark(Sheet1);
    }

    /// <summary>
    /// <see cref="Source"/> with the totals cell A4 holding <c>'Total</c> under the quote prefix
    /// flag, and the table part giving the same text as the field's label.
    /// </summary>
    private static MemoryStream QuotedLabelSource()
    {
        // The setter takes ''Total as the quote prefix and 'Total, but keeps ''Total as the label.
        var source = Source(label: "''Total");
        source.RewritePart(TablePart(source), xml => xml.Replace("totalsRowLabel=\"''Total\"",
            "totalsRowLabel=\"'Total\"", StringComparison.Ordinal));
        return source;
    }

    private static string TablePart(Stream package) => package.PartsUnder("xl/tables/").Single();

    private static async Task AssertValues(MemoryStream package)
    {
        package.Position = 0;
        using var wb = new XLWorkbook(package);
        var ws = wb.Worksheet("Data");

        await Assert.That(ws.Cell("A1").GetText()).IsEqualTo("Name");
        await Assert.That(ws.Cell("B2").GetDouble()).IsEqualTo(1.5);
        await Assert.That(ws.Cell("A3").GetText()).IsEqualTo("b");
    }

    private static async Task AssertTotals(MemoryStream package, string label)
    {
        package.Position = 0;
        using var wb = new XLWorkbook(package);
        var ws = wb.Worksheet("Data");

        await Assert.That(ws.Cell("A4").GetText()).IsEqualTo(label);
        await Assert.That(ws.Table("Sales").Field("Name").TotalsRowLabel).IsEqualTo(label);
    }
}
