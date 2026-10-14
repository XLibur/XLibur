using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Tables;

/// <summary>
/// A cell in a table's totals row that has no formula is a cell like any other: the load reads it and
/// the save writes it as the file had it, whatever the field's label in the table part says (#742).
/// </summary>
/// <remarks>
/// Excel shows the cell, not the label (checked over COM). A file whose totals cell holds text and
/// whose field has no label opens as it is, and so does one whose totals cell is blank under a label.
/// </remarks>
public class TotalsRowCellTests
{
    private const string A4 = "<x:c r=\"A4\"([^>]*?) t=\"s\"><x:v>\\d+</x:v></x:c>";

    [Test]
    [Arguments(false)]
    [Arguments(true)]
    public async Task Text_in_the_totals_cell_of_a_field_with_no_label_is_kept(bool rewrite)
    {
        using var source = Source(label: "my note");
        source.RewritePart(TablePart(source), xml => xml.Replace(" totalsRowLabel=\"my note\"", string.Empty,
            StringComparison.Ordinal));

        using var saved = source.LoadAndSave(options: Options(rewrite));

        using var wb = new XLWorkbook(saved);
        var ws = wb.Worksheet("Data");
        await Assert.That(ws.Cell("A4").GetText()).IsEqualTo("my note");
        await Assert.That(ws.Table("Sales").Field("Name").TotalsRowLabel).IsNull();
    }

    [Test]
    public async Task The_load_keeps_a_number_in_the_totals_cell_under_a_label()
    {
        using var source = NumberUnderLabelSource();

        using var wb = new XLWorkbook(source);
        var ws = wb.Worksheet("Data");

        await Assert.That(ws.Cell("A4").GetDouble()).IsEqualTo(5);
        await Assert.That(ws.Table("Sales").Field("Name").TotalsRowLabel).IsEqualTo("Total");
    }

    [Test]
    [Arguments(false)]
    [Arguments(true)]
    public async Task The_save_keeps_a_number_in_the_totals_cell_under_a_label(bool rewrite)
    {
        using var source = NumberUnderLabelSource();

        using var saved = source.LoadAndSave(options: Options(rewrite));

        using var wb = new XLWorkbook(saved);
        var ws = wb.Worksheet("Data");
        await Assert.That(ws.Cell("A4").GetDouble()).IsEqualTo(5);
        await Assert.That(ws.Table("Sales").Field("Name").TotalsRowLabel).IsEqualTo("Total");
    }

    [Test]
    [Arguments(false)]
    [Arguments(true)]
    public async Task An_inline_string_label_in_the_totals_cell_is_saved(bool rewrite)
    {
        using var source = Source(label: "Total");
        source.RewriteSheet1(xml => Regex.Replace(xml, A4,
            "<x:c r=\"A4\"$1 t=\"inlineStr\"><x:is><x:t>Total</x:t></x:is></x:c>"));
        await Assert.That(source.Sheet1Xml()).Contains("inlineStr");

        using var saved = source.LoadAndSave(options: Options(rewrite));

        using var wb = new XLWorkbook(saved);
        var ws = wb.Worksheet("Data");
        await Assert.That(ws.Cell("A4").GetText()).IsEqualTo("Total");
        await Assert.That(ws.Table("Sales").Field("Name").TotalsRowLabel).IsEqualTo("Total");
    }

    [Test]
    [Arguments(false)]
    [Arguments(true)]
    public async Task A_blank_totals_cell_with_a_comment_under_a_label_is_saved(bool rewrite)
    {
        using var source = Source(label: "Total", comment: "note");
        source.RewriteSheet1(xml => Regex.Replace(xml, A4, "<x:c r=\"A4\"$1 />"));

        using var saved = source.LoadAndSave(options: Options(rewrite));

        using var wb = new XLWorkbook(saved);
        var ws = wb.Worksheet("Data");
        await Assert.That(ws.Cell("A4").Value.IsBlank).IsTrue();
        await Assert.That(ws.Cell("A4").GetComment().Text).IsEqualTo("note");
        await Assert.That(ws.Table("Sales").Field("Name").TotalsRowLabel).IsEqualTo("Total");
    }

    [Test]
    public async Task Setting_a_label_on_a_loaded_table_writes_it_into_the_cell()
    {
        using var source = NumberUnderLabelSource();

        using var saved = source.LoadAndSave(wb =>
            wb.Worksheet("Data").Table("Sales").Field("Name").TotalsRowLabel = "Sum");

        using var reloaded = new XLWorkbook(saved);
        var ws = reloaded.Worksheet("Data");
        await Assert.That(ws.Cell("A4").GetText()).IsEqualTo("Sum");
        await Assert.That(ws.Table("Sales").Field("Name").TotalsRowLabel).IsEqualTo("Sum");
    }

    [Test]
    public async Task Setting_a_totals_function_on_a_loaded_table_writes_its_formula()
    {
        using var source = NumberUnderLabelSource();

        using var saved = source.LoadAndSave(wb =>
            wb.Worksheet("Data").Table("Sales").Field("Name").TotalsRowFunction = XLTotalsRowFunction.Count);

        using var reloaded = new XLWorkbook(saved);
        var ws = reloaded.Worksheet("Data");
        await Assert.That(ws.Cell("A4").FormulaA1).Contains("SUBTOTAL(103");
        await Assert.That(ws.Table("Sales").Field("Name").TotalsRowFunction).IsEqualTo(XLTotalsRowFunction.Count);
    }

    private static SaveOptions Options(bool rewrite) => new() { RewriteUnchangedSheets = rewrite };

    /// <summary>
    /// <see cref="Source"/> with the totals cell A4 holding the number 5, under the label Total.
    /// </summary>
    private static MemoryStream NumberUnderLabelSource()
    {
        var source = Source(label: "Total");
        source.RewriteSheet1(xml => Regex.Replace(xml, A4, "<x:c r=\"A4\"$1><x:v>5</x:v></x:c>"));
        return source;
    }

    /// <summary>
    /// Sheet Data holds the table Sales over A1:B4: Name and Amount, ("a", 1.5) and ("b", 2), and a
    /// totals row whose Name cell A4 has <paramref name="label"/> as the field's label and as its text.
    /// </summary>
    private static MemoryStream Source(string label, string? comment = null)
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
            table.Field("Name").TotalsRowLabel = label;
            if (comment is not null)
                ws.Cell("A4").CreateComment().AddText(comment);

            wb.SaveAs(stream);
        }

        stream.Position = 0;
        return stream;
    }

    private static string TablePart(Stream package) => package.PartsUnder("xl/tables/").Single();
}
