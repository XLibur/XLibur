using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Tables;

/// <summary>
/// The load reads a totals-row formula and leaves it as the file has it: it does not write the formula
/// of the column's totals function over it, which made the formula dirty and lost its cached value on
/// save (#749).
/// </summary>
public class TotalsRowFormulaLoadTests
{
    private const string LineBreaks = "TryToLoad.TableHeadersWithLineBreaks.xlsx";

    [Test]
    public async Task The_load_leaves_a_totals_function_formula_clean_and_as_the_file_has_it()
    {
        using var wb = new XLWorkbook(Resource(LineBreaks));
        var ws = wb.Worksheet(1);

        await AssertLoaded(ws.Cell("H7"), "SUBTOTAL(109,Table1[Purchase\n price])", 33059);
        await AssertLoaded(ws.Cell("I7"), "SUBTOTAL(109,Table1[Estimated\ncurrent value])", 35156.490000000005);
        await AssertLoaded(ws.Cell("K7"), "SUBTOTAL(103,Table1[Photo?])", 5);

        var table = ws.Table("Table1");
        await Assert.That(table.Field("Purchase\n price").TotalsRowFunction).IsEqualTo(XLTotalsRowFunction.Sum);
        await Assert.That(table.Field("Photo?").TotalsRowFunction).IsEqualTo(XLTotalsRowFunction.Count);
    }

    [Test]
    [Arguments(false)]
    [Arguments(true)]
    public async Task A_totals_formula_keeps_its_cached_value_through_a_load_and_save(bool rewrite)
    {
        using var source = Resource(LineBreaks);

        using var saved = source.LoadAndSave(options: new SaveOptions { RewriteUnchangedSheets = rewrite });

        using var wb = new XLWorkbook(saved);
        var ws = wb.Worksheet(1);
        await AssertLoaded(ws.Cell("H7"), "SUBTOTAL(109,Table1[Purchase\n price])", 33059);
        await AssertLoaded(ws.Cell("I7"), "SUBTOTAL(109,Table1[Estimated\ncurrent value])", 35156.490000000005);
        await AssertLoaded(ws.Cell("K7"), "SUBTOTAL(103,Table1[Photo?])", 5);
    }

    [Test]
    public async Task The_load_leaves_a_custom_totals_formula_clean()
    {
        using var source = Source(field => field.TotalsRowFormulaA1 = "SUM([Amount])*2");

        using var wb = new XLWorkbook(source);
        var field = wb.Worksheet("Data").Table("Sales").Field("Amount");

        await AssertLoaded(wb.Worksheet("Data").Cell("B4"), "SUM([Amount])*2", 7);
        await Assert.That(field.TotalsRowFunction).IsEqualTo(XLTotalsRowFunction.Custom);
        await Assert.That(field.TotalsRowLabel).IsNull();
    }

    [Test]
    public async Task The_load_writes_the_formula_of_a_totals_function_into_a_totals_cell_without_one()
    {
        // Not a file Excel writes, but the load has always given such a cell its function's formula.
        using var source = Source(field => field.TotalsRowFunction = XLTotalsRowFunction.Sum);
        source.RewriteSheet1(xml => Regex.Replace(xml, "<x:c r=\"B4\"([^>]*)>.*?</x:c>", "<x:c r=\"B4\"$1 />"));
        await Assert.That(source.Sheet1Xml()).DoesNotContain("SUBTOTAL");

        using var wb = new XLWorkbook(source);

        await Assert.That(wb.Worksheet("Data").Cell("B4").FormulaA1).IsEqualTo("SUBTOTAL(109,[Amount])");
        await Assert.That(wb.Worksheet("Data").Cell("B4").Value).IsEqualTo(3.5);
    }

    [Test]
    public async Task A_totals_formula_loaded_without_a_cached_value_is_calculated()
    {
        // The load keeps the formula; one with no cached value still loads dirty, as any formula does.
        using var source = Source(field => field.TotalsRowFunction = XLTotalsRowFunction.Sum);
        source.RewriteSheet1(xml => Regex.Replace(xml, "(<x:c r=\"B4\"[^>]*>.*?</x:f>)<x:v>[^<]*</x:v>", "$1"));
        await Assert.That(source.Sheet1Xml()).Contains("SUBTOTAL(109,[Amount])</x:f></x:c>");

        using var wb = new XLWorkbook(source);
        var cell = wb.Worksheet("Data").Cell("B4");

        await Assert.That(cell.NeedsRecalculation).IsTrue();
        await Assert.That(cell.Value).IsEqualTo(3.5);
    }

    [Test]
    [Arguments("Purchase\n price")]
    [Arguments("Estimated\ncurrent value")]
    [Arguments("Photo?")]
    public async Task A_structured_reference_resolves_a_column_with_a_line_break_or_a_question_mark(string column)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = column;
        ws.Cell("A2").Value = 1.5;
        ws.Cell("A3").Value = 2;
        ws.Range("A1:A3").CreateTable("T");

        await Assert.That(ws.Evaluate($"SUM(T[{column}])")).IsEqualTo(3.5);
        await Assert.That(ws.Evaluate($"SUM(T[[{column}]])")).IsEqualTo(3.5);
        await Assert.That(ws.Evaluate($"SUM(T[[#Data],[{column}]])")).IsEqualTo(3.5);
    }

    private static async Task AssertLoaded(IXLCell cell, string formula, double cached)
    {
        await Assert.That(cell.FormulaA1).IsEqualTo(formula);
        await Assert.That(cell.NeedsRecalculation).IsFalse();
        await Assert.That(cell.CachedValue).IsEqualTo(cached);
    }

    private static MemoryStream Resource(string name)
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(name));
        var copy = new MemoryStream();
        stream.CopyTo(copy);
        copy.Position = 0;
        return copy;
    }

    /// <summary>
    /// Sheet Data holds the table Sales over A1:B4: Name and Amount, ("a", 1.5) and ("b", 2), and a
    /// totals row whose Amount cell B4 <paramref name="setTotals"/> fills, saved with its value.
    /// </summary>
    private static MemoryStream Source(Action<IXLTableField> setTotals)
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
            setTotals(table.Field("Amount"));

            wb.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = true });
        }

        stream.Position = 0;
        return stream;
    }
}
