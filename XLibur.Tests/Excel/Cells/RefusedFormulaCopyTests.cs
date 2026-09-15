using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;

namespace XLibur.Tests.Excel.Cells;

/// <summary>
/// #508: copying a formula the parser refuses keeps its text exactly as it is. Its references are
/// unknown, so there is nothing to move (ADR 0002). Every copy converted the formula to R1C1 at the
/// source and back to A1 at the target, and the conversion threw <see cref="ExpressionParseException"/>
/// on a refused formula, so the copy failed.
/// </summary>
/// <remarks>
/// Each refused text holds a relative reference, or looks as if it does, so a copy that moved it by
/// the offset, or guessed at it with a regex, would change it.
/// </remarks>
public class RefusedFormulaCopyTests
{
    /// <summary>Text the parser cannot read, which <c>FormulaA1</c> stores all the same.</summary>
    private const string Incomplete = "1+";

    /// <summary>An unclosed call around a relative reference.</summary>
    private const string Unclosed = "SUM(A1";

    /// <summary>
    /// The formula bar shows an external reference as <c>'[file.xlsx]Sheet'!A1</c>, but a file stores
    /// it as <c>[1]Sheet!A1</c>, and the parser refuses the displayed form.
    /// </summary>
    private const string FormulaBarExternal = "'[file.xlsx]Sheet'!A1+B2";

    [Test]
    [Arguments(Incomplete)]
    [Arguments(Unclosed)]
    [Arguments(FormulaBarExternal)]
    public async Task Cell_CopyTo_a_cell_keeps_the_text(string refused)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C2").FormulaA1 = refused;

        ws.Cell("C2").CopyTo(ws.Cell("E5"));

        await Assert.That(ws.Cell("E5").FormulaA1).IsEqualTo(refused);
        await Assert.That(ws.Cell("C2").FormulaA1).IsEqualTo(refused);
    }

    [Test]
    [Arguments(Incomplete)]
    [Arguments(Unclosed)]
    [Arguments(FormulaBarExternal)]
    public async Task Cell_CopyTo_an_address_on_another_sheet_keeps_the_text(string refused)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Source");
        var other = wb.AddWorksheet("Other");
        ws.Cell("C2").FormulaA1 = refused;

        ws.Cell("C2").CopyTo("Other!E5");

        await Assert.That(other.Cell("E5").FormulaA1).IsEqualTo(refused);
    }

    [Test]
    [Arguments(Incomplete)]
    [Arguments(Unclosed)]
    [Arguments(FormulaBarExternal)]
    public async Task Cell_CopyFrom_a_cell_keeps_the_text(string refused)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C2").FormulaA1 = refused;

        ws.Cell("E5").CopyFrom(ws.Cell("C2"));

        await Assert.That(ws.Cell("E5").FormulaA1).IsEqualTo(refused);
    }

    [Test]
    [Arguments(Incomplete)]
    [Arguments(Unclosed)]
    [Arguments(FormulaBarExternal)]
    public async Task Cell_CopyFrom_a_range_keeps_the_text(string refused)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C2").FormulaA1 = refused;
        ws.Cell("D2").Value = 7;

        ws.Cell("E5").CopyFrom(ws.Range("C2:D2"));

        await Assert.That(ws.Cell("E5").FormulaA1).IsEqualTo(refused);
        await Assert.That(ws.Cell("F5").Value.GetNumber()).IsEqualTo(7.0);
    }

    [Test]
    [Arguments(Incomplete)]
    [Arguments(Unclosed)]
    [Arguments(FormulaBarExternal)]
    public async Task Range_CopyTo_a_cell_on_another_sheet_keeps_the_text(string refused)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Source");
        var other = wb.AddWorksheet("Other");
        ws.Cell("C2").FormulaA1 = refused;
        ws.Cell("C3").FormulaA1 = "C2+A1";

        ws.Range("C2:C3").CopyTo(other.Cell("E5"));

        await Assert.That(other.Cell("E5").FormulaA1).IsEqualTo(refused);
        await Assert.That(other.Cell("E6").FormulaA1).IsEqualTo("E5+C4");
    }

    [Test]
    [Arguments(Incomplete)]
    [Arguments(Unclosed)]
    [Arguments(FormulaBarExternal)]
    public async Task Range_CopyTo_a_range_keeps_the_text(string refused)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C2").FormulaA1 = refused;

        ws.Range("C2:D3").CopyTo(ws.Range("E5:F6"));

        await Assert.That(ws.Cell("E5").FormulaA1).IsEqualTo(refused);
    }

    [Test]
    [Arguments(Incomplete)]
    [Arguments(Unclosed)]
    [Arguments(FormulaBarExternal)]
    public async Task Row_CopyTo_a_row_keeps_the_text(string refused)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B1").FormulaA1 = refused;

        ws.Row(1).CopyTo(ws.Row(4));

        await Assert.That(ws.Cell("B4").FormulaA1).IsEqualTo(refused);
    }

    [Test]
    [Arguments(Incomplete)]
    [Arguments(Unclosed)]
    [Arguments(FormulaBarExternal)]
    public async Task Worksheet_CopyTo_keeps_the_text(string refused)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Source");
        ws.Cell("C2").FormulaA1 = refused;

        var copy = ws.CopyTo("Copy");

        await Assert.That(copy.Cell("C2").FormulaA1).IsEqualTo(refused);
    }

    [Test]
    [Arguments(Incomplete)]
    [Arguments(Unclosed)]
    [Arguments(FormulaBarExternal)]
    public async Task Transpose_keeps_the_text(string refused)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B1").FormulaA1 = refused;
        ws.Cell("A2").Value = 5;

        ws.Range("A1:B2").Transpose(XLTransposeOptions.ReplaceCells);

        await Assert.That(ws.Cell("A2").FormulaA1).IsEqualTo(refused);
        await Assert.That(ws.Cell("B1").Value.GetNumber()).IsEqualTo(5.0);
    }

    /// <summary>
    /// The copy is still the refused formula: reading its value, or its R1C1 text, throws as reading
    /// the source does. Only the copy stopped throwing.
    /// </summary>
    [Test]
    public async Task The_copy_is_still_a_refused_formula()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C2").FormulaA1 = Unclosed;

        ws.Cell("C2").CopyTo(ws.Cell("E5"));

        await Assert.That(ws.Cell("E5").HasFormula).IsTrue();
        await Assert.That(ws.Cell("E5").NeedsRecalculation).IsTrue();
        await Assert.That(() => ws.Cell("E5").Value).Throws<ExpressionParseException>();
        await Assert.That(() => ws.Cell("E5").FormulaR1C1).Throws<ExpressionParseException>();
    }

    /// <summary>A formula the parser reads still moves its relative references by the offset.</summary>
    [Test]
    public async Task A_formula_the_parser_reads_still_moves_by_the_offset()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C2").FormulaA1 = "A1+$B$2+SUM(A1:B1)";

        ws.Cell("C2").CopyTo(ws.Cell("E5"));

        await Assert.That(ws.Cell("E5").FormulaA1).IsEqualTo("C4+$B$2+SUM(C4:D4)");
    }

    [Test]
    [Arguments(Incomplete)]
    [Arguments(Unclosed)]
    [Arguments(FormulaBarExternal)]
    public async Task A_data_validation_formula_keeps_the_text_when_its_cell_is_copied(string refused)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C2").CreateDataValidation().Custom(refused);

        ws.Cell("C2").CopyTo(ws.Cell("E5"));

        await Assert.That(ws.Cell("E5").GetDataValidation().Value).IsEqualTo(refused);
        await Assert.That(ws.Cell("C2").GetDataValidation().Value).IsEqualTo(refused);
    }

    [Test]
    [Arguments(Incomplete)]
    [Arguments(Unclosed)]
    [Arguments(FormulaBarExternal)]
    public async Task A_conditional_format_formula_keeps_the_text_when_its_range_is_copied(string refused)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C2").Value = 1;
        ws.Range("C2:C3").AddConditionalFormat().WhenIsTrue(refused).Fill.SetBackgroundColor(XLColor.Red);

        ws.Range("C2:C3").CopyTo(ws.Cell("E5"));

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(2);
        await Assert.That(ws.ConditionalFormats.Select(cf => cf.Values[1].Value))
            .IsEquivalentTo(new[] { refused, refused });
    }

    /// <summary>
    /// D78 / #508, pre-existing behaviour, kept on purpose. A data table's formula text is only a
    /// placeholder, such as <c>{TABLE(A1,B1}</c>, which the parser always refuses. XLibur has no model
    /// of what a copy of a data table becomes, so copying one still throws, as it did before #508.
    /// Copied as a refused normal formula, the placeholder would lose the table's type, range and
    /// inputs, a save would write it as <c>&lt;f&gt;{TABLE(A1,B1}&lt;/f&gt;</c>, and Excel would report
    /// unreadable content.
    /// </summary>
    [Test]
    public async Task Copying_a_data_table_still_throws_as_before()
    {
        using var stream = TestHelper.GetStreamFromResource(
            TestHelper.GetResourcePath(@"Other\Formulas\DataTableFormula-Excel-Input.xlsx"));
        using var wb = new XLWorkbook(stream);
        var master = wb.Worksheets
            .SelectMany(ws => ws.CellsUsed())
            .Cast<XLCell>()
            .First(c => c.Formula is { Type: FormulaType.DataTable });
        var sheet = master.Worksheet;

        await Assert.That(() => master.CopyTo(sheet.Cell(200, 20))).Throws<ExpressionParseException>();
        await Assert.That(() => master.AsRange().CopyTo(sheet.Cell(300, 20))).Throws<ExpressionParseException>();
        await Assert.That(() => sheet.CopyTo("Copy")).Throws<ExpressionParseException>();
    }

    /// <summary>
    /// Pre-existing behaviour, kept on purpose: only a normal formula whose text the parser refuses is
    /// copied as it is (#508). A refused array formula still makes the copy throw, as before.
    /// </summary>
    [Test]
    public async Task Copying_a_refused_array_formula_still_throws_as_before()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("C2:C3").FormulaArrayA1 = Unclosed;

        await Assert.That(() => ws.Cell("C2").CopyTo(ws.Cell("E5"))).Throws<ExpressionParseException>();
    }

    /// <summary>
    /// Pre-existing behaviour, unchanged by #508: a cell of an array formula the parser reads is copied
    /// as a normal formula, with its references moved by the offset.
    /// </summary>
    [Test]
    public async Task An_array_formula_the_parser_reads_is_copied_as_a_normal_formula_as_before()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("C2:C3").FormulaArrayA1 = "A1:A2*2";

        ws.Cell("C2").CopyTo(ws.Cell("E5"));

        await Assert.That(ws.Cell("E5").FormulaA1).IsEqualTo("C4:C5*2");
        await Assert.That(ws.Cell("E5").HasArrayFormula).IsFalse();
    }
}
