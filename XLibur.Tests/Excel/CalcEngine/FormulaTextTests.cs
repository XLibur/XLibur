using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Parser;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.CalcEngine.Visitors;
using XLibur.Excel.ConditionalFormats;
using XLibur.Excel.Coordinates;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// Formula text reaches the parser through one module, and a refused formula means one thing to
/// each caller. The defects are the ones spec 54 was written for: each path that hands formula text
/// to the parser used to repeat the steps around the call, and each got a different one wrong.
/// </summary>
public class FormulaTextTests
{
    /// <summary>
    /// The form the formula bar shows for an external reference. A file stores <c>[1]Sheet1!A1</c>;
    /// <see cref="IXLCell.FormulaA1"/> accepts this form, and the parser refuses it.
    /// </summary>
    private const string RefusedExternalReference = "'[Book2.xlsx]Sheet1'!A1";
    /// <summary>
    /// D49. A rename used to reach the refused formula part way through, after the calc engine had
    /// already renamed the sheet, and throw the parser's own exception: the sheet kept its old name,
    /// and a formula that referred to it evaluated to <c>#REF!</c>.
    /// </summary>
    [Test]
    public async Task D49_renaming_a_sheet_completes_and_leaves_a_refused_formula_unchanged()
    {
        using var wb = new XLWorkbook();
        var sheet1 = wb.AddWorksheet("Sheet1");
        var sheet2 = wb.AddWorksheet("Sheet2");
        sheet2.Cell("A1").FormulaA1 = RefusedExternalReference;
        sheet2.Cell("B1").FormulaA1 = "Sheet1!A1+1";
        sheet1.Cell("A1").Value = 5;
        await Assert.That(sheet2.Cell("B1").Value).IsEqualTo(6);

        sheet1.Name = "Data";

        await Assert.That(sheet1.Name).IsEqualTo("Data");
        await Assert.That(sheet2.Cell("A1").FormulaA1).IsEqualTo(RefusedExternalReference);
        await Assert.That(sheet2.Cell("B1").FormulaA1).IsEqualTo("Data!A1+1");
        await Assert.That(sheet2.Cell("B1").Value).IsEqualTo(6);
    }

    /// <summary>
    /// D49's shape was a rename that changed some holders of the sheet name and then threw before the
    /// rest. After a rename past a refused formula, every holder has the new name: the sheet, the
    /// workbook's lookup by name, and the calc engine.
    /// </summary>
    [Test]
    public async Task After_a_rename_past_a_refused_formula_the_sheet_the_workbook_and_the_calc_engine_agree()
    {
        using var wb = new XLWorkbook();
        var sheet1 = wb.AddWorksheet("Sheet1");
        var sheet2 = wb.AddWorksheet("Sheet2");
        sheet2.Cell("A1").FormulaA1 = RefusedExternalReference;
        sheet2.Cell("B1").FormulaA1 = "Sheet1!A1+1";
        sheet1.Cell("A1").Value = 5;
        await Assert.That(sheet2.Cell("B1").Value).IsEqualTo(6);

        sheet1.Name = "Data";

        await Assert.That(sheet1.Name).IsEqualTo("Data");
        await Assert.That(wb.Worksheet("Data")).IsSameReferenceAs(sheet1);
        await Assert.That(wb.TryGetWorksheet("Sheet1", out IXLWorksheet? _)).IsFalse();

        // The calc engine agrees: a change on the renamed sheet reaches the formula that refers to it.
        sheet1.Cell("A1").Value = 10;
        await Assert.That(sheet2.Cell("B1").Value).IsEqualTo(11);
    }

    /// <summary>
    /// D50. Evaluation stripped the future-function prefix only when it was written in lower case.
    /// </summary>
    [Test]
    public async Task D50_an_upper_case_future_function_prefix_evaluates()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaA1 = "_XLFN.CONCAT(\"a\",\"b\")";

        await Assert.That(ws.Cell("A1").Value).IsEqualTo("ab");
    }

    /// <summary>
    /// D51. The check that stops SUBTOTAL counting a nested SUBTOTAL parsed every candidate cell and
    /// let the parser's exception out, about a cell the caller never read. A refused formula does not
    /// call SUBTOTAL as far as anyone can tell, so its cell counts.
    /// </summary>
    /// <remarks>
    /// A workbook Excel wrote holds a cached value beside a formula the parser refuses, and a load
    /// leaves the cell in that state: the formula clean and its cached value in the cell. SUM already
    /// counts that value.
    /// </remarks>
    [Test]
    public async Task D51_SUBTOTAL_counts_a_cell_whose_formula_is_refused()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var a1 = (XLCell)ws.Cell("A1");
        a1.FormulaA1 = RefusedExternalReference + "+SUBTOTAL(9,B5)";
        ((XLWorksheet)ws).Internals.CellsCollection.ValueSlice.SetCellValue(a1.SheetPoint, 10);
        a1.Formula!.MarkClean();
        ws.Cell("A2").Value = 2;
        ws.Cell("C1").FormulaA1 = "SUM(A1:A2)";
        ws.Cell("C2").FormulaA1 = "SUBTOTAL(9,A1:A2)";

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(12);
        await Assert.That(ws.Cell("C2").Value).IsEqualTo(12);
    }

    /// <summary>
    /// D51, when the refused formula has no value yet. Its cell counts, so SUBTOTAL fails the way SUM
    /// over the same cell fails: with XLibur's <see cref="ExpressionParseException"/>, not the
    /// parser's own exception.
    /// </summary>
    [Test]
    public async Task D51_SUBTOTAL_over_a_dirty_refused_formula_fails_as_SUM_does()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaA1 = RefusedExternalReference + "+SUBTOTAL(9,B5)";
        ws.Cell("A2").Value = 2;
        ws.Cell("C1").FormulaA1 = "SUBTOTAL(9,A1:A2)";
        ws.Cell("C2").FormulaA1 = "SUM(A1:A2)";

        await Assert.That(() => ws.Cell("C1").Value).Throws<ExpressionParseException>();
        await Assert.That(() => ws.Cell("C2").Value).Throws<ExpressionParseException>();
    }

    /// <summary>
    /// D52. Evaluation did not hide the colon in a column name from the parser, which read it as a
    /// range operator. The text survived a save and a load; only its value was wrong.
    /// </summary>
    [Test]
    public async Task D52_a_table_column_whose_name_contains_a_colon_evaluates()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "Name";
            ws.Cell("B1").Value = "Start: Date";
            ws.Cell("A2").Value = "a";
            ws.Cell("B2").Value = 3;
            ws.Cell("A3").Value = "b";
            ws.Cell("B3").Value = 4;
            ws.Range("A1:B3").CreateTable("Table1");
            ws.Cell("D1").FormulaA1 = "SUM(Table1[Start: Date])";

            await Assert.That(ws.Cell("D1").Value).IsEqualTo(7);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheet("Sheet1");
            wb.RecalculateAllFormulas();

            await Assert.That(ws.Cell("D1").FormulaA1).IsEqualTo("SUM(Table1[Start: Date])");
            await Assert.That(ws.Cell("D1").Value).IsEqualTo(7);
        }
    }

    /// <summary>
    /// Spec 54, task 1.3: the review read that a conditional format's comparer runs the R1C1
    /// converter, which throws on a refused formula. Executed, it did: the format could not go in a
    /// <see cref="HashSet{T}"/>. It can now, and the set finds it again.
    /// </summary>
    [Test]
    public async Task A_conditional_format_holding_a_refused_formula_can_be_put_in_a_HashSet_and_found_again()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var format = new XLConditionalFormat((XLRange)ws.Range("A1:A5"));
        format.Values.Add(new XLFormula("=" + RefusedExternalReference));

        var set = new HashSet<IXLConditionalFormat>(XLConditionalFormat.NoRangeComparer) { format };

        await Assert.That(set).Count().IsEqualTo(1);
        await Assert.That(set.Contains(format)).IsTrue();
    }

    /// <summary>
    /// The references in a refused formula are unknown, so nothing says two formats that hold one
    /// hold the same formula. Such a format equals only itself.
    /// </summary>
    [Test]
    public async Task A_conditional_format_holding_a_refused_formula_equals_no_other_format()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var first = new XLConditionalFormat((XLRange)ws.Range("A1:A5"));
        first.Values.Add(new XLFormula("=" + RefusedExternalReference));
        var second = new XLConditionalFormat((XLRange)ws.Range("B1:B5"));
        second.Values.Add(new XLFormula("=" + RefusedExternalReference));
        var comparer = XLConditionalFormat.NoRangeComparer;

        await Assert.That(comparer.Equals(first, second)).IsFalse();
        await Assert.That(comparer.Equals(first, first)).IsTrue();
        await Assert.That(comparer.GetHashCode(first)).IsEqualTo(comparer.GetHashCode(first));
    }

    /// <summary>
    /// Two formats with the same refused text over adjacent ranges would be merged if they compared
    /// equal, and a merge re-points the formula to the merged range's first cell. They are not merged,
    /// and each keeps its range and its text through a save and a reload (ADR 0002).
    /// </summary>
    [Test]
    public async Task Two_formats_with_the_same_refused_formula_over_adjacent_ranges_are_not_merged()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Range("A1:A5").AddConditionalFormat().WhenIsTrue(RefusedExternalReference)
                .Fill.SetBackgroundColor(XLColor.Red);
            ws.Range("B1:B5").AddConditionalFormat().WhenIsTrue(RefusedExternalReference)
                .Fill.SetBackgroundColor(XLColor.Red);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var reloaded = new XLWorkbook(ms);
        var formats = reloaded.Worksheet("Sheet1").ConditionalFormats
            .OrderBy(f => f.Ranges.Single().RangeAddress.ToString())
            .ToList();

        await Assert.That(formats).Count().IsEqualTo(2);
        await Assert.That(formats[0].Ranges.Single().RangeAddress.ToString()).IsEqualTo("A1:A5");
        await Assert.That(formats[1].Ranges.Single().RangeAddress.ToString()).IsEqualTo("B1:B5");
        foreach (var format in formats)
            await Assert.That(format.Values.Single().Value.Value).IsEqualTo(RefusedExternalReference);
    }

    /// <summary>
    /// The control for the test above: the same two formats with a formula the parser accepts are
    /// merged, so "not merged" there is the refusal's doing and not the layout's.
    /// </summary>
    [Test]
    public async Task Two_formats_with_the_same_accepted_formula_over_adjacent_ranges_are_merged()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Range("A1:A5").AddConditionalFormat().WhenIsTrue("$C$1>0").Fill.SetBackgroundColor(XLColor.Red);
            ws.Range("B1:B5").AddConditionalFormat().WhenIsTrue("$C$1>0").Fill.SetBackgroundColor(XLColor.Red);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var reloaded = new XLWorkbook(ms);
        var format = reloaded.Worksheet("Sheet1").ConditionalFormats.Single();

        await Assert.That(format.Ranges.Single().RangeAddress.ToString()).IsEqualTo("A1:B5");
    }

    /// <summary>
    /// The consequence of the comparer that a caller met: a save consolidates conditional formats by
    /// default, and consolidation converted each formula to R1C1, so the save threw.
    /// </summary>
    [Test]
    public async Task A_workbook_whose_conditional_format_holds_a_refused_formula_saves()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Range("A1:A5").AddConditionalFormat().WhenIsTrue(RefusedExternalReference)
            .Fill.SetBackgroundColor(XLColor.Red);
        using var ms = new MemoryStream();

        wb.SaveAs(ms);

        ms.Position = 0;
        using var reloaded = new XLWorkbook(ms);
        var reloadedFormula = reloaded.Worksheet("Sheet1").ConditionalFormats.Single().Values.Single().Value.Value;
        await Assert.That(reloadedFormula).IsEqualTo(RefusedExternalReference);
    }

    // The module on its own. Each operation runs over every corpus row. For text the parser accepts,
    // it must give the answer of the path it replaces. For text the parser refuses, it must return a
    // refusal that carries the text, and leave the text unchanged.

    /// <summary>
    /// The corpus forms the parser refuses. Evaluation takes a leading <c>=</c> off before the parser
    /// sees the text, so <c>=A1+1</c> evaluates; handed to the parser as it is, it is refused.
    /// </summary>
    private static readonly HashSet<string> RefusedForms =
    [
        "external_formula_bar_form", "leading_equals", "empty", "whitespace_only", "unparseable",
        "refused_subtotal_call",
    ];

    [Test]
    [MethodDataSource(typeof(FormulaTextCorpusTests), nameof(FormulaTextCorpusTests.Rows))]
    public async Task TryWalk_refuses_exactly_the_refused_forms(FormulaTextCorpusTests.CorpusRow row)
    {
        var accepted = FormulaText.TryWalk(row.Text, new List<string>(), ColumnProbe.Instance,
            FormulaNotation.A1, out _, out var refusal);

        await Assert.That(accepted).IsEqualTo(!RefusedForms.Contains(row.Form));
        if (!accepted)
        {
            await Assert.That(refusal.Text).IsEqualTo(row.Text);
            await Assert.That(refusal.Cause).IsNotNull();
        }
    }

    /// <summary>
    /// The parser reads a column name with its colon hidden. The factory must still get the name as
    /// the formula writes it, or a table lookup by that name fails.
    /// </summary>
    [Test]
    [MethodDataSource(typeof(FormulaTextCorpusTests), nameof(FormulaTextCorpusTests.Rows))]
    public async Task TryWalk_hands_the_factory_each_column_name_as_it_is_written(
        FormulaTextCorpusTests.CorpusRow row)
    {
        var columns = new List<string>();
        FormulaText.TryWalk(row.Text, columns, ColumnProbe.Instance, FormulaNotation.A1, out _, out _);

        var expected = row.Form switch
        {
            "structured_reference" => "Name",
            "colon_in_column_name" => "Start: Date",
            _ => string.Empty,
        };
        await Assert.That(string.Join("|", columns)).IsEqualTo(expected);
    }

    [Test]
    [MethodDataSource(typeof(FormulaTextCorpusTests), nameof(FormulaTextCorpusTests.Rows))]
    public async Task TryConvert_to_R1C1_gives_what_conversion_gives(FormulaTextCorpusTests.CorpusRow row)
    {
        var converted = FormulaText.TryConvert(row.Text, new Point(3, 3), FormulaNotation.R1C1,
            out var r1c1, out var refusal);

        if (RefusedForms.Contains(row.Form))
        {
            await Assert.That(converted).IsFalse();
            await Assert.That(r1c1).IsEqualTo(row.Text);
            await Assert.That(refusal.Text).IsEqualTo(row.Text);
            return;
        }

        await Assert.That(converted).IsTrue();
        await Assert.That(r1c1).IsEqualTo(row.Cells["to_r1c1"]);
    }

    [Test]
    [MethodDataSource(typeof(FormulaTextCorpusTests), nameof(FormulaTextCorpusTests.Rows))]
    public async Task TryRewrite_gives_what_a_rename_gives(FormulaTextCorpusTests.CorpusRow row)
    {
        var modifier = new RenameRefModVisitor
        {
            Sheets = new Dictionary<string, string?> { { "Sheet1", "Data" } }
        };

        var rewritten = FormulaText.TryRewrite(row.Text, "Data", new Point(2, 8), modifier,
            out var text, out var refusal);

        if (RefusedForms.Contains(row.Form))
        {
            await Assert.That(rewritten).IsFalse();
            await Assert.That(text).IsEqualTo(row.Text);
            await Assert.That(refusal.Text).IsEqualTo(row.Text);
            return;
        }

        await Assert.That(rewritten).IsTrue();
        await Assert.That(text).IsEqualTo(row.Cells["rename"]);
    }

    [Test]
    [MethodDataSource(typeof(FormulaTextCorpusTests), nameof(FormulaTextCorpusTests.Rows))]
    public async Task AddFuturePrefixes_gives_what_setting_a_formula_gives(FormulaTextCorpusTests.CorpusRow row)
    {
        var prefixed = FormulaText.AddFuturePrefixes(row.Text, "Sheet1", new Point(3, 3));

        await Assert.That(prefixed).IsEqualTo(row.Cells["add_prefix"]);
    }

    [Test]
    [Arguments("_xlfn.CONCAT", true, "CONCAT")]
    [Arguments("_XLFN.CONCAT", true, "CONCAT")]
    [Arguments("_Xlfn.concat", true, "concat")]
    [Arguments("_xlfn._xlws.FILTER", true, "FILTER")]
    [Arguments("_XLFN._XLWS.FILTER", true, "FILTER")]
    [Arguments("CONCAT", false, "CONCAT")]
    [Arguments("_xlws.FILTER", false, "_xlws.FILTER")]
    [Arguments("_xlfn", false, "_xlfn")]
    public async Task TryStripFuturePrefix_takes_the_prefix_off_in_any_case(string name, bool stripped, string bare)
    {
        var (actualStripped, actualBare) = Strip(name);

        await Assert.That(actualStripped).IsEqualTo(stripped);
        await Assert.That(actualBare).IsEqualTo(bare);

        static (bool Stripped, string Bare) Strip(string name)
        {
            var result = FormulaText.TryStripFuturePrefix(name.AsSpan(), out var bareName);
            return (result, bareName.ToString());
        }
    }

    [Test]
    [Arguments("=A1+1", "A1+1")]
    [Arguments("A1+1", "A1+1")]
    [Arguments("==A1", "=A1")]
    [Arguments("", "")]
    [Arguments(" =A1", " =A1")]
    public async Task WithoutLeadingEquals_takes_off_one_leading_equals_sign(string text, string expected)
    {
        await Assert.That(FormulaText.WithoutLeadingEquals(text)).IsEqualTo(expected);
    }

    [Test]
    public async Task TryWalk_reads_R1C1_text()
    {
        var accepted = FormulaText.TryWalk("SUM(R[-2]C[-2]:R[-2]C)", new List<string>(), ColumnProbe.Instance,
            FormulaNotation.R1C1, out _, out _);

        await Assert.That(accepted).IsTrue();
    }

    [Test]
    [Arguments("R[-2]C[-2]", "A1")]
    [Arguments("SUM(Table1[Start: Date])", "SUM(Table1[Start: Date])")]
    public async Task TryConvert_to_A1_reads_R1C1_text(string r1c1, string a1)
    {
        var converted = FormulaText.TryConvert(r1c1, new Point(3, 3), FormulaNotation.A1, out var text, out _);

        await Assert.That(converted).IsTrue();
        await Assert.That(text).IsEqualTo(a1);
    }

    /// <summary>
    /// The public edges turn a refusal into <see cref="ExpressionParseException"/>. The parser's own
    /// exception stays inside it, so the position the parser reported still reaches the caller.
    /// </summary>
    [Test]
    public async Task A_refusal_becomes_ExpressionParseException_with_the_parsers_exception_inside()
    {
        FormulaText.TryConvert(RefusedExternalReference, new Point(1, 1), FormulaNotation.R1C1, out _,
            out var refusal);

        var exception = refusal.ToException();

        await Assert.That(exception.InnerException).IsSameReferenceAs(refusal.Cause);
        await Assert.That(exception.Message).IsEqualTo(refusal.Message);
        await Assert.That(refusal.Message).Contains("char 0");
    }

    /// <summary>
    /// #543. The parser reads a function with any number of arguments. The calc engine's factory
    /// refuses one with too many or too few, by throwing <see cref="ExpressionParseException"/> while
    /// the parser reads the text. That is a refusal, as the parser's own is, so the parse returns it,
    /// and a public edge throws the same exception that evaluation threw before.
    /// </summary>
    [Test]
    [Arguments("ABS(1,2)", "Too many parameters for function 'ABS'.Expected a minimum of 1 and a maximum of 1.")]
    [Arguments("ABS()", "Too few parameters for function 'ABS'. Expected a minimum of 1 and a maximum of 1.")]
    public async Task Issue543_a_function_with_the_wrong_number_of_arguments_is_a_refusal(string text, string message)
    {
        using var wb = new XLWorkbook();
        var parser = new FormulaParser(wb.CalcEngine.Functions);

        var accepted = parser.TryGetAst(text, isA1: true, out _, out var refusal);

        await Assert.That(accepted).IsFalse();
        await Assert.That(refusal.Message).IsEqualTo(message);
        var thrown = await Assert.That(() => parser.GetAst(text, isA1: true)).Throws<ExpressionParseException>();
        await Assert.That(thrown!.Message).IsEqualTo(message);
        await Assert.That(thrown.InnerException).IsNull();
    }

    /// <summary>
    /// #543. Any factory that refuses a part of the text with <see cref="ExpressionParseException"/>
    /// gives a refusal, and the refusal throws that exception itself.
    /// </summary>
    [Test]
    public async Task A_factory_that_throws_ExpressionParseException_gives_a_refusal()
    {
        var refused = new ExpressionParseException("refused by the factory");

        var accepted = FormulaText.TryWalk("SUM(A1)", new List<string>(), new ThrowingProbe(refused),
            FormulaNotation.A1, out _, out var refusal);

        await Assert.That(accepted).IsFalse();
        await Assert.That(refusal.Cause).IsSameReferenceAs(refused);
        await Assert.That(refusal.Message).IsEqualTo("refused by the factory");
        await Assert.That(refusal.ToException()).IsSameReferenceAs(refused);
    }

    /// <summary>
    /// #543. Only a refusal becomes a value. Any other exception from a factory is a defect, and it
    /// reaches the caller.
    /// </summary>
    [Test]
    public async Task Any_other_exception_from_a_factory_reaches_the_caller()
    {
        var defect = new InvalidOperationException("a defect in the factory");

        await Assert.That(() => FormulaText.TryWalk("SUM(A1)", new List<string>(), new ThrowingProbe(defect),
            FormulaNotation.A1, out _, out _)).Throws<InvalidOperationException>();
    }

    /// <summary>Throws the exception it is given for the first function in the text.</summary>
    private sealed class ThrowingProbe(Exception exception) : CollectVisitor<List<string>>
    {
        public override object? Function(List<string> context, SymbolRange range, ReadOnlySpan<char> functionName,
            IReadOnlyList<object?> arguments)
            => throw exception;
    }

    /// <summary>Records every column name a structured reference names.</summary>
    private sealed class ColumnProbe : CollectVisitor<List<string>>
    {
        internal static readonly ColumnProbe Instance = new();

        public override object? StructureReference(List<string> context, SymbolRange range, string table,
            StructuredReferenceArea area, string? firstColumn, string? lastColumn)
        {
            if (firstColumn is not null)
                context.Add(firstColumn);

            if (lastColumn is not null && lastColumn != firstColumn)
                context.Add(lastColumn);

            return null;
        }
    }
}
