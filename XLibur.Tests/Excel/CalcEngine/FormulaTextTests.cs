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
    /// converter, which throws on a refused formula. Executed, it does: the format cannot go in a
    /// <see cref="HashSet{T}"/>.
    /// </summary>
    [Test]
    public async Task A_conditional_format_holding_a_refused_formula_can_be_put_in_a_HashSet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var format = new XLConditionalFormat((XLRange)ws.Range("A1:A5"));
        format.Values.Add(new XLFormula("=" + RefusedExternalReference));

        var set = new HashSet<IXLConditionalFormat>(XLConditionalFormat.NoRangeComparer) { format };

        await Assert.That(set).Count().IsEqualTo(1);
    }

    /// <summary>
    /// The consequence of the comparer above that a caller meets: a save consolidates conditional
    /// formats by default, and consolidation converts each formula to R1C1, so the save throws.
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
}
