using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// Functions that read the values of a reference all go through the one sparse walk,
/// <c>CalcContext.UsedPointsWalk</c>. It visits only the cells that hold something, yet it must
/// still hand them over in row-major order — left to right, then top to bottom — because NPV, IRR
/// and MIRR depend on the order.
/// </summary>
public class ReferenceValueOrderTests
{
    /// <summary>
    /// A 2-column block whose row-major and column-major orders differ, with a blank cell to skip
    /// and one value at the very bottom of the sheet, reached only by a whole-column reference.
    /// </summary>
    private static IXLWorksheet NewSheet(out XLWorkbook wb, string left, string right, params XLCellValue[] values)
    {
        wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell(left + "1").Value = values[0];
        ws.Cell(right + "1").Value = values[1];
        ws.Cell(left + "2").Value = values[2];
        ws.Cell(right + "2").Value = values[3];
        // Row 3 of the left column stays blank.
        ws.Cell(right + "3").Value = values[4];
        ws.Cell(left + "1048576").Value = values[5];
        return ws;
    }

    [Test]
    public async Task Npv_ReadsAReferenceInRowMajorOrder()
    {
        var ws = NewSheet(out var wb, "A", "B", 1, 2, 3, 4, 5, 6);
        using (wb)
        {
            // 1/2 + 2/4 + 3/8 + 4/16 + 5/32. Column-major order (1, 3, 2, 4, 5) would give 1.90625.
            await Assert.That((double)ws.Evaluate("NPV(1, A1:B3)")).IsEqualTo(1.78125);
            await Assert.That((double)ws.Evaluate("NPV(1, A1:B3)")).IsEqualTo((double)ws.Evaluate("NPV(1, 1, 2, 3, 4, 5)"));

            // The whole-column reference adds the bottom cell last.
            await Assert.That((double)ws.Evaluate("NPV(1, A:B)")).IsEqualTo((double)ws.Evaluate("NPV(1, 1, 2, 3, 4, 5, 6)"));
        }
    }

    [Test]
    public async Task IrrAndMirr_ReadAReferenceInRowMajorOrder()
    {
        var ws = NewSheet(out var wb, "D", "E", -100, 10, 20, 110, 30, 5);
        using (wb)
        {
            await Assert.That((double)ws.Evaluate("IRR(D1:E3)")).IsEqualTo((double)ws.Evaluate("IRR({-100,10,20,110,30})")).Within(1e-12);
            await Assert.That((double)ws.Evaluate("IRR(D:E)")).IsEqualTo((double)ws.Evaluate("IRR({-100,10,20,110,30,5})")).Within(1e-12);
            await Assert.That((double)ws.Evaluate("IRR(D:E)")).IsNotEqualTo((double)ws.Evaluate("IRR({-100,20,5,10,110,30})"));

            await Assert.That((double)ws.Evaluate("MIRR(D:E, 0.1, 0.12)")).IsEqualTo((double)ws.Evaluate("MIRR({-100,10,20,110,30,5}, 0.1, 0.12)")).Within(1e-12);
        }
    }

    /// <summary>
    /// A dynamic-array anchor inside the reference that has not been evaluated yet: the cells it
    /// spills into are empty until it runs. Reading the anchor evaluates it, and the cells it then
    /// spills into, later in row-major order, must still be read.
    /// </summary>
    [Test]
    [Arguments("SEQUENCE(3)", "NPV(1, A1:A3)", 1.375, false)] // 1/2 + 2/4 + 3/8
    [Arguments("SEQUENCE(3)", "NPV(1, A1:A3)", 1.375, true)]
    [Arguments("SEQUENCE(3)", "NPV(1, A:A)", 1.375, false)]
    [Arguments("SEQUENCE(2, 2)", "NPV(1, A1:B2)", 1.625, false)] // 1/2 + 2/4 + 3/8 + 4/16, row by row
    [Arguments("SEQUENCE(2, 2)", "NPV(1, A1:B2)", 1.625, true)]
    [Arguments("SEQUENCE(3)", "SUM(A1:A3)", 6d, false)]
    [Arguments("SEQUENCE(3)", "SUM(A:A)", 6d, false)]
    [Arguments("SEQUENCE(3)", "SUM(A:A)", 6d, true)]
    [Arguments("SEQUENCE(3)", "COUNT(A1:A3)", 3d, false)]
    [Arguments("SEQUENCE(3, 1, 0)", "OR(A1:A3)", true, false)] // {0;1;2}: only the spilled cells are TRUE.
    [Arguments("SEQUENCE(3, 1, 0)", "OR(A1:A3)", true, true)]
    [Arguments("SEQUENCE(3, 1, 1, -1)", "AND(A1:A3)", false, false)] // {1;0;-1}: only a spilled cell is FALSE.
    [Arguments("SEQUENCE(3, 1, 1, -1)", "AND(A1:A3)", false, true)]
    public async Task UnevaluatedSpillInsideTheReferenceIsRead(string spill, string formula, object expected, bool inCell)
    {
        var ws = NewSpillSheet(out var wb, spill);
        using (wb)
        {
            XLCellValue actual;
            if (inCell)
            {
                // A cell formula, recalculated through the calculation chain.
                ws.Cell("Z1").FormulaA1 = formula;
                actual = ws.Cell("Z1").Value;
            }
            else
            {
                // Evaluated outside any cell, which evaluates dirty formulas recursively.
                actual = ws.Evaluate(formula);
            }

            await Assert.That(actual).IsEqualTo(XLCellValue.FromObject(expected));
        }
    }

    /// <summary>
    /// The anchor is outside the reference: a spilled cell inside it is read, which evaluates the
    /// dirty anchor, whose new, larger spill writes cells further down the reference.
    /// </summary>
    [Test]
    [Arguments("NPV(1, A2:A5)", 2.25, false)] // 2/2 + 3/4 + 4/8
    [Arguments("NPV(1, A2:A5)", 2.25, true)]
    [Arguments("SUM(A2:A5)", 9d, false)]
    [Arguments("SUM(A2:A5)", 9d, true)]
    public async Task SpillGrownByAnAnchorOutsideTheReferenceIsRead(string formula, double expected, bool inCell)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("B1").Value = 2;
        ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(B1)");
        await Assert.That(ws.Cell("A1").Value).IsEqualTo(1); // Evaluates and spills into A1:A2.
        await Assert.That(ws.Cell("A2").Value).IsEqualTo(2);

        ws.Cell("B1").Value = 4; // A1 is dirty, keeps its A1:A2 footprint; A3:A4 are still empty.
        XLCellValue actual;
        if (inCell)
        {
            ws.Cell("Z1").FormulaA1 = formula;
            actual = ws.Cell("Z1").Value;
        }
        else
        {
            actual = ws.Evaluate(formula);
        }

        await Assert.That((double)actual).IsEqualTo(expected);
    }

    /// <summary>
    /// SUBTOTAL, AGGREGATE and the criteria functions read a reference through their own sparse
    /// readers, which must also see the cells a dirty anchor spills into once it is read.
    /// </summary>
    [Test]
    [Arguments("SUBTOTAL(9,A1:A3)", 6d, false)]
    [Arguments("SUBTOTAL(9,A1:A3)", 6d, true)]
    [Arguments("SUBTOTAL(109,A1:A3)", 6d, false)]
    [Arguments("SUBTOTAL(3,A1:A3)", 3d, false)] // COUNTA goes through the same reader.
    [Arguments("AGGREGATE(9,0,A1:A3)", 6d, false)]
    [Arguments("AGGREGATE(9,5,A1:A3)", 6d, false)]
    [Arguments("SUMIF(A1:A3,\">0\")", 6d, false)]
    [Arguments("SUMIF(A1:A3,\">0\")", 6d, true)]
    [Arguments("COUNTIF(A1:A3,\">1\")", 2d, false)]
    [Arguments("SUMIFS(A1:A3,A1:A3,\">0\")", 6d, false)]
    [Arguments("AVERAGEIF(A:A,\">0\")", 2d, false)]
    public async Task UnevaluatedSpillIsReadByFilteredAndCriteriaReaders(string formula, double expected, bool inCell)
    {
        var ws = NewSpillSheet(out var wb, "SEQUENCE(3)");
        using (wb)
            await Assert.That((double)EvaluateFormula(ws, formula, inCell)).IsEqualTo(expected);
    }

    /// <summary>
    /// The hidden-row filter still applies to the cells found after a restart: row 2 is hidden, so
    /// only A1 and A3 of the spilled <c>{1;2;3}</c> count.
    /// </summary>
    [Test]
    [Arguments("SUBTOTAL(109,A1:A3)", 4d, false)]
    [Arguments("SUBTOTAL(109,A1:A3)", 4d, true)]
    [Arguments("SUBTOTAL(9,A1:A3)", 6d, false)] // 9 counts hidden rows.
    [Arguments("AGGREGATE(9,5,A1:A3)", 4d, false)]
    [Arguments("AGGREGATE(9,5,A1:A3)", 4d, true)]
    public async Task HiddenRowFilterAppliesAfterASpillRestart(string formula, double expected, bool inCell)
    {
        var ws = NewSpillSheet(out var wb, "SEQUENCE(3)");
        using (wb)
        {
            ws.Row(2).Hide();
            await Assert.That((double)EvaluateFormula(ws, formula, inCell)).IsEqualTo(expected);
        }
    }

    /// <summary>
    /// The anchor's own row is hidden, so the anchor is left out, but the cells it spills into are
    /// visible and count: 2 + 3. The anchor must still be evaluated so that it spills.
    /// </summary>
    [Test]
    [Arguments("SUBTOTAL(109,A1:A3)", 5d, false)]
    [Arguments("SUBTOTAL(109,A1:A3)", 5d, true)]
    [Arguments("AGGREGATE(9,5,A1:A3)", 5d, false)]
    [Arguments("AGGREGATE(9,5,A1:A3)", 5d, true)]
    [Arguments("SUBTOTAL(9,A1:A3)", 6d, false)]
    public async Task SpillOfAnAnchorOnAHiddenRowIsRead(string formula, double expected, bool inCell)
    {
        var ws = NewSpillSheet(out var wb, "SEQUENCE(3)");
        using (wb)
        {
            ws.Row(1).Hide();
            await Assert.That((double)EvaluateFormula(ws, formula, inCell)).IsEqualTo(expected);
        }
    }

    /// <summary>
    /// The nested-SUBTOTAL filter still applies to the cells found after a restart: A4 holds a
    /// SUBTOTAL of its own and is below the spill, so the walk reaches it only after the restart.
    /// </summary>
    [Test]
    [Arguments("SUBTOTAL(9,A1:A4)", 6d, false)]
    [Arguments("SUBTOTAL(9,A1:A4)", 6d, true)]
    [Arguments("AGGREGATE(9,0,A1:A4)", 6d, false)]
    [Arguments("AGGREGATE(9,0,A1:A4)", 6d, true)]
    [Arguments("SUM(A1:A4)", 106d, false)] // The unfiltered reader does count A4.
    public async Task NestedSubtotalFilterAppliesAfterASpillRestart(string formula, double expected, bool inCell)
    {
        var ws = NewSpillSheet(out var wb, "SEQUENCE(3)");
        using (wb)
        {
            ws.Cell("B1").Value = 100;
            ws.Cell("A4").FormulaA1 = "SUBTOTAL(9,B1)";
            await Assert.That((double)EvaluateFormula(ws, formula, inCell)).IsEqualTo(expected);
        }
    }

    private static XLCellValue EvaluateFormula(IXLWorksheet ws, string formula, bool inCell)
    {
        if (!inCell)
            return ws.Evaluate(formula); // Evaluates dirty formulas recursively.

        // A cell formula, recalculated through the calculation chain.
        ws.Cell("Z1").FormulaA1 = formula;
        return ws.Cell("Z1").Value;
    }

    private static IXLWorksheet NewSpillSheet(out XLWorkbook wb, string spill)
    {
        wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").SetDynamicFormulaA1(spill);
        return ws;
    }

    [Test]
    public async Task WholeColumnReferenceToASparseSheet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("J1").Value = true;
        ws.Cell("J1048576").Value = false;
        ws.Cell("K1").Value = 2;
        ws.Cell("K1048576").Value = 3;

        // Only the two used cells of each column are read; the million blanks in between are not.
        await Assert.That((bool)ws.Evaluate("AND(J:J)")).IsFalse();
        await Assert.That((bool)ws.Evaluate("OR(J:J)")).IsTrue();
        await Assert.That(ws.Evaluate("MULTINOMIAL(K:K)")).IsEqualTo(10); // 5! / (2! 3!)
        await Assert.That(ws.Evaluate("AND(L:L)")).IsEqualTo(XLError.IncompatibleValue); // No values at all.
    }
}
