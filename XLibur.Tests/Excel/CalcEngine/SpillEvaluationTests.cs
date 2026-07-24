using NUnit.Framework;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

// Phase B1 — in-memory spilling of dynamic-array formulas.
// A dynamic-array formula in a single anchor cell auto-fills its computed footprint into the
// neighbouring cells, which stay formula-less. A blocked footprint (existing content) or one that
// runs past the sheet edge collapses to a #SPILL! error on the anchor. Re-evaluating to a smaller
// footprint clears the cells the previous result no longer covers.
[TestFixture]
public class SpillEvaluationTests
{
    private static XLWorksheet NewSheet(out XLWorkbook wb)
    {
        wb = new XLWorkbook();
        return (XLWorksheet)wb.AddWorksheet("Sheet1");
    }

    [Test]
    public void Spill_ColumnVector_FillsFootprint()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(3)");

            // Reading the anchor evaluates the formula and spills into A2:A3.
            Assert.AreEqual(1, ws.Cell("A1").Value);
            Assert.AreEqual(2, ws.Cell("A2").Value);
            Assert.AreEqual(3, ws.Cell("A3").Value);
        }
    }

    [Test]
    public void Spill_TwoDimensional_FillsGrid()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(2, 3)");

            Assert.AreEqual(1, ws.Cell("A1").Value);
            Assert.AreEqual(2, ws.Cell("B1").Value);
            Assert.AreEqual(3, ws.Cell("C1").Value);
            Assert.AreEqual(4, ws.Cell("A2").Value);
            Assert.AreEqual(5, ws.Cell("B2").Value);
            Assert.AreEqual(6, ws.Cell("C2").Value);
        }
    }

    [Test]
    public void Spill_AnchorHoldsFormula_SpilledCellsDoNot()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(3)");
            Assert.AreEqual(1, ws.Cell("A1").Value); // trigger the spill

            Assert.IsTrue(ws.Cell("A1").HasFormula, "Anchor must keep the formula");
            Assert.IsFalse(ws.Cell("A2").HasFormula, "Spilled cell must be formula-less");
            Assert.IsFalse(ws.Cell("A3").HasFormula, "Spilled cell must be formula-less");
        }
    }

    [Test]
    public void Spill_BlockedByExistingValue_ProducesSpillErrorAndPreservesBlocker()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A2").Value = "block";
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(3)");

            // The footprint A1:A3 collides with A2, so only the anchor is written.
            Assert.AreEqual(XLError.SpillRange, ws.Cell("A1").Value);
            Assert.AreEqual("block", ws.Cell("A2").Value, "Blocking value must be untouched");
            Assert.IsTrue(ws.Cell("A3").IsEmpty(), "No value is written to blocked-spill cells");
        }
    }

    [Test]
    public void Spill_BlockedByFormula_ProducesSpillError()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A3").FormulaA1 = "1+1";
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(3)");

            Assert.AreEqual(XLError.SpillRange, ws.Cell("A1").Value);
            Assert.IsTrue(ws.Cell("A2").IsEmpty(), "No value is written to blocked-spill cells");
        }
    }

    [Test]
    public void Spill_ShrinkingResult_ClearsStaleCells()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("B1").Value = 1;
            ws.Cell("B2").Value = 2;
            ws.Cell("B3").Value = 3;
            ws.Cell("A1").SetDynamicFormulaA1("UNIQUE(B1:B3)");

            Assert.AreEqual(1, ws.Cell("A1").Value); // spills A1:A3 = {1;2;3}
            Assert.AreEqual(3, ws.Cell("A3").Value);

            // Collapse a source value so only two distinct values remain: the same formula
            // instance now spills A1:A2 only and must clear the stale A3.
            ws.Cell("B3").Value = 2;
            Assert.AreEqual(1, ws.Cell("A1").Value);
            Assert.AreEqual(2, ws.Cell("A2").Value);
            Assert.IsTrue(ws.Cell("A3").IsEmpty(), "Stale cell of the previous footprint must be cleared");
        }
    }

    [Test]
    public void Spill_DependentOfSpilledCell_RecalculatesWhenSourceChanges()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("D1").Value = 1;
            ws.Cell("D2").Value = 2;
            ws.Cell("D3").Value = 3;
            ws.Cell("A1").SetDynamicFormulaA1("UNIQUE(D1:D3)"); // spills A1:A3 = {1;2;3}
            ws.Cell("C1").FormulaA1 = "A3*10";                  // depends on the spilled A3

            wb.CalcEngine.Recalculate(wb, null);
            Assert.AreEqual(30, ws.Cell("C1").Value);

            // Change a source cell so the spilled A3 becomes 5.
            ws.Cell("D3").Value = 5;
            Assert.IsTrue(ws.Cell("C1").NeedsRecalculation,
                "A dependent of a spilled (non-anchor) cell must be invalidated when the array's source changes");

            wb.CalcEngine.Recalculate(wb, null);
            Assert.AreEqual(50, ws.Cell("C1").Value);
        }
    }

    [Test]
    public void Spill_DependentBeforeAnchor_RecalculatesAfterInitialSpill()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("D1").Value = 1;
            ws.Cell("D2").Value = 2;
            ws.Cell("D3").Value = 3;
            // Anchor at C1 spills C1:C3; the dependent at A1 sits positionally BEFORE the anchor
            // and reads the spilled (non-anchor) C3 directly.
            ws.Cell("C1").SetDynamicFormulaA1("UNIQUE(D1:D3)");
            ws.Cell("A1").FormulaA1 = "C3*10";

            // Establish the spill so the spill-owner lookup knows C3 belongs to C1.
            wb.CalcEngine.Recalculate(wb, null);

            // A later source change must recompute the dependent with the fresh spilled value:
            // reading C3 now forces the dirty anchor C1 to evaluate first.
            ws.Cell("D3").Value = 5;
            wb.CalcEngine.Recalculate(wb, null);
            Assert.AreEqual(50, ws.Cell("A1").Value);
        }
    }

    [Test]
    [Ignore("Remaining limitation: on the VERY FIRST evaluation the spill footprint is unknown " +
            "until the anchor runs, so a dependent positioned before a not-yet-spilled anchor still " +
            "reads a blank cell. Full ordering here needs a calc-chain pre-pass that sizes arrays " +
            "before evaluation. Post-first-spill ordering is covered by " +
            nameof(Spill_DependentBeforeAnchor_RecalculatesAfterInitialSpill) + ".")]
    public void Spill_DependentBeforeAnchor_FirstEvaluationOrdering()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("D1").Value = 1;
            ws.Cell("D2").Value = 2;
            ws.Cell("D3").Value = 3;
            ws.Cell("C1").SetDynamicFormulaA1("UNIQUE(D1:D3)");
            ws.Cell("A1").FormulaA1 = "C3*10";

            wb.CalcEngine.Recalculate(wb, null);
            Assert.AreEqual(30, ws.Cell("A1").Value);
        }
    }

    [Test]
    public void SpillOperator_ReferencesWholeFootprint()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(3)"); // spills A1:A3 = {1;2;3}
            ws.Cell("C1").FormulaA1 = "SUM(A1#)";

            wb.CalcEngine.Recalculate(wb, null);
            Assert.AreEqual(6, ws.Cell("C1").Value);
        }
    }

    [Test]
    public void SpillOperator_NonAnchorCell_ReturnsRefError()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            // A1 holds no dynamic array, so A1# is #REF!.
            Assert.AreEqual(XLError.CellReference, ws.Evaluate("A1#"));
        }
    }

    [Test]
    public void SpillOperator_TracksFootprintWhenItShrinks()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("D1").Value = 1;
            ws.Cell("D2").Value = 2;
            ws.Cell("D3").Value = 3;
            ws.Cell("A1").SetDynamicFormulaA1("UNIQUE(D1:D3)"); // spills A1:A3
            ws.Cell("C1").FormulaA1 = "SUM(A1#)";

            wb.CalcEngine.Recalculate(wb, null);
            Assert.AreEqual(6, ws.Cell("C1").Value); // 1+2+3

            // Collapse to two distinct values: A1# now covers A1:A2 only.
            ws.Cell("D3").Value = 1;
            wb.CalcEngine.Recalculate(wb, null);
            Assert.AreEqual(3, ws.Cell("C1").Value); // 1+2
        }
    }

    [Test]
    public void SpillOperator_EvaluatesAnchorFirst_EvenWhenDependentComesBefore()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            // The dependent A1=SUM(C1#) sits positionally BEFORE its anchor C1. Because the
            // spill operator's range includes the anchor cell (which holds the dirty formula),
            // reading it forces the anchor to evaluate first — so this orders correctly, unlike
            // a plain read of a non-anchor spilled cell.
            ws.Cell("C1").SetDynamicFormulaA1("SEQUENCE(3)"); // spills C1:C3
            ws.Cell("A1").FormulaA1 = "SUM(C1#)";

            wb.CalcEngine.Recalculate(wb, null);
            Assert.AreEqual(6, ws.Cell("A1").Value);
        }
    }

    [Test]
    public void Spill_SurvivesRowInsertAndReSpills()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(3)"); // spills A1:A3
            wb.RecalculateAllFormulas();
            Assert.AreEqual(3, ws.Cell("A3").Value);

            // A structural insert relocates the anchor A1 -> A2. It stays a dynamic array and
            // re-spills over the shifted footprint A2:A4 after recalculation.
            ws.Row(1).InsertRowsAbove(1);
            wb.RecalculateAllFormulas();

            Assert.IsTrue(ws.Cell("A2").HasFormula, "Anchor must stay dynamic after the shift");
            Assert.AreEqual(1, ws.Cell("A2").Value);
            Assert.AreEqual(3, ws.Cell("A4").Value);
            Assert.IsFalse(ws.Cell("A3").HasFormula, "Spilled cell stays formula-less after the shift");
        }
    }

    [Test]
    public void Spill_PastSheetEdge_ProducesSpillError()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            // Anchor on the last row: a 2-row result would need a row beyond the sheet.
            var anchor = ws.Cell(XLHelper.MaxRowNumber, 1);
            anchor.SetDynamicFormulaA1("SEQUENCE(2)");

            Assert.AreEqual(XLError.SpillRange, anchor.Value);
        }
    }

    [Test]
    public void Spill_HorizontalVector_FillsRow()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(1, 3)"); // spills A1:C1

            Assert.AreEqual(1, ws.Cell("A1").Value);
            Assert.AreEqual(2, ws.Cell("B1").Value);
            Assert.AreEqual(3, ws.Cell("C1").Value);
        }
    }

    [Test]
    public void Spill_GrowingResult_FillsNewCells()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("D1").Value = 1;
            ws.Cell("D2").Value = 2;
            ws.Cell("D3").Value = 2;
            ws.Cell("A1").SetDynamicFormulaA1("UNIQUE(D1:D3)"); // {1;2} -> A1:A2

            Assert.AreEqual(1, ws.Cell("A1").Value); // trigger the spill
            Assert.AreEqual(2, ws.Cell("A2").Value);
            Assert.IsTrue(ws.Cell("A3").IsEmpty(), "Only two distinct values initially");

            // A third distinct value grows the footprint into the previously-empty A3.
            ws.Cell("D3").Value = 3;
            Assert.AreEqual(1, ws.Cell("A1").Value);
            Assert.AreEqual(3, ws.Cell("A3").Value);
        }
    }

    [Test]
    public void Spill_ErrorIsReportedByErrorFunctions()
    {
        // A real #SPILL! cell reports through ERROR.TYPE (9) and ISERROR — exercising the
        // XLError.SpillRange enum member end to end (the literal can't be parsed).
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A2").Value = "block";
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(3)");
            Assert.AreEqual(XLError.SpillRange, ws.Cell("A1").Value);

            ws.Cell("C1").FormulaA1 = "ERROR.TYPE(A1)";
            ws.Cell("C2").FormulaA1 = "ISERROR(A1)";
            Assert.AreEqual(9, ws.Cell("C1").Value);
            Assert.AreEqual(true, ws.Cell("C2").Value);
        }
    }

    [Test]
    public void Spill_RecoversAfterBlockerClearedAndAnchorReevaluates()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("D1").Value = 1;
            ws.Cell("D2").Value = 2;
            ws.Cell("D3").Value = 3;
            ws.Cell("A2").Value = "block";
            ws.Cell("A1").SetDynamicFormulaA1("UNIQUE(D1:D3)");
            Assert.AreEqual(XLError.SpillRange, ws.Cell("A1").Value);

            // Clear the blocker and change a source so the anchor re-evaluates: the spill recovers.
            ws.Cell("A2").Value = Blank.Value;
            ws.Cell("D3").Value = 4;
            Assert.AreEqual(1, ws.Cell("A1").Value);
            Assert.AreEqual(2, ws.Cell("A2").Value);
            Assert.AreEqual(4, ws.Cell("A3").Value);
        }
    }

    [Test]
    public void Spill_SurvivesColumnInsertAndReSpills()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(1, 3)"); // spills A1:C1
            wb.RecalculateAllFormulas();
            Assert.AreEqual(3, ws.Cell("C1").Value);

            // A column insert relocates the anchor A1 -> B1; it re-spills over B1:D1.
            ws.Column(1).InsertColumnsBefore(1);
            wb.RecalculateAllFormulas();

            Assert.IsTrue(ws.Cell("B1").HasFormula, "Anchor must stay dynamic after the shift");
            Assert.AreEqual(1, ws.Cell("B1").Value);
            Assert.AreEqual(3, ws.Cell("D1").Value);
        }
    }

    [Test]
    public void Spill_DependentBeforeAnchor_OrdersCorrectlyOnInteractiveRead()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("D1").Value = 1;
            ws.Cell("D2").Value = 2;
            ws.Cell("D3").Value = 3;
            ws.Cell("C1").SetDynamicFormulaA1("UNIQUE(D1:D3)");
            ws.Cell("A1").FormulaA1 = "C3*10";

            // Establish the spill by reading the anchor, then the dependent.
            Assert.AreEqual(1, ws.Cell("C1").Value);
            Assert.AreEqual(30, ws.Cell("A1").Value);

            // A plain .Value read of the dependent after a source change must order the dirty
            // anchor first (via the fallback to a full, dependency-ordered recalculation).
            ws.Cell("D3").Value = 5;
            Assert.AreEqual(50, ws.Cell("A1").Value);
        }
    }
}
