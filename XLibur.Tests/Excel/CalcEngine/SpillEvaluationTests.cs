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
}
