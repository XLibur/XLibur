using NUnit.Framework;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

// SEQUENCE / UNIQUE / SORT / SORTBY / FILTER / XLOOKUP / XMATCH.
// Array results are exercised through legacy CSE array formulas (FormulaArrayA1) over a correctly
// sized range; scalar results (and the top-left collapse) through ws.Evaluate.
[TestFixture]
public class DynamicArrayFunctionTests
{
    private static XLWorksheet NewSheet(out XLWorkbook wb)
    {
        wb = new XLWorkbook();
        return (XLWorksheet)wb.AddWorksheet("Sheet1");
    }

    [Test]
    public void Sequence_FillsRowMajor()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Range("A1:B3").FormulaArrayA1 = "SEQUENCE(3, 2)";
            Assert.AreEqual(1, ws.Cell("A1").Value);
            Assert.AreEqual(2, ws.Cell("B1").Value);
            Assert.AreEqual(3, ws.Cell("A2").Value);
            Assert.AreEqual(4, ws.Cell("B2").Value);
            Assert.AreEqual(5, ws.Cell("A3").Value);
            Assert.AreEqual(6, ws.Cell("B3").Value);

            // Start and step.
            ws.Range("D1:D3").FormulaArrayA1 = "SEQUENCE(3, 1, 10, 5)";
            Assert.AreEqual(10, ws.Cell("D1").Value);
            Assert.AreEqual(15, ws.Cell("D2").Value);
            Assert.AreEqual(20, ws.Cell("D3").Value);
        }
    }

    [Test]
    public void Sequence_ScalarContext_ReturnsTopLeft()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual(1, ws.Evaluate("SEQUENCE(3, 2)"));
        }
    }

    [Test]
    public void Unique_ReturnsDistinctValuesInOrder()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 2;
            ws.Cell("A3").Value = 2;
            ws.Cell("A4").Value = 3;
            ws.Cell("A5").Value = 1;

            ws.Range("C1:C3").FormulaArrayA1 = "UNIQUE(A1:A5)";
            Assert.AreEqual(1, ws.Cell("C1").Value);
            Assert.AreEqual(2, ws.Cell("C2").Value);
            Assert.AreEqual(3, ws.Cell("C3").Value);
        }
    }

    [Test]
    public void Unique_ExactlyOnce_KeepsValuesAppearingOnce()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 2;
            ws.Cell("A3").Value = 2;
            ws.Cell("A4").Value = 3;

            // by_col FALSE, exactly_once TRUE -> only 1 and 3.
            ws.Range("C1:C2").FormulaArrayA1 = "UNIQUE(A1:A4, FALSE, TRUE)";
            Assert.AreEqual(1, ws.Cell("C1").Value);
            Assert.AreEqual(3, ws.Cell("C2").Value);
        }
    }

    [Test]
    public void Sort_OrdersRows()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 3;
            ws.Cell("A2").Value = 1;
            ws.Cell("A3").Value = 4;
            ws.Cell("A4").Value = 1;

            ws.Range("C1:C4").FormulaArrayA1 = "SORT(A1:A4)";
            Assert.AreEqual(1, ws.Cell("C1").Value);
            Assert.AreEqual(1, ws.Cell("C2").Value);
            Assert.AreEqual(3, ws.Cell("C3").Value);
            Assert.AreEqual(4, ws.Cell("C4").Value);

            // Descending.
            ws.Range("D1:D4").FormulaArrayA1 = "SORT(A1:A4, 1, -1)";
            Assert.AreEqual(4, ws.Cell("D1").Value);
            Assert.AreEqual(3, ws.Cell("D2").Value);
            Assert.AreEqual(1, ws.Cell("D3").Value);
            Assert.AreEqual(1, ws.Cell("D4").Value);
        }
    }

    [Test]
    public void SortBy_OrdersBySeparateKey()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = "c";
            ws.Cell("A2").Value = "a";
            ws.Cell("A3").Value = "b";
            ws.Cell("B1").Value = 3;
            ws.Cell("B2").Value = 1;
            ws.Cell("B3").Value = 2;

            ws.Range("C1:C3").FormulaArrayA1 = "SORTBY(A1:A3, B1:B3)";
            Assert.AreEqual("a", ws.Cell("C1").Value);
            Assert.AreEqual("b", ws.Cell("C2").Value);
            Assert.AreEqual("c", ws.Cell("C3").Value);
        }
    }

    [Test]
    public void Filter_KeepsMatchingRows()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 10;
            ws.Cell("A2").Value = 20;
            ws.Cell("A3").Value = 30;
            ws.Cell("A4").Value = 40;

            ws.Range("C1:C2").FormulaArrayA1 = "FILTER(A1:A4, A1:A4>25)";
            Assert.AreEqual(30, ws.Cell("C1").Value);
            Assert.AreEqual(40, ws.Cell("C2").Value);
        }
    }

    [Test]
    public void Filter_NoMatch_ReturnsIfEmpty()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 2;
            Assert.AreEqual("none", ws.Evaluate("FILTER(A1:A2, A1:A2>9, \"none\")"));
        }
    }

    [Test]
    public void XLookup_ExactMatch()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = "apple";
            ws.Cell("A2").Value = "banana";
            ws.Cell("A3").Value = "cherry";
            ws.Cell("B1").Value = 10;
            ws.Cell("B2").Value = 20;
            ws.Cell("B3").Value = 30;

            Assert.AreEqual(20, ws.Evaluate("XLOOKUP(\"banana\", A1:A3, B1:B3)"));
            // Not found with a provided fallback.
            Assert.AreEqual("missing", ws.Evaluate("XLOOKUP(\"kiwi\", A1:A3, B1:B3, \"missing\")"));
            // Not found without fallback -> #N/A.
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("XLOOKUP(\"kiwi\", A1:A3, B1:B3)"));
        }
    }

    [Test]
    public void XLookup_NextSmallerMatchMode()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 3;
            ws.Cell("A3").Value = 5;
            ws.Cell("B1").Value = "low";
            ws.Cell("B2").Value = "mid";
            ws.Cell("B3").Value = "high";

            // 4 has no exact match; match mode -1 falls back to the next smaller (3 -> "mid").
            Assert.AreEqual("mid", ws.Evaluate("XLOOKUP(4, A1:A3, B1:B3, , -1)"));
        }
    }

    [Test]
    public void XMatch_ReturnsPosition()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = "apple";
            ws.Cell("A2").Value = "banana";
            ws.Cell("A3").Value = "cherry";

            Assert.AreEqual(2, ws.Evaluate("XMATCH(\"banana\", A1:A3)"));
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("XMATCH(\"kiwi\", A1:A3)"));

            ws.Cell("D1").Value = 1;
            ws.Cell("D2").Value = 3;
            ws.Cell("D3").Value = 5;
            // Next-smaller: 4 -> position of 3.
            Assert.AreEqual(2, ws.Evaluate("XMATCH(4, D1:D3, -1)"));
        }
    }
}
