using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;
// SEQUENCE / UNIQUE / SORT / SORTBY / FILTER / XLOOKUP / XMATCH.
// Array results are exercised through legacy CSE array formulas (FormulaArrayA1) over a correctly
// sized range; scalar results (and the top-left collapse) through ws.Evaluate.
public class DynamicArrayFunctionTests
{
    private static XLWorksheet NewSheet(out XLWorkbook wb)
    {
        wb = new XLWorkbook();
        return (XLWorksheet)wb.AddWorksheet("Sheet1");
    }

    [Test]
    public async Task Sequence_FillsRowMajor()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Range("A1:B3").FormulaArrayA1 = "SEQUENCE(3, 2)";
            await Assert.That(ws.Cell("A1").Value).IsEqualTo(1);
            await Assert.That(ws.Cell("B1").Value).IsEqualTo(2);
            await Assert.That(ws.Cell("A2").Value).IsEqualTo(3);
            await Assert.That(ws.Cell("B2").Value).IsEqualTo(4);
            await Assert.That(ws.Cell("A3").Value).IsEqualTo(5);
            await Assert.That(ws.Cell("B3").Value).IsEqualTo(6);

            // Start and step.
            ws.Range("D1:D3").FormulaArrayA1 = "SEQUENCE(3, 1, 10, 5)";
            await Assert.That(ws.Cell("D1").Value).IsEqualTo(10);
            await Assert.That(ws.Cell("D2").Value).IsEqualTo(15);
            await Assert.That(ws.Cell("D3").Value).IsEqualTo(20);
        }
    }

    [Test]
    public async Task Sequence_ScalarContext_ReturnsTopLeft()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            await Assert.That(ws.Evaluate("SEQUENCE(3, 2)")).IsEqualTo(1);
        }
    }

    [Test]
    public async Task Unique_ReturnsDistinctValuesInOrder()
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
            await Assert.That(ws.Cell("C1").Value).IsEqualTo(1);
            await Assert.That(ws.Cell("C2").Value).IsEqualTo(2);
            await Assert.That(ws.Cell("C3").Value).IsEqualTo(3);
        }
    }

    [Test]
    public async Task Unique_ExactlyOnce_KeepsValuesAppearingOnce()
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
            await Assert.That(ws.Cell("C1").Value).IsEqualTo(1);
            await Assert.That(ws.Cell("C2").Value).IsEqualTo(3);
        }
    }

    [Test]
    public async Task Sort_OrdersRows()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 3;
            ws.Cell("A2").Value = 1;
            ws.Cell("A3").Value = 4;
            ws.Cell("A4").Value = 1;

            ws.Range("C1:C4").FormulaArrayA1 = "SORT(A1:A4)";
            await Assert.That(ws.Cell("C1").Value).IsEqualTo(1);
            await Assert.That(ws.Cell("C2").Value).IsEqualTo(1);
            await Assert.That(ws.Cell("C3").Value).IsEqualTo(3);
            await Assert.That(ws.Cell("C4").Value).IsEqualTo(4);

            // Descending.
            ws.Range("D1:D4").FormulaArrayA1 = "SORT(A1:A4, 1, -1)";
            await Assert.That(ws.Cell("D1").Value).IsEqualTo(4);
            await Assert.That(ws.Cell("D2").Value).IsEqualTo(3);
            await Assert.That(ws.Cell("D3").Value).IsEqualTo(1);
            await Assert.That(ws.Cell("D4").Value).IsEqualTo(1);
        }
    }

    [Test]
    public async Task SortBy_OrdersBySeparateKey()
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
            await Assert.That(ws.Cell("C1").Value).IsEqualTo("a");
            await Assert.That(ws.Cell("C2").Value).IsEqualTo("b");
            await Assert.That(ws.Cell("C3").Value).IsEqualTo("c");
        }
    }

    [Test]
    public async Task Filter_KeepsMatchingRows()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 10;
            ws.Cell("A2").Value = 20;
            ws.Cell("A3").Value = 30;
            ws.Cell("A4").Value = 40;

            ws.Range("C1:C2").FormulaArrayA1 = "FILTER(A1:A4, A1:A4>25)";
            await Assert.That(ws.Cell("C1").Value).IsEqualTo(30);
            await Assert.That(ws.Cell("C2").Value).IsEqualTo(40);
        }
    }

    [Test]
    public async Task Filter_NoMatch_ReturnsIfEmpty()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 2;
            await Assert.That(ws.Evaluate("FILTER(A1:A2, A1:A2>9, \"none\")")).IsEqualTo("none");
        }
    }

    [Test]
    public async Task XLookup_ExactMatch()
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

            await Assert.That(ws.Evaluate("XLOOKUP(\"banana\", A1:A3, B1:B3)")).IsEqualTo(20);
            // Not found with a provided fallback.
            await Assert.That(ws.Evaluate("XLOOKUP(\"kiwi\", A1:A3, B1:B3, \"missing\")")).IsEqualTo("missing");
            // Not found without fallback -> #N/A.
            await Assert.That(ws.Evaluate("XLOOKUP(\"kiwi\", A1:A3, B1:B3)")).IsEqualTo(XLError.NoValueAvailable);
        }
    }

    [Test]
    public async Task XLookup_NextSmallerMatchMode()
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
            await Assert.That(ws.Evaluate("XLOOKUP(4, A1:A3, B1:B3, , -1)")).IsEqualTo("mid");
        }
    }

    [Test]
    public async Task XMatch_ReturnsPosition()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = "apple";
            ws.Cell("A2").Value = "banana";
            ws.Cell("A3").Value = "cherry";

            await Assert.That(ws.Evaluate("XMATCH(\"banana\", A1:A3)")).IsEqualTo(2);
            await Assert.That(ws.Evaluate("XMATCH(\"kiwi\", A1:A3)")).IsEqualTo(XLError.NoValueAvailable);

            ws.Cell("D1").Value = 1;
            ws.Cell("D2").Value = 3;
            ws.Cell("D3").Value = 5;
            // Next-smaller: 4 -> position of 3.
            await Assert.That(ws.Evaluate("XMATCH(4, D1:D3, -1)")).IsEqualTo(2);
        }
    }
}
