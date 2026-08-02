using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// AGGREGATE, which applies one of nineteen aggregates with an options argument controlling what is
/// left out of the data set. Expected values are the same as the equivalent standalone function
/// over the same data, which is what AGGREGATE is documented to compute.
/// </summary>
[SetCulture("en-US")]
public class AggregateTests
{
    private static IXLWorksheet NewSheet(out XLWorkbook wb)
    {
        wb = new XLWorkbook();
        return wb.AddWorksheet("Sheet1");
    }

    /// <summary>Put 1, 2, 3, 4 and 5 into A1:A5.</summary>
    private static void SeedNumbers(IXLWorksheet ws)
    {
        for (var row = 1; row <= 5; row++)
            ws.Cell(row, 1).Value = row;
    }

    [Test]
    [Arguments(1, 3d)] // AVERAGE
    [Arguments(2, 5d)] // COUNT
    [Arguments(3, 5d)] // COUNTA
    [Arguments(4, 5d)] // MAX
    [Arguments(5, 1d)] // MIN
    [Arguments(6, 120d)] // PRODUCT
    [Arguments(7, 1.5811388300841898d)] // STDEV.S
    [Arguments(8, 1.4142135623730951d)] // STDEV.P
    [Arguments(9, 15d)] // SUM
    [Arguments(10, 2.5d)] // VAR.S
    [Arguments(11, 2d)] // VAR.P
    [Arguments(12, 3d)] // MEDIAN
    public async Task Aggregate_AppliesTheNumberedFunction(int functionNumber, double expected)
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedNumbers(ws);
            ws.Cell("C1").FormulaA1 = $"AGGREGATE({functionNumber}, 0, A1:A5)";

            await Assert.That((double)ws.Cell("C1").Value).IsEqualTo(expected).Within(1e-12);
        }
    }

    [Test]
    public async Task Aggregate_MatchesTheStandaloneFunctions()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedNumbers(ws);
            ws.Cell("C1").FormulaA1 = "AGGREGATE(1, 0, A1:A5) - AVERAGE(A1:A5)";
            ws.Cell("C2").FormulaA1 = "AGGREGATE(9, 0, A1:A5) - SUM(A1:A5)";
            ws.Cell("C3").FormulaA1 = "AGGREGATE(7, 0, A1:A5) - STDEV(A1:A5)";
            ws.Cell("C4").FormulaA1 = "AGGREGATE(12, 0, A1:A5) - MEDIAN(A1:A5)";

            foreach (var address in new[] { "C1", "C2", "C3", "C4" })
                await Assert.That((double)ws.Cell(address).Value).IsEqualTo(0d).Within(1e-12);
        }
    }

    [Test]
    public async Task Aggregate_Mode()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 2;
            ws.Cell("A3").Value = 2;
            ws.Cell("A4").Value = 3;
            ws.Cell("C1").FormulaA1 = "AGGREGATE(13, 0, A1:A4)";

            await Assert.That((double)ws.Cell("C1").Value).IsEqualTo(2d);
        }
    }

    [Test]
    [Arguments("AGGREGATE(14, 0, A1:A5, 2)", 4d)] // LARGE
    [Arguments("AGGREGATE(15, 0, A1:A5, 2)", 2d)] // SMALL
    [Arguments("AGGREGATE(16, 0, A1:A5, 0.5)", 3d)] // PERCENTILE.INC
    [Arguments("AGGREGATE(17, 0, A1:A5, 1)", 2d)] // QUARTILE.INC
    // With five values the exclusive rank of the median is 0.5*(5+1) = 3, the third value.
    [Arguments("AGGREGATE(18, 0, A1:A5, 0.5)", 3d)] // PERCENTILE.EXC
    [Arguments("AGGREGATE(19, 0, A1:A5, 2)", 3d)] // QUARTILE.EXC
    public async Task Aggregate_OrderStatisticsTakeAK(string formula, double expected)
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedNumbers(ws);
            ws.Cell("C1").FormulaA1 = formula;

            await Assert.That((double)ws.Cell("C1").Value).IsEqualTo(expected).Within(1e-12);
        }
    }

    [Test]
    public async Task Aggregate_ExclusivePercentileRejectsUnreachableRanks()
    {
        // With five values the exclusive percentile only spans 1/6 to 5/6.
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedNumbers(ws);
            ws.Cell("C1").FormulaA1 = "AGGREGATE(18, 0, A1:A5, 0.1)";
            ws.Cell("C2").FormulaA1 = "AGGREGATE(18, 0, A1:A5, 0.9)";
            ws.Cell("C3").FormulaA1 = "AGGREGATE(19, 0, A1:A5, 0)"; // Only quartiles 1..3 exist.
            ws.Cell("C4").FormulaA1 = "AGGREGATE(19, 0, A1:A5, 4)";

            foreach (var address in new[] { "C1", "C2", "C3", "C4" })
                await Assert.That(ws.Cell(address).Value).IsEqualTo(XLError.NumberInvalid);
        }
    }

    [Test]
    public async Task Aggregate_IgnoresErrorsOnlyWhenAskedTo()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").FormulaA1 = "1/0";
            ws.Cell("A3").Value = 3;

            ws.Cell("C1").FormulaA1 = "AGGREGATE(9, 0, A1:A3)"; // Errors propagate.
            ws.Cell("C2").FormulaA1 = "AGGREGATE(9, 2, A1:A3)"; // Option 2 ignores them.
            ws.Cell("C3").FormulaA1 = "AGGREGATE(9, 6, A1:A3)"; // As does option 6.
            ws.Cell("C4").FormulaA1 = "AGGREGATE(4, 2, A1:A3)"; // MAX over the surviving values.

            await Assert.That(ws.Cell("C1").Value).IsEqualTo(XLError.DivisionByZero);
            await Assert.That((double)ws.Cell("C2").Value).IsEqualTo(4d);
            await Assert.That((double)ws.Cell("C3").Value).IsEqualTo(4d);
            await Assert.That((double)ws.Cell("C4").Value).IsEqualTo(3d);
        }
    }

    [Test]
    public async Task Aggregate_IgnoresHiddenRowsOnlyWhenAskedTo()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedNumbers(ws);
            ws.Row(2).Hide();
            ws.Row(4).Hide();

            ws.Cell("C1").FormulaA1 = "AGGREGATE(9, 0, A1:A5)"; // Every row counts.
            ws.Cell("C2").FormulaA1 = "AGGREGATE(9, 1, A1:A5)"; // Option 1 skips hidden rows.
            ws.Cell("C3").FormulaA1 = "AGGREGATE(9, 5, A1:A5)"; // As does option 5.
            ws.Cell("C4").FormulaA1 = "AGGREGATE(2, 1, A1:A5)"; // COUNT over the visible rows.

            await Assert.That((double)ws.Cell("C1").Value).IsEqualTo(15d);
            await Assert.That((double)ws.Cell("C2").Value).IsEqualTo(9d); // 1 + 3 + 5.
            await Assert.That((double)ws.Cell("C3").Value).IsEqualTo(9d);
            await Assert.That((double)ws.Cell("C4").Value).IsEqualTo(3d);
        }
    }

    [Test]
    public async Task Aggregate_IgnoresHiddenRowsAndErrorsTogether()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").FormulaA1 = "1/0";
            ws.Cell("A3").Value = 3;
            ws.Cell("A4").Value = 4;
            ws.Row(4).Hide();

            ws.Cell("C1").FormulaA1 = "AGGREGATE(9, 3, A1:A4)";
            ws.Cell("C2").FormulaA1 = "AGGREGATE(9, 7, A1:A4)";

            await Assert.That((double)ws.Cell("C1").Value).IsEqualTo(4d); // 1 + 3.
            await Assert.That((double)ws.Cell("C2").Value).IsEqualTo(4d);
        }
    }

    [Test]
    public async Task Aggregate_SkipsNestedSubtotalsAndAggregates()
    {
        // A subtotal inside the range is not counted again by the aggregate that spans it.
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 2;
            ws.Cell("A3").FormulaA1 = "SUBTOTAL(9, A1:A2)";
            ws.Cell("A4").FormulaA1 = "AGGREGATE(9, 0, A1:A2)";
            ws.Cell("A5").Value = 4;

            ws.Cell("C1").FormulaA1 = "AGGREGATE(9, 0, A1:A5)";
            ws.Cell("C2").FormulaA1 = "SUM(A1:A5)";

            await Assert.That((double)ws.Cell("C1").Value).IsEqualTo(7d); // 1 + 2 + 4.
            await Assert.That((double)ws.Cell("C2").Value).IsEqualTo(13d); // SUM counts them all.
        }
    }

    [Test]
    public async Task Aggregate_TakesSeveralRanges()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedNumbers(ws);
            ws.Cell("B1").Value = 10;
            ws.Cell("B2").Value = 20;

            ws.Cell("D1").FormulaA1 = "AGGREGATE(9, 0, A1:A5, B1:B2)";

            await Assert.That((double)ws.Cell("D1").Value).IsEqualTo(45d);
        }
    }

    [Test]
    [Arguments("AGGREGATE(0, 0, A1:A5)")] // Function numbers run 1..19.
    [Arguments("AGGREGATE(20, 0, A1:A5)")]
    [Arguments("AGGREGATE(9, -1, A1:A5)")] // Options run 0..7.
    [Arguments("AGGREGATE(9, 8, A1:A5)")]
    [Arguments("AGGREGATE(14, 0, A1:A5)")] // The order statistics need a k.
    [Arguments("AGGREGATE(18, 0, A1:A5)")]
    public async Task Aggregate_OutOfRangeArgumentsReturnIncompatibleValue(string formula)
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedNumbers(ws);
            ws.Cell("C1").FormulaA1 = formula;

            await Assert.That(ws.Cell("C1").Value).IsEqualTo(XLError.IncompatibleValue);
        }
    }

    [Test]
    public async Task Aggregate_OptionDefaultsAndReadsArgumentsFromCells()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedNumbers(ws);
            ws.Cell("C1").Value = 9;
            ws.Cell("C2").Value = 0;
            ws.Cell("D1").FormulaA1 = "AGGREGATE(C1, C2, A1:A5)";

            await Assert.That((double)ws.Cell("D1").Value).IsEqualTo(15d);
        }
    }

    [Test]
    public async Task Subtotal_CoversBothTheVisibleAndTheHiddenRowVariants()
    {
        // SUBTOTAL 1..11 count hidden rows and 101..111 do not; this pins that the whole pair of
        // ranges is wired up, which AGGREGATE's options build on.
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedNumbers(ws);
            ws.Row(2).Hide();

            for (var i = 0; i < 11; i++)
            {
                ws.Cell(i + 1, 3).FormulaA1 = $"SUBTOTAL({i + 1}, A1:A5)";
                ws.Cell(i + 1, 4).FormulaA1 = $"SUBTOTAL({i + 101}, A1:A5)";
            }

            // Only the aggregates that depend on the hidden value differ between the two ranges.
            await Assert.That((double)ws.Cell("C9").Value).IsEqualTo(15d); // SUM over everything.
            await Assert.That((double)ws.Cell("D9").Value).IsEqualTo(13d); // SUM without row 2.
            await Assert.That((double)ws.Cell("C2").Value).IsEqualTo(5d); // COUNT.
            await Assert.That((double)ws.Cell("D2").Value).IsEqualTo(4d);
            await Assert.That((double)ws.Cell("C6").Value).IsEqualTo(120d); // PRODUCT.
            await Assert.That((double)ws.Cell("D6").Value).IsEqualTo(60d);
        }
    }
}
