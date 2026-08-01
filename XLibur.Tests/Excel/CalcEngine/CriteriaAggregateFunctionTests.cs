using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;
// AVERAGEIF / AVERAGEIFS / MAXIFS / MINIFS — the criteria aggregate functions built on the same
// TallyCriteria machinery as SUMIFS / COUNTIFS.
public class CriteriaAggregateFunctionTests
{
    private const double Tolerance = 1e-9;

    // Region | Category | Sales
    // North  | A        | 100
    // South  | B        | 200
    // North  | A        | 300
    // West   | B        | 400
    // North  | B        | 500
    private static XLWorkbook CreateSampleWorkbook()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        string[] regions = { "North", "South", "North", "West", "North" };
        string[] categories = { "A", "B", "A", "B", "B" };
        double[] sales = { 100, 200, 300, 400, 500 };
        for (var i = 0; i < regions.Length; i++)
        {
            ws.Cell(i + 1, 1).Value = regions[i];
            ws.Cell(i + 1, 2).Value = categories[i];
            ws.Cell(i + 1, 3).Value = sales[i];
        }

        return wb;
    }

    [Test]
    public async Task AverageIf_WithSeparateAverageRange()
    {
        using var wb = CreateSampleWorkbook();
        var ws = wb.Worksheet("Data");

        // Average of Sales where Region = North -> (100 + 300 + 500) / 3
        await Assert.That((double)ws.Evaluate("AVERAGEIF(A1:A5, \"North\", C1:C5)")).IsEqualTo(300d).Within(Tolerance);
    }

    [Test]
    public async Task AverageIf_WithoutAverageRange_AveragesTheCriteriaRange()
    {
        using var wb = CreateSampleWorkbook();
        var ws = wb.Worksheet("Data");

        // Average of the values in C1:C5 that are > 150 -> (200 + 300 + 400 + 500) / 4
        await Assert.That((double)ws.Evaluate("AVERAGEIF(C1:C5, \">150\")")).IsEqualTo(350d).Within(Tolerance);
    }

    [Test]
    public async Task AverageIf_NoMatch_ReturnsDivisionByZero()
    {
        using var wb = CreateSampleWorkbook();
        var ws = wb.Worksheet("Data");

        await Assert.That(ws.Evaluate("AVERAGEIF(A1:A5, \"East\", C1:C5)")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    public async Task AverageIfs_MultipleCriteria()
    {
        using var wb = CreateSampleWorkbook();
        var ws = wb.Worksheet("Data");

        // Average of Sales where Region = North AND Category = A -> (100 + 300) / 2
        await Assert.That((double)ws.Evaluate("AVERAGEIFS(C1:C5, A1:A5, \"North\", B1:B5, \"A\")")).IsEqualTo(200d).Within(Tolerance);
    }

    [Test]
    public async Task AverageIfs_NoMatch_ReturnsDivisionByZero()
    {
        using var wb = CreateSampleWorkbook();
        var ws = wb.Worksheet("Data");

        await Assert.That(ws.Evaluate("AVERAGEIFS(C1:C5, A1:A5, \"East\")")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    public async Task MaxIfs_ReturnsMaxOfMatchingCells()
    {
        using var wb = CreateSampleWorkbook();
        var ws = wb.Worksheet("Data");

        await Assert.That((double)ws.Evaluate("MAXIFS(C1:C5, A1:A5, \"North\")")).IsEqualTo(500d).Within(Tolerance);
        // Region = North AND Category = A -> max(100, 300)
        await Assert.That((double)ws.Evaluate("MAXIFS(C1:C5, A1:A5, \"North\", B1:B5, \"A\")")).IsEqualTo(300d).Within(Tolerance);
    }

    [Test]
    public async Task MinIfs_ReturnsMinOfMatchingCells()
    {
        using var wb = CreateSampleWorkbook();
        var ws = wb.Worksheet("Data");

        await Assert.That((double)ws.Evaluate("MINIFS(C1:C5, A1:A5, \"North\")")).IsEqualTo(100d).Within(Tolerance);
        // Region = North AND Category = A -> min(100, 300)
        await Assert.That((double)ws.Evaluate("MINIFS(C1:C5, A1:A5, \"North\", B1:B5, \"A\")")).IsEqualTo(100d).Within(Tolerance);
    }

    [Test]
    public async Task MaxIfs_And_MinIfs_NoMatch_ReturnZero()
    {
        using var wb = CreateSampleWorkbook();
        var ws = wb.Worksheet("Data");

        // Excel returns 0 (not an error) when no cell satisfies the criteria.
        await Assert.That((double)ws.Evaluate("MAXIFS(C1:C5, A1:A5, \"East\")")).IsEqualTo(0d).Within(Tolerance);
        await Assert.That((double)ws.Evaluate("MINIFS(C1:C5, A1:A5, \"East\")")).IsEqualTo(0d).Within(Tolerance);
    }

    [Test]
    public async Task CriteriaAndValueRangeSizeMismatch_ReturnsValueError()
    {
        using var wb = CreateSampleWorkbook();
        var ws = wb.Worksheet("Data");

        // Value range (5 rows) and criteria range (4 rows) differ in size.
        await Assert.That(ws.Evaluate("AVERAGEIFS(C1:C5, A1:A4, \"North\")")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(ws.Evaluate("MAXIFS(C1:C5, A1:A4, \"North\")")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(ws.Evaluate("MINIFS(C1:C5, A1:A4, \"North\")")).IsEqualTo(XLError.IncompatibleValue);
    }
}
