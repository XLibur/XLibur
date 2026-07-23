using NUnit.Framework;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

// AVERAGEIF / AVERAGEIFS / MAXIFS / MINIFS — the criteria aggregate functions built on the same
// TallyCriteria machinery as SUMIFS / COUNTIFS.
[TestFixture]
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
    public void AverageIf_WithSeparateAverageRange()
    {
        using var wb = CreateSampleWorkbook();
        var ws = wb.Worksheet("Data");

        // Average of Sales where Region = North -> (100 + 300 + 500) / 3
        Assert.AreEqual(300d, (double)ws.Evaluate("AVERAGEIF(A1:A5, \"North\", C1:C5)"), Tolerance);
    }

    [Test]
    public void AverageIf_WithoutAverageRange_AveragesTheCriteriaRange()
    {
        using var wb = CreateSampleWorkbook();
        var ws = wb.Worksheet("Data");

        // Average of the values in C1:C5 that are > 150 -> (200 + 300 + 400 + 500) / 4
        Assert.AreEqual(350d, (double)ws.Evaluate("AVERAGEIF(C1:C5, \">150\")"), Tolerance);
    }

    [Test]
    public void AverageIf_NoMatch_ReturnsDivisionByZero()
    {
        using var wb = CreateSampleWorkbook();
        var ws = wb.Worksheet("Data");

        Assert.AreEqual(XLError.DivisionByZero, ws.Evaluate("AVERAGEIF(A1:A5, \"East\", C1:C5)"));
    }

    [Test]
    public void AverageIfs_MultipleCriteria()
    {
        using var wb = CreateSampleWorkbook();
        var ws = wb.Worksheet("Data");

        // Average of Sales where Region = North AND Category = A -> (100 + 300) / 2
        Assert.AreEqual(200d, (double)ws.Evaluate("AVERAGEIFS(C1:C5, A1:A5, \"North\", B1:B5, \"A\")"), Tolerance);
    }

    [Test]
    public void AverageIfs_NoMatch_ReturnsDivisionByZero()
    {
        using var wb = CreateSampleWorkbook();
        var ws = wb.Worksheet("Data");

        Assert.AreEqual(XLError.DivisionByZero, ws.Evaluate("AVERAGEIFS(C1:C5, A1:A5, \"East\")"));
    }

    [Test]
    public void MaxIfs_ReturnsMaxOfMatchingCells()
    {
        using var wb = CreateSampleWorkbook();
        var ws = wb.Worksheet("Data");

        Assert.AreEqual(500d, (double)ws.Evaluate("MAXIFS(C1:C5, A1:A5, \"North\")"), Tolerance);
        // Region = North AND Category = A -> max(100, 300)
        Assert.AreEqual(300d, (double)ws.Evaluate("MAXIFS(C1:C5, A1:A5, \"North\", B1:B5, \"A\")"), Tolerance);
    }

    [Test]
    public void MinIfs_ReturnsMinOfMatchingCells()
    {
        using var wb = CreateSampleWorkbook();
        var ws = wb.Worksheet("Data");

        Assert.AreEqual(100d, (double)ws.Evaluate("MINIFS(C1:C5, A1:A5, \"North\")"), Tolerance);
        // Region = North AND Category = A -> min(100, 300)
        Assert.AreEqual(100d, (double)ws.Evaluate("MINIFS(C1:C5, A1:A5, \"North\", B1:B5, \"A\")"), Tolerance);
    }

    [Test]
    public void MaxIfs_And_MinIfs_NoMatch_ReturnZero()
    {
        using var wb = CreateSampleWorkbook();
        var ws = wb.Worksheet("Data");

        // Excel returns 0 (not an error) when no cell satisfies the criteria.
        Assert.AreEqual(0d, (double)ws.Evaluate("MAXIFS(C1:C5, A1:A5, \"East\")"), Tolerance);
        Assert.AreEqual(0d, (double)ws.Evaluate("MINIFS(C1:C5, A1:A5, \"East\")"), Tolerance);
    }

    [Test]
    public void CriteriaAndValueRangeSizeMismatch_ReturnsValueError()
    {
        using var wb = CreateSampleWorkbook();
        var ws = wb.Worksheet("Data");

        // Value range (5 rows) and criteria range (4 rows) differ in size.
        Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("AVERAGEIFS(C1:C5, A1:A4, \"North\")"));
        Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("MAXIFS(C1:C5, A1:A4, \"North\")"));
        Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("MINIFS(C1:C5, A1:A4, \"North\")"));
    }
}
