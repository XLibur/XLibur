using NUnit.Framework;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

// SMALL / RANK / PERCENTILE / QUARTILE / MODE and their modern aliases.
[TestFixture]
public class StatisticalRankPercentileTests
{
    private const double Tolerance = 1e-9;

    // A1:A7 = 3, 1, 4, 1, 5, 9, 2  ->  sorted: 1, 1, 2, 3, 4, 5, 9
    private static XLWorksheet SampleSheet(out XLWorkbook wb)
    {
        wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("Data");
        double[] values = { 3, 1, 4, 1, 5, 9, 2 };
        for (var i = 0; i < values.Length; i++)
            ws.Cell(i + 1, 1).Value = values[i];
        return ws;
    }

    [Test]
    public void Small_ReturnsKthSmallest()
    {
        var ws = SampleSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual(1d, (double)ws.Evaluate("SMALL(A1:A7, 1)"), Tolerance);
            Assert.AreEqual(1d, (double)ws.Evaluate("SMALL(A1:A7, 2)"), Tolerance);
            Assert.AreEqual(2d, (double)ws.Evaluate("SMALL(A1:A7, 3)"), Tolerance);
            Assert.AreEqual(9d, (double)ws.Evaluate("SMALL(A1:A7, 7)"), Tolerance);
            // Mirror of LARGE.
            Assert.AreEqual(9d, (double)ws.Evaluate("LARGE(A1:A7, 1)"), Tolerance);
            Assert.AreEqual(3d, (double)ws.Evaluate("SMALL({5,3,8,1}, 2)"), Tolerance);
        }
    }

    [Test]
    public void Small_OutOfRangeK_ReturnsNumberInvalid()
    {
        var ws = SampleSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual(XLError.NumberInvalid, ws.Evaluate("SMALL(A1:A7, 8)"));
            Assert.AreEqual(XLError.NumberInvalid, ws.Evaluate("SMALL(A1:A7, 0)"));
        }
    }

    [Test]
    public void Rank_DescendingByDefault_AscendingWhenOrderNonZero()
    {
        var ws = SampleSheet(out var wb);
        using (wb)
        {
            // Descending (default): 9=1, 5=2, 4=3
            Assert.AreEqual(3, ws.Evaluate("RANK(4, A1:A7)"));
            // Tied values (two 1s) share the top rank of the group.
            Assert.AreEqual(6, ws.Evaluate("RANK(1, A1:A7)"));
            // Ascending: values below 4 are {1,1,2,3} -> rank 5
            Assert.AreEqual(5, ws.Evaluate("RANK(4, A1:A7, 1)"));
            // RANK.EQ is an alias.
            Assert.AreEqual(3, ws.Evaluate("RANK.EQ(4, A1:A7)"));
        }
    }

    [Test]
    public void Rank_NumberNotPresent_ReturnsNotAvailable()
    {
        var ws = SampleSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("RANK(7, A1:A7)"));
        }
    }

    [Test]
    public void Mode_ReturnsMostFrequentValue()
    {
        var ws = SampleSheet(out var wb);
        using (wb)
        {
            // Only 1 repeats.
            Assert.AreEqual(1d, (double)ws.Evaluate("MODE(A1:A7)"), Tolerance);
            // Ties resolve to the value whose first occurrence is earliest.
            Assert.AreEqual(4d, (double)ws.Evaluate("MODE(4, 4, 2, 2)"), Tolerance);
            Assert.AreEqual(1d, (double)ws.Evaluate("MODE.SNGL(A1:A7)"), Tolerance);
        }
    }

    [Test]
    public void Mode_NoRepeats_ReturnsNotAvailable()
    {
        var ws = SampleSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("MODE(1, 2, 3, 4)"));
        }
    }

    [Test]
    public void Percentile_InterpolatesBetweenRanks()
    {
        var ws = SampleSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual(1d, (double)ws.Evaluate("PERCENTILE(A1:A7, 0)"), Tolerance);
            Assert.AreEqual(9d, (double)ws.Evaluate("PERCENTILE(A1:A7, 1)"), Tolerance);
            Assert.AreEqual(3d, (double)ws.Evaluate("PERCENTILE(A1:A7, 0.5)"), Tolerance);
            Assert.AreEqual(1.5d, (double)ws.Evaluate("PERCENTILE(A1:A7, 0.25)"), Tolerance);
            Assert.AreEqual(1.5d, (double)ws.Evaluate("PERCENTILE.INC(A1:A7, 0.25)"), Tolerance);
        }
    }

    [Test]
    public void Percentile_OutOfRange_ReturnsNumberInvalid()
    {
        var ws = SampleSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual(XLError.NumberInvalid, ws.Evaluate("PERCENTILE(A1:A7, 1.1)"));
            Assert.AreEqual(XLError.NumberInvalid, ws.Evaluate("PERCENTILE(A1:A7, -0.1)"));
        }
    }

    [Test]
    public void Quartile_MapsToInclusivePercentiles()
    {
        var ws = SampleSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual(1d, (double)ws.Evaluate("QUARTILE(A1:A7, 0)"), Tolerance);
            Assert.AreEqual(1.5d, (double)ws.Evaluate("QUARTILE(A1:A7, 1)"), Tolerance);
            Assert.AreEqual(3d, (double)ws.Evaluate("QUARTILE(A1:A7, 2)"), Tolerance);
            Assert.AreEqual(4.5d, (double)ws.Evaluate("QUARTILE(A1:A7, 3)"), Tolerance);
            Assert.AreEqual(9d, (double)ws.Evaluate("QUARTILE(A1:A7, 4)"), Tolerance);
            Assert.AreEqual(4.5d, (double)ws.Evaluate("QUARTILE.INC(A1:A7, 3)"), Tolerance);
        }
    }

    [Test]
    public void Quartile_OutOfRange_ReturnsNumberInvalid()
    {
        var ws = SampleSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual(XLError.NumberInvalid, ws.Evaluate("QUARTILE(A1:A7, 5)"));
        }
    }
}
