using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;
// SMALL / RANK / PERCENTILE / QUARTILE / MODE and their modern aliases.
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
    public async Task Small_ReturnsKthSmallest()
    {
        var ws = SampleSheet(out var wb);
        using (wb)
        {
            await Assert.That((double)ws.Evaluate("SMALL(A1:A7, 1)")).IsEqualTo(1d).Within(Tolerance);
            await Assert.That((double)ws.Evaluate("SMALL(A1:A7, 2)")).IsEqualTo(1d).Within(Tolerance);
            await Assert.That((double)ws.Evaluate("SMALL(A1:A7, 3)")).IsEqualTo(2d).Within(Tolerance);
            await Assert.That((double)ws.Evaluate("SMALL(A1:A7, 7)")).IsEqualTo(9d).Within(Tolerance);
            // Mirror of LARGE.
            await Assert.That((double)ws.Evaluate("LARGE(A1:A7, 1)")).IsEqualTo(9d).Within(Tolerance);
            await Assert.That((double)ws.Evaluate("SMALL({5,3,8,1}, 2)")).IsEqualTo(3d).Within(Tolerance);
        }
    }

    [Test]
    public async Task Small_OutOfRangeK_ReturnsNumberInvalid()
    {
        var ws = SampleSheet(out var wb);
        using (wb)
        {
            await Assert.That(ws.Evaluate("SMALL(A1:A7, 8)")).IsEqualTo(XLError.NumberInvalid);
            await Assert.That(ws.Evaluate("SMALL(A1:A7, 0)")).IsEqualTo(XLError.NumberInvalid);
        }
    }

    [Test]
    public async Task Rank_DescendingByDefault_AscendingWhenOrderNonZero()
    {
        var ws = SampleSheet(out var wb);
        using (wb)
        {
            // Descending (default): 9=1, 5=2, 4=3
            await Assert.That(ws.Evaluate("RANK(4, A1:A7)")).IsEqualTo(3);
            // Tied values (two 1s) share the top rank of the group.
            await Assert.That(ws.Evaluate("RANK(1, A1:A7)")).IsEqualTo(6);
            // Ascending: values below 4 are {1,1,2,3} -> rank 5
            await Assert.That(ws.Evaluate("RANK(4, A1:A7, 1)")).IsEqualTo(5);
            // RANK.EQ is an alias.
            await Assert.That(ws.Evaluate("RANK.EQ(4, A1:A7)")).IsEqualTo(3);
        }
    }

    [Test]
    public async Task Rank_NumberNotPresent_ReturnsNotAvailable()
    {
        var ws = SampleSheet(out var wb);
        using (wb)
        {
            await Assert.That(ws.Evaluate("RANK(7, A1:A7)")).IsEqualTo(XLError.NoValueAvailable);
        }
    }

    [Test]
    public async Task Mode_ReturnsMostFrequentValue()
    {
        var ws = SampleSheet(out var wb);
        using (wb)
        {
            // Only 1 repeats.
            await Assert.That((double)ws.Evaluate("MODE(A1:A7)")).IsEqualTo(1d).Within(Tolerance);
            // Ties resolve to the value whose first occurrence is earliest.
            await Assert.That((double)ws.Evaluate("MODE(4, 4, 2, 2)")).IsEqualTo(4d).Within(Tolerance);
            await Assert.That((double)ws.Evaluate("MODE.SNGL(A1:A7)")).IsEqualTo(1d).Within(Tolerance);
        }
    }

    [Test]
    public async Task Mode_NoRepeats_ReturnsNotAvailable()
    {
        var ws = SampleSheet(out var wb);
        using (wb)
        {
            await Assert.That(ws.Evaluate("MODE(1, 2, 3, 4)")).IsEqualTo(XLError.NoValueAvailable);
        }
    }

    [Test]
    public async Task Percentile_InterpolatesBetweenRanks()
    {
        var ws = SampleSheet(out var wb);
        using (wb)
        {
            await Assert.That((double)ws.Evaluate("PERCENTILE(A1:A7, 0)")).IsEqualTo(1d).Within(Tolerance);
            await Assert.That((double)ws.Evaluate("PERCENTILE(A1:A7, 1)")).IsEqualTo(9d).Within(Tolerance);
            await Assert.That((double)ws.Evaluate("PERCENTILE(A1:A7, 0.5)")).IsEqualTo(3d).Within(Tolerance);
            await Assert.That((double)ws.Evaluate("PERCENTILE(A1:A7, 0.25)")).IsEqualTo(1.5d).Within(Tolerance);
            await Assert.That((double)ws.Evaluate("PERCENTILE.INC(A1:A7, 0.25)")).IsEqualTo(1.5d).Within(Tolerance);
        }
    }

    [Test]
    public async Task Percentile_OutOfRange_ReturnsNumberInvalid()
    {
        var ws = SampleSheet(out var wb);
        using (wb)
        {
            await Assert.That(ws.Evaluate("PERCENTILE(A1:A7, 1.1)")).IsEqualTo(XLError.NumberInvalid);
            await Assert.That(ws.Evaluate("PERCENTILE(A1:A7, -0.1)")).IsEqualTo(XLError.NumberInvalid);
        }
    }

    [Test]
    public async Task Quartile_MapsToInclusivePercentiles()
    {
        var ws = SampleSheet(out var wb);
        using (wb)
        {
            await Assert.That((double)ws.Evaluate("QUARTILE(A1:A7, 0)")).IsEqualTo(1d).Within(Tolerance);
            await Assert.That((double)ws.Evaluate("QUARTILE(A1:A7, 1)")).IsEqualTo(1.5d).Within(Tolerance);
            await Assert.That((double)ws.Evaluate("QUARTILE(A1:A7, 2)")).IsEqualTo(3d).Within(Tolerance);
            await Assert.That((double)ws.Evaluate("QUARTILE(A1:A7, 3)")).IsEqualTo(4.5d).Within(Tolerance);
            await Assert.That((double)ws.Evaluate("QUARTILE(A1:A7, 4)")).IsEqualTo(9d).Within(Tolerance);
            await Assert.That((double)ws.Evaluate("QUARTILE.INC(A1:A7, 3)")).IsEqualTo(4.5d).Within(Tolerance);
        }
    }

    [Test]
    public async Task Quartile_OutOfRange_ReturnsNumberInvalid()
    {
        var ws = SampleSheet(out var wb);
        using (wb)
        {
            await Assert.That(ws.Evaluate("QUARTILE(A1:A7, 5)")).IsEqualTo(XLError.NumberInvalid);
        }
    }
}
