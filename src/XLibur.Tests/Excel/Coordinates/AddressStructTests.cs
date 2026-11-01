using System;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.Tables;

namespace XLibur.Tests.Excel.Coordinates;

/// <summary>
/// <see cref="IXLCell.Address"/>, <see cref="IXLAddressable.RangeAddress"/> and the range's
/// <see cref="IXLRangeAddress.FirstAddress"/> and <see cref="IXLRangeAddress.LastAddress"/> used
/// to return interfaces over internal structs, so every call boxed a copy: 40 bytes for a cell
/// address, 72 for a range address, 112 for <c>range.RangeAddress.FirstAddress.ColumnNumber</c>
/// (#699). They now return the public structs <see cref="XLAddress"/> and
/// <see cref="XLRangeAddress"/>.
/// </summary>
/// <remarks>
/// Every test reads through the public interfaces (<see cref="IXLCell"/>, <see cref="IXLRange"/>),
/// because code inside XLibur that holds the concrete types never boxed.
/// </remarks>
public class AddressStructTests
{
    private const int Calls = 10_000;

    [Test]
    public async Task ReadingACellAddressColumn_AllocatesNothing()
    {
        using var wb = new XLWorkbook();
        IXLCell cell = wb.AddWorksheet("Sheet1").Cell(3, 2);

        var (bytes, sum) = MeasureAllocation(() =>
        {
            var sum = 0L;
            for (var i = 0; i < Calls; i++)
                sum += cell.Address.ColumnNumber;
            return sum;
        });

        await Assert.That(sum).IsEqualTo(2L * Calls);
        await Assert.That(bytes).IsEqualTo(0);
    }

    [Test]
    public async Task ReadingACellAddressRowAndColumn_AllocatesNothing()
    {
        using var wb = new XLWorkbook();
        IXLCell cell = wb.AddWorksheet("Sheet1").Cell(3, 2);

        var (bytes, sum) = MeasureAllocation(() =>
        {
            var sum = 0L;
            for (var i = 0; i < Calls; i++)
                sum += cell.Address.RowNumber + cell.Address.ColumnNumber;
            return sum;
        });

        await Assert.That(sum).IsEqualTo(5L * Calls);
        await Assert.That(bytes).IsEqualTo(0);
    }

    [Test]
    public async Task ReadingARangeAddressCorners_AllocatesNothing()
    {
        using var wb = new XLWorkbook();
        IXLRange range = wb.AddWorksheet("Sheet1").Range("B2:D5");

        var (bytes, sum) = MeasureAllocation(() =>
        {
            var sum = 0L;
            for (var i = 0; i < Calls; i++)
                sum += range.RangeAddress.FirstAddress.ColumnNumber + range.RangeAddress.LastAddress.RowNumber;
            return sum;
        });

        await Assert.That(sum).IsEqualTo(7L * Calls);
        await Assert.That(bytes).IsEqualTo(0);
    }

    [Test]
    public async Task IntersectionAndRelative_ReturnTheStructWithoutAllocating()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        // Boxed once here, outside the measurement: the parameters are still interfaces.
        IXLRangeAddress range = ws.Range("B2:D5").RangeAddress;
        IXLRangeAddress other = ws.Range("C3:F9").RangeAddress;
        IXLRangeAddress source = ws.Range("A1:A1").RangeAddress;
        IXLRangeAddress target = ws.Range("K10:K10").RangeAddress;

        await Assert.That(range.Intersection(other).ToString()).IsEqualTo("C3:D5");
        await Assert.That(range.Relative(source, target).ToString()).IsEqualTo("L11:N14");

        var (bytes, sum) = MeasureAllocation(() =>
        {
            var sum = 0L;
            for (var i = 0; i < Calls; i++)
                sum += range.Intersection(other).FirstAddress.ColumnNumber
                       + range.Relative(source, target).LastAddress.RowNumber;
            return sum;
        });

        await Assert.That(sum).IsEqualTo(17L * Calls);
        await Assert.That(bytes).IsEqualTo(0);
    }

    [Test]
    public async Task ReadingTableFieldNamesAgain_AllocatesNothing()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "Name";
        ws.Cell("B1").Value = "Age";
        ws.Cell("A2").Value = "Ann";
        ws.Cell("B2").Value = 30;
        var table = (XLTable)ws.Range("A1:B2").CreateTable();

        // The table kept its last range address as an interface, and compared the current one
        // with it on every read, which boxed the current one each time.
        var (bytes, count) = MeasureAllocation(() =>
        {
            var count = 0L;
            for (var i = 0; i < Calls; i++)
                count += table.FieldNames.Count;
            return count;
        });

        await Assert.That(count).IsEqualTo(2L * Calls);
        await Assert.That(bytes).IsEqualTo(0);
    }

    /// <summary>
    /// <see cref="XLAddress"/> used to cache its A1 text, and a caller that kept one boxed
    /// <see cref="IXLAddress"/> formatted it for free after the first call. A readonly struct cannot
    /// cache, so each call now builds its string. It must build only that one string: building
    /// the row and column text first and concatenating it made up to three, and
    /// <c>string.Concat(object[])</c> boxed the chars and the column number as well.
    /// </summary>
    [Test]
    [Arguments("ToString(A1)")]
    [Arguments("ToStringRelative()")]
    [Arguments("ToStringFixed(A1)")]
    [Arguments("ToString(R1C1)")]
    [Arguments("ToStringFixed(R1C1)")]
    public async Task FormattingAHeldAddress_AllocatesOnlyTheResult(string form)
    {
        using var wb = new XLWorkbook();
        // Row and column both from 300 up, which .NET's cache of small numbers' text does not cover.
        IXLAddress address = wb.AddWorksheet("Sheet1").Cell(1000, 400).Address;
        Func<string> format = form switch
        {
            "ToString(A1)" => () => address.ToString(XLReferenceStyle.A1),
            "ToStringRelative()" => () => address.ToStringRelative(),
            "ToStringFixed(A1)" => () => address.ToStringFixed(XLReferenceStyle.A1),
            "ToString(R1C1)" => () => address.ToString(XLReferenceStyle.R1C1),
            _ => () => address.ToStringFixed(XLReferenceStyle.R1C1),
        };
        var text = format();

        var (bytes, length) = MeasureAllocation(() =>
        {
            var length = 0L;
            for (var i = 0; i < Calls; i++)
                length += format().Length;
            return length;
        });
        var (oneStringEach, _) = MeasureAllocation(() =>
        {
            var length = 0L;
            for (var i = 0; i < Calls; i++)
                length += new string('x', text.Length).Length;
            return length;
        });

        await Assert.That(length).IsEqualTo((long)text.Length * Calls);
        await Assert.That(bytes).IsEqualTo(oneStringEach);
    }

    [Test]
    public async Task CellAddresses_CompareByValue()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        // Before #699 each call boxed a new copy, and == compared the two references: always false.
        await Assert.That(ws.Cell("B3").Address == ws.Cell(3, 2).Address).IsTrue();
        await Assert.That(ws.Cell("B3").Address != ws.Cell(3, 2).Address).IsFalse();
        await Assert.That(ws.Cell("B3").Address == ws.Cell("C3").Address).IsFalse();
        await Assert.That(ws.Cell("B3").Address != ws.Cell("B4").Address).IsTrue();
    }

    [Test]
    public async Task CellAddressEquality_DoesNotCompareTheWorksheet()
    {
        using var wb = new XLWorkbook();
        var first = wb.AddWorksheet("Sheet1").Cell("B3").Address;
        var second = wb.AddWorksheet("Sheet2").Cell("B3").Address;

        // Row, column and the $ flags only, as XLAddress compared internally before it was public.
        await Assert.That(first == second).IsTrue();
        await Assert.That(first.Equals(second)).IsTrue();
        await Assert.That(first.GetHashCode()).IsEqualTo(second.GetHashCode());
    }

    [Test]
    public async Task RangeAddresses_CompareByValueAndWorksheet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var other = wb.AddWorksheet("Sheet2");

        await Assert.That(ws.Range("B2:D5").RangeAddress == ws.Range(2, 2, 5, 4).RangeAddress).IsTrue();
        await Assert.That(ws.Range("B2:D5").RangeAddress == ws.Range("B2:D6").RangeAddress).IsFalse();
        await Assert.That(ws.Range("B2:D5").RangeAddress != other.Range("B2:D5").RangeAddress).IsTrue();
    }

    [Test]
    public async Task TheStructs_StillServeAsTheInterfaces()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        IXLAddress address = ws.Cell("B3").Address;
        IXLRangeAddress rangeAddress = ws.Range("B2:D5").RangeAddress;

        await Assert.That(address.ToString()).IsEqualTo("B3");
        await Assert.That(address.Worksheet).IsSameReferenceAs(ws);
        await Assert.That(address.Equals(ws.Cell(3, 2).Address)).IsTrue();
        await Assert.That(rangeAddress.Contains(address)).IsTrue();
        await Assert.That(rangeAddress.FirstAddress.ToString()).IsEqualTo("B2");
        await Assert.That(rangeAddress.Worksheet).IsSameReferenceAs(ws);
        await Assert.That(ws.Range(ws.Cell("B2").Address, ws.Cell("D5").Address).RangeAddress.ToString())
            .IsEqualTo("B2:D5");
    }

    /// <summary>
    /// Runs <paramref name="loop"/> twice and returns the bytes the second run allocated on this
    /// thread, with its result. The first run warms up, so that tiered JIT and lazy
    /// initialisation are not counted.
    /// </summary>
    private static (long Bytes, long Result) MeasureAllocation(Func<long> loop)
    {
        loop();

        var before = GC.GetAllocatedBytesForCurrentThread();
        var result = loop();
        return (GC.GetAllocatedBytesForCurrentThread() - before, result);
    }
}
