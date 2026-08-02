using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Ranges;

/// <summary>
/// Covers <see cref="XLRanges"/>'s value semantics and text form, which spec 05 C2 rewrote to stop
/// flattening the collection's per-worksheet indexes on every call.
/// <para>
/// Equality here means equality, not coverage: <c>IXLRanges.Contains(IXLRange)</c> asks whether some
/// range in the collection <em>covers</em> the argument, which is a different question. The
/// <see cref="EqualityIsNotCoverage"/> case is what stops the two being confused again, since the
/// faster implementation is the wrong one.
/// </para>
/// </summary>
public class XLRangesEqualityTests
{
    [Test]
    public async Task EqualCollectionsAreEqualRegardlessOfInsertionOrder()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var left = new XLRanges { ws.Range("A1:B2"), ws.Range("D4:E5") };
        var right = new XLRanges { ws.Range("D4:E5"), ws.Range("A1:B2") };

        await Assert.That(left.Equals(right)).IsTrue();
        await Assert.That(left.GetHashCode()).IsEqualTo(right.GetHashCode());
    }

    [Test]
    public async Task CollectionsOfDifferentSizeAreNotEqual()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var left = new XLRanges { ws.Range("A1:B2"), ws.Range("D4:E5") };
        var right = new XLRanges { ws.Range("A1:B2") };

        await Assert.That(left.Equals(right)).IsFalse();
        await Assert.That(right.Equals(left)).IsFalse();
    }

    [Test]
    public async Task CollectionsOfSameSizeWithDifferentRangesAreNotEqual()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var left = new XLRanges { ws.Range("A1:B2") };
        var right = new XLRanges { ws.Range("A1:B3") };

        await Assert.That(left.Equals(right)).IsFalse();
    }

    /// <summary>
    /// A1:B2 sits inside A1:C3, so a coverage test would call these equal. Equality must not.
    /// </summary>
    [Test]
    public async Task EqualityIsNotCoverage()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var contained = new XLRanges { ws.Range("A1:B2") };
        var container = new XLRanges { ws.Range("A1:C3") };

        await Assert.That(contained.Equals(container)).IsFalse();
    }

    [Test]
    public async Task NullIsNotEqual()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var ranges = new XLRanges { ws.Range("A1:B2") };

        object? nullObject = null;

        await Assert.That(ranges.Equals(null)).IsFalse();
        await Assert.That(ranges!.Equals(nullObject)).IsFalse();
        await Assert.That(ranges.Equals("not a range collection")).IsFalse();
    }

    [Test]
    public async Task RangesSpanningWorksheetsCompareEqual()
    {
        using var wb = new XLWorkbook();
        var first = wb.AddWorksheet("First");
        var second = wb.AddWorksheet("Second");

        var left = new XLRanges { first.Range("A1:B2"), second.Range("C3:D4") };
        var right = new XLRanges { first.Range("A1:B2"), second.Range("C3:D4") };

        await Assert.That(left.Equals(right)).IsTrue();
    }

    [Test]
    public async Task ToStringJoinsRangesWithCommas()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var ranges = new XLRanges { ws.Range("A1:B2"), ws.Range("D4:E5") };

        var text = ranges.ToString();

        await Assert.That(text).Contains("A1:B2");
        await Assert.That(text).Contains("D4:E5");
        await Assert.That(text).Contains(",");
        await Assert.That(text.EndsWith(',')).IsFalse();
    }

    [Test]
    public async Task ToStringOfSingleRangeHasNoSeparator()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var ranges = new XLRanges { ws.Range("A1:B2") };

        await Assert.That(ranges.ToString()).DoesNotContain(",");
    }

    [Test]
    public async Task ToStringOfEmptyCollectionIsEmpty()
    {
        var ranges = new XLRanges();

        await Assert.That(ranges.ToString()).IsEqualTo(string.Empty);
    }

    [Test]
    public async Task EmptyCollectionsAreEqual()
    {
        var left = new XLRanges();
        var right = new XLRanges();

        await Assert.That(left.Equals(right)).IsTrue();
        await Assert.That(left.GetHashCode()).IsEqualTo(right.GetHashCode());
    }
}
