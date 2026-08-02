using System;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Ranges;

/// <summary>
/// Pins which exception each bad address produces. Before this was fixed, a null address escaped
/// as <see cref="NullReferenceException"/> and an empty one — or an empty half such as "A1:" —
/// escaped as <see cref="IndexOutOfRangeException"/>, both of which are internal detail leaking
/// out of a public API rather than a contract a caller can act on.
///
/// These tests exist to stop that regressing, so they assert the exception type deliberately
/// rather than just asserting that something was thrown.
/// </summary>
public class XLRangeAddressArgumentTests
{
    [Test]
    public async Task NullAddressThrowsArgumentNullException()
    {
        using var wb = new XLWorkbook();
        var range = wb.AddWorksheet("Sheet1").Range("A1:C3");

        // Cast required: a bare null is ambiguous between Range(string) and Range(IXLRangeAddress).
        await Assert.That(() => range.Range((string)null!)).Throws<ArgumentNullException>();
    }

    [Test]
    public async Task EmptyAddressThrowsArgumentException()
    {
        using var wb = new XLWorkbook();
        var range = wb.AddWorksheet("Sheet1").Range("A1:C3");

        // ArgumentNullException derives from ArgumentException, so Throws<ArgumentException>
        // would also pass for the null case above. ThrowsExactly keeps the two apart.
        await Assert.That(() => range.Range("")).ThrowsExactly<ArgumentException>();
    }

    [Test]
    [Arguments(":")]
    [Arguments("A1:")]
    [Arguments(":B2")]
    [Arguments("$")]
    [Arguments("$:$")]
    public async Task HalfWrittenAddressThrowsFormatException(string address)
    {
        using var wb = new XLWorkbook();
        var range = wb.AddWorksheet("Sheet1").Range("A1:C3");

        await Assert.That(() => range.Range(address)).Throws<FormatException>();
    }

    [Test]
    public async Task MalformedAddressStillThrowsFormatException()
    {
        using var wb = new XLWorkbook();
        var range = wb.AddWorksheet("Sheet1").Range("A1:C3");

        await Assert.That(() => range.Range("not an address")).Throws<FormatException>();
    }

    [Test]
    public async Task AddressPastSheetLimitsStillThrowsOverflowException()
    {
        using var wb = new XLWorkbook();
        var range = wb.AddWorksheet("Sheet1").Range("A1:C3");

        await Assert.That(() => range.Range("ZZZZZZ9999999")).Throws<OverflowException>();
    }

    [Test]
    public async Task UnknownNameStillThrowsArgumentOutOfRangeException()
    {
        using var wb = new XLWorkbook();
        var range = wb.AddWorksheet("Sheet1").Range("A1:C3");

        await Assert.That(() => range.Range("NoSuchDefinedName")).Throws<ArgumentOutOfRangeException>();
    }

    /// <summary>
    /// The guards must not have narrowed what is accepted, so the shapes either side of them are
    /// checked too: a plain range, an anchored one, a sheet-qualified one and a single cell.
    /// </summary>
    [Test]
    [Arguments("B2:C3", "B2:C3")]
    [Arguments("$B$2:$C$3", "B2:C3")]
    [Arguments("Sheet1!B2:C3", "B2:C3")]
    [Arguments("B2", "B2:B2")]
    public async Task ValidAddressesAreUnaffected(string address, string expected)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var range = ws.Range("A1:C3");

        var resolved = range.Range(address);

        await Assert.That(resolved.RangeAddress.ToStringRelative()).IsEqualTo(expected);
    }

    /// <summary>
    /// Whitespace already threw <see cref="ArgumentException"/> before this change and still
    /// does — pinned so the new empty-string guard is not later widened to trim, which would
    /// silently reclassify it.
    /// </summary>
    [Test]
    public async Task WhitespaceAddressThrowsArgumentException()
    {
        using var wb = new XLWorkbook();
        var range = wb.AddWorksheet("Sheet1").Range("A1:C3");

        await Assert.That(() => range.Range(" ")).Throws<ArgumentException>();
    }
}
