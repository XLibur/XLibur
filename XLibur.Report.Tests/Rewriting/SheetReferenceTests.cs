using System.Threading.Tasks;
using XLibur.Report.Rewriting;

namespace XLibur.Report.Tests.Rewriting;

public class SheetReferenceTests
{
    private static SheetReference Parse(string text)
    {
        SheetReference.TryParse(text, out var reference);
        return reference;
    }

    [Test]
    [Arguments("Data!$B$3:$B$8", "Data", 3, 2, 8, 2)]
    [Arguments("Data!$B$3", "Data", 3, 2, 3, 2)]
    [Arguments("$B$3:$C$8", null, 3, 2, 8, 3)]
    [Arguments("B3:C8", null, 3, 2, 8, 3)]
    [Arguments("'Sales 2026'!$A$1:$A$2", "Sales 2026", 1, 1, 2, 1)]
    public async Task ParsesTheFormsAChartUses(
        string text, string? sheet, int firstRow, int firstColumn, int lastRow, int lastColumn)
    {
        await Assert.That(SheetReference.TryParse(text, out var reference)).IsTrue();
        await Assert.That(reference.SheetName).IsEqualTo(sheet);
        await Assert.That(reference.FirstRow).IsEqualTo(firstRow);
        await Assert.That(reference.FirstColumn).IsEqualTo(firstColumn);
        await Assert.That(reference.LastRow).IsEqualTo(lastRow);
        await Assert.That(reference.LastColumn).IsEqualTo(lastColumn);
    }

    /// <summary>A doubled quote is a literal one inside a sheet name.</summary>
    [Test]
    public async Task ParsesASheetNameContainingAQuote()
    {
        await Assert.That(SheetReference.TryParse("'Bob''s data'!$A$1", out var reference)).IsTrue();
        await Assert.That(reference.SheetName).IsEqualTo("Bob's data");
    }

    [Test]
    public async Task OrdersTheCornersWhicheverWayTheyWereWritten()
    {
        var reference = Parse("Data!$C$8:$B$3");

        await Assert.That(reference.FirstRow).IsEqualTo(3);
        await Assert.That(reference.FirstColumn).IsEqualTo(2);
        await Assert.That(reference.LastRow).IsEqualTo(8);
        await Assert.That(reference.LastColumn).IsEqualTo(3);
    }

    /// <summary>
    /// The forms with no fixed extent, or no single sheet, are refused rather than guessed at — a
    /// wrong reference is worse for a report than a stale one.
    /// </summary>
    [Test]
    [Arguments("")]
    [Arguments("   ")]
    [Arguments("Data!$B:$B")]
    [Arguments("Data!$3:$8")]
    [Arguments("(Data!$B$3,Data!$D$3)")]
    [Arguments("Data!$B$3,Data!$D$3")]
    [Arguments("First:Last!$B$3")]
    [Arguments("Data!nonsense")]
    [Arguments("Data!")]
    [Arguments("'unterminated!$B$3")]
    public async Task RefusesWhatItCannotSafelyRewrite(string text)
    {
        await Assert.That(SheetReference.TryParse(text, out _)).IsFalse();
    }

    [Test]
    [Arguments("Data", 3, 2, 8, 2, "Data!$B$3:$B$8")]
    [Arguments("Data", 3, 2, 3, 2, "Data!$B$3")]
    [Arguments(null, 3, 2, 8, 3, "$B$3:$C$8")]
    [Arguments("Sales 2026", 1, 1, 2, 1, "'Sales 2026'!$A$1:$A$2")]
    [Arguments("Bob's data", 1, 1, 1, 1, "'Bob''s data'!$A$1")]
    [Arguments("Plain_1.a", 1, 1, 1, 1, "Plain_1.a!$A$1")]
    public async Task WritesTheReferenceBackTheWayExcelStoresOne(
        string? sheet, int firstRow, int firstColumn, int lastRow, int lastColumn, string expected)
    {
        var reference = new SheetReference(sheet, firstRow, firstColumn, lastRow, lastColumn);

        await Assert.That(reference.ToText()).IsEqualTo(expected);
    }

    [Test]
    [Arguments("Data!$B$3:$B$8")]
    [Arguments("'Sales 2026'!$A$1:$A$2")]
    [Arguments("Data!$B$3")]
    public async Task ParsingAndWritingRoundTrip(string text)
    {
        await Assert.That(Parse(text).ToText()).IsEqualTo(text);
    }
}
