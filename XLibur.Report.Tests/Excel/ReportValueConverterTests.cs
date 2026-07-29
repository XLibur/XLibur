using System;
using System.Numerics;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Report.Excel;

namespace XLibur.Report.Tests.Excel;

public class ReportValueConverterTests
{
    private enum Region
    {
        North,
    }

    [Test]
    public async Task NullBecomesBlank()
    {
        await Assert.That(ReportValueConverter.ToCellValue(null).IsBlank).IsTrue();
    }

    [Test]
    public async Task TextStaysText()
    {
        var value = ReportValueConverter.ToCellValue("hello");

        await Assert.That(value.IsText).IsTrue();
        await Assert.That(value.GetText()).IsEqualTo("hello");
    }

    [Test]
    public async Task BooleanStaysLogical()
    {
        var value = ReportValueConverter.ToCellValue(true);

        await Assert.That(value.IsBoolean).IsTrue();
        await Assert.That(value.GetBoolean()).IsTrue();
    }

    /// <summary>
    /// The point of evaluating expressions to objects rather than strings: a money total has to
    /// reach Excel as a number so it can be summed, charted and number-formatted.
    /// </summary>
    [Test]
    [Arguments(9.5)]
    public async Task DecimalBecomesANumber(double expected)
    {
        var value = ReportValueConverter.ToCellValue((decimal)expected);

        await Assert.That(value.IsNumber).IsTrue();
        await Assert.That(value.GetNumber()).IsEqualTo(expected);
    }

    [Test]
    public async Task IntegerBecomesANumber()
    {
        var value = ReportValueConverter.ToCellValue(42);

        await Assert.That(value.IsNumber).IsTrue();
        await Assert.That(value.GetNumber()).IsEqualTo(42d);
    }

    [Test]
    public async Task LongBecomesANumber()
    {
        var value = ReportValueConverter.ToCellValue(9_000_000_000L);

        await Assert.That(value.IsNumber).IsTrue();
        await Assert.That(value.GetNumber()).IsEqualTo(9_000_000_000d);
    }

    [Test]
    public async Task BigIntegerBecomesANumber()
    {
        var value = ReportValueConverter.ToCellValue(new BigInteger(1234));

        await Assert.That(value.IsNumber).IsTrue();
        await Assert.That(value.GetNumber()).IsEqualTo(1234d);
    }

    [Test]
    public async Task DateTimeStaysADate()
    {
        var value = ReportValueConverter.ToCellValue(new DateTime(2026, 3, 14));

        await Assert.That(value.IsDateTime).IsTrue();
        await Assert.That(value.GetDateTime()).IsEqualTo(new DateTime(2026, 3, 14));
    }

    [Test]
    public async Task DateTimeOffsetBecomesItsDateTime()
    {
        var value = ReportValueConverter.ToCellValue(new DateTimeOffset(new DateTime(2026, 3, 14), TimeSpan.Zero));

        await Assert.That(value.IsDateTime).IsTrue();
        await Assert.That(value.GetDateTime()).IsEqualTo(new DateTime(2026, 3, 14));
    }

    [Test]
    public async Task TimeSpanStaysATimeSpan()
    {
        var value = ReportValueConverter.ToCellValue(TimeSpan.FromHours(2));

        await Assert.That(value.IsTimeSpan).IsTrue();
        await Assert.That(value.GetTimeSpan()).IsEqualTo(TimeSpan.FromHours(2));
    }

    [Test]
    public async Task ErrorStaysAnError()
    {
        var value = ReportValueConverter.ToCellValue(XLError.DivisionByZero);

        await Assert.That(value.IsError).IsTrue();
        await Assert.That(value.GetError()).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    public async Task CellValuePassesThrough()
    {
        XLCellValue original = 12.5;

        var value = ReportValueConverter.ToCellValue(original);

        await Assert.That(value.IsNumber).IsTrue();
        await Assert.That(value.GetNumber()).IsEqualTo(12.5);
    }

    /// <summary>An enum's name reads better in a report than its underlying number.</summary>
    [Test]
    public async Task EnumBecomesItsName()
    {
        var value = ReportValueConverter.ToCellValue(Region.North);

        await Assert.That(value.IsText).IsTrue();
        await Assert.That(value.GetText()).IsEqualTo("North");
    }

    [Test]
    public async Task UnrecognisedTypeFallsBackToText()
    {
        var value = ReportValueConverter.ToCellValue(new Uri("https://example.com/report"));

        await Assert.That(value.IsText).IsTrue();
        await Assert.That(value.GetText()).IsEqualTo("https://example.com/report");
    }

    [Test]
    public async Task GuidFallsBackToInvariantText()
    {
        var id = Guid.Parse("2f1c9c8e-1c1e-4a54-9b9a-6b0f4a3a1d55");

        var value = ReportValueConverter.ToCellValue(id);

        await Assert.That(value.IsText).IsTrue();
        await Assert.That(value.GetText()).IsEqualTo("2f1c9c8e-1c1e-4a54-9b9a-6b0f4a3a1d55");
    }
}
