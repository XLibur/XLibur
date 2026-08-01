using System;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Cells;
// ReSharper disable once InconsistentNaming
public class XLCellValueTests
{
    [Test]
    public async Task Creation_Blank()
    {
        XLCellValue blank = Blank.Value;
        await Assert.That(blank.Type).IsEqualTo(XLDataType.Blank);
        await Assert.That(blank.IsBlank).IsTrue();
    }

    [Test]
    public async Task Creation_Boolean()
    {
        XLCellValue logical = true;
        await Assert.That(logical.Type).IsEqualTo(XLDataType.Boolean);
        await Assert.That(logical.GetBoolean()).IsTrue();
        await Assert.That(logical.IsBoolean).IsTrue();
    }

    [Test]
    public async Task Creation_Number()
    {
        XLCellValue number = 14.0;
        await Assert.That(number.Type).IsEqualTo(XLDataType.Number);
        await Assert.That(number.IsNumber).IsTrue();
        await Assert.That(number.GetNumber()).IsEqualTo(14.0);
    }

    [Test]
    [Arguments(double.NaN)]
    [Arguments(double.PositiveInfinity)]
    [Arguments(double.NegativeInfinity)]
    public async Task Creation_Number_CantBeNonNumber(double nonNumber)
    {
        await Assert.That(() => _ = (XLCellValue)nonNumber).Throws<ArgumentException>();
    }

    // Decimal is not allowed as a member of an attribute, so TestCase can't be used.
    private static readonly object[] DecimalTestCases =
    [
        new object[] { 5.875m, 5.875d },
        new object[] { decimal.MaxValue, 7.922816251426434E+28 },
        new object[] { 1.0E-28m, 1.0000000000000001E-28d }
    ];

    [Test]
    [MethodDataSource(nameof(DecimalTestCases))]
    public async Task Creation_Decimal(decimal decimalNumber, double expectedNumber)
    {
        XLCellValue cellValue = decimalNumber;
        await Assert.That(cellValue.IsNumber).IsTrue();
        await Assert.That(cellValue.GetNumber()).IsEqualTo(expectedNumber);
    }

    [Test]
    public async Task Creation_Text()
    {
        XLCellValue text = "Hello World";
        await Assert.That(text.Type).IsEqualTo(XLDataType.Text);
        await Assert.That(text.GetText()).IsEqualTo("Hello World");
    }

    [Test]
    public async Task NullString_IsConvertedToBlank()
    {
        XLCellValue value = (string)null;
        await Assert.That(value.IsBlank).IsTrue();
        await Assert.That(value.IsText).IsFalse();
    }

    [Test]
    public async Task Creation_Text_HasLimitedLength()
    {
        var longText = new string('A', 32768);
        await Assert.That(() => _ = (XLCellValue)longText).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task Creation_Error()
    {
        XLCellValue error = XLError.NumberInvalid;
        await Assert.That(error.Type).IsEqualTo(XLDataType.Error);
        await Assert.That(error.IsError).IsTrue();
        await Assert.That(error.GetError()).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task Creation_DateTime()
    {
        XLCellValue dateTime = new DateTime(2021, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
        await Assert.That(dateTime.Type).IsEqualTo(XLDataType.DateTime);
        await Assert.That(dateTime.IsDateTime).IsTrue();
        await Assert.That(dateTime.GetDateTime()).IsEqualTo(new DateTime(2021, 1, 1, 0, 0, 0, DateTimeKind.Unspecified));
    }

    [Test]
    public async Task Creation_TimeSpan()
    {
        XLCellValue dateTime = new TimeSpan(10, 1, 2, 3, 456);
        await Assert.That(dateTime.Type).IsEqualTo(XLDataType.TimeSpan);
        await Assert.That(dateTime.IsTimeSpan).IsTrue();
        await Assert.That(dateTime.GetTimeSpan()).IsEqualTo(new TimeSpan(10, 1, 2, 3, 456));
    }

    [Test]
    public async Task Creation_FromObject()
    {
        await Assert.That(XLCellValue.FromObject(null).Type).IsEqualTo(XLDataType.Blank);
        await Assert.That(XLCellValue.FromObject(Blank.Value).Type).IsEqualTo(XLDataType.Blank);
        await Assert.That(XLCellValue.FromObject(true).Type).IsEqualTo(XLDataType.Boolean);
        await Assert.That(XLCellValue.FromObject("Hello World").Type).IsEqualTo(XLDataType.Text);
        await Assert.That(XLCellValue.FromObject(XLError.NumberInvalid).Type).IsEqualTo(XLDataType.Error);
        await Assert.That(XLCellValue.FromObject(new DateTime(2021, 1, 1, 0, 0, 0, DateTimeKind.Unspecified)).Type).IsEqualTo(XLDataType.DateTime);
        await Assert.That(XLCellValue.FromObject(new TimeSpan(10, 1, 2, 3, 456)).Type).IsEqualTo(XLDataType.TimeSpan);
        await Assert.That(XLCellValue.FromObject((sbyte)42).Type).IsEqualTo(XLDataType.Number);
        await Assert.That(XLCellValue.FromObject((byte)42).Type).IsEqualTo(XLDataType.Number);
        await Assert.That(XLCellValue.FromObject((short)42).Type).IsEqualTo(XLDataType.Number);
        await Assert.That(XLCellValue.FromObject((ushort)42).Type).IsEqualTo(XLDataType.Number);
        await Assert.That(XLCellValue.FromObject(42).Type).IsEqualTo(XLDataType.Number);
        await Assert.That(XLCellValue.FromObject((uint)42).Type).IsEqualTo(XLDataType.Number);
        await Assert.That(XLCellValue.FromObject((long)42).Type).IsEqualTo(XLDataType.Number);
        await Assert.That(XLCellValue.FromObject((ulong)42).Type).IsEqualTo(XLDataType.Number);
        await Assert.That(XLCellValue.FromObject((float)42).Type).IsEqualTo(XLDataType.Number);
        await Assert.That(XLCellValue.FromObject((double)42).Type).IsEqualTo(XLDataType.Number);
        await Assert.That(XLCellValue.FromObject((decimal)42).Type).IsEqualTo(XLDataType.Number);
        await Assert.That(XLCellValue.FromObject(DayOfWeek.Sunday).Type).IsEqualTo(XLDataType.Text);
    }

    [Test]
    public async Task NumberTypes_HaveUnambiguousConversion()
    {
        {
            sbyte sbyteNumber = 5;
            XLCellValue sbyteCellValue = sbyteNumber;
            await Assert.That(sbyteCellValue.IsNumber).IsTrue();
            await Assert.That(sbyteCellValue.GetNumber()).IsEqualTo(5d);
        }
        {
            byte byteNumber = 6;
            XLCellValue byteCellValue = byteNumber;
            await Assert.That(byteCellValue.IsNumber).IsTrue();
            await Assert.That(byteCellValue.GetNumber()).IsEqualTo(6d);
        }
        {
            short shortNumber = 7;
            XLCellValue shortCellValue = shortNumber;
            await Assert.That(shortCellValue.IsNumber).IsTrue();
            await Assert.That(shortCellValue.GetNumber()).IsEqualTo(7d);
        }
        {
            ushort ushortNumber = 8;
            XLCellValue ushortCellValue = ushortNumber;
            await Assert.That(ushortCellValue.IsNumber).IsTrue();
            await Assert.That(ushortCellValue.GetNumber()).IsEqualTo(8d);
        }
        {
            var intNumber = 9;
            XLCellValue intCellValue = intNumber;
            await Assert.That(intCellValue.IsNumber).IsTrue();
            await Assert.That(intCellValue.GetNumber()).IsEqualTo(9d);
        }
        {
            uint uintNumber = 10;
            XLCellValue uintCellValue = uintNumber;
            await Assert.That(uintCellValue.IsNumber).IsTrue();
            await Assert.That(uintCellValue.GetNumber()).IsEqualTo(10d);
        }
        {
            long longNumber = 11;
            XLCellValue longCellValue = longNumber;
            await Assert.That(longCellValue.IsNumber).IsTrue();
            await Assert.That(longCellValue.GetNumber()).IsEqualTo(11d);
        }
        {
            ulong ulongNumber = 12;
            XLCellValue ulongCellValue = ulongNumber;
            await Assert.That(ulongCellValue.IsNumber).IsTrue();
            await Assert.That(ulongCellValue.GetNumber()).IsEqualTo(12d);
        }
        {
            var floatNumber = 13.5f;
            XLCellValue floatCellValue = floatNumber;
            await Assert.That(floatCellValue.IsNumber).IsTrue();
            await Assert.That(floatCellValue.GetNumber()).IsEqualTo(13.5d);
        }
        {
            var doubleNumber = 14.5;
            XLCellValue doubleCellValue = doubleNumber;
            await Assert.That(doubleCellValue.IsNumber).IsTrue();
            await Assert.That(doubleCellValue.GetNumber()).IsEqualTo(14.5d);
        }
        {
            var decimalNumber = 15.75m;
            XLCellValue decimalCellValue = decimalNumber;
            await Assert.That(decimalCellValue.IsNumber).IsTrue();
            await Assert.That(decimalCellValue.GetNumber()).IsEqualTo(15.75d);
        }
    }

    [Test]
    [SuppressMessage("ReSharper", "ExpressionIsAlwaysNull")]
    public async Task NullableNumber_WithNullValue_AreConvertedToBlank()
    {
        {
            sbyte? sbyteNull = null;
            XLCellValue sbyteCellValue = sbyteNull;
            await Assert.That(sbyteCellValue.IsNumber).IsFalse();
            await Assert.That(sbyteCellValue.IsBlank).IsTrue();
        }
        {
            byte? byteNull = null;
            XLCellValue byteCellValue = byteNull;
            await Assert.That(byteCellValue.IsNumber).IsFalse();
            await Assert.That(byteCellValue.IsBlank).IsTrue();
        }
        {
            short? shortNull = null;
            XLCellValue shortCellValue = shortNull;
            await Assert.That(shortCellValue.IsNumber).IsFalse();
            await Assert.That(shortCellValue.IsBlank).IsTrue();
        }
        {
            ushort? ushortNull = null;
            XLCellValue ushortCellValue = ushortNull;
            await Assert.That(ushortCellValue.IsNumber).IsFalse();
            await Assert.That(ushortCellValue.IsBlank).IsTrue();
        }
        {
            int? intNull = null;
            XLCellValue intCellValue = intNull;
            await Assert.That(intCellValue.IsNumber).IsFalse();
            await Assert.That(intCellValue.IsBlank).IsTrue();
        }
        {
            uint? uintNull = null;
            XLCellValue uintCellValue = uintNull;
            await Assert.That(uintCellValue.IsNumber).IsFalse();
            await Assert.That(uintCellValue.IsBlank).IsTrue();
        }
        {
            long? longNull = null;
            XLCellValue longCellValue = longNull;
            await Assert.That(longCellValue.IsNumber).IsFalse();
            await Assert.That(longCellValue.IsBlank).IsTrue();
        }
        {
            ulong? ulongNull = null;
            XLCellValue ulongCellValue = ulongNull;
            await Assert.That(ulongCellValue.IsNumber).IsFalse();
            await Assert.That(ulongCellValue.IsBlank).IsTrue();
        }
        {
            float? floatValue = null;
            XLCellValue floatCellValue = floatValue;
            await Assert.That(floatCellValue.IsNumber).IsFalse();
            await Assert.That(floatCellValue.IsBlank).IsTrue();
        }
        {
            double? doubleValue = null;
            XLCellValue doubleCellValue = doubleValue;
            await Assert.That(doubleCellValue.IsNumber).IsFalse();
            await Assert.That(doubleCellValue.IsBlank).IsTrue();
        }
        {
            decimal? decimalValue = null;
            XLCellValue decimalCellValue = decimalValue;
            await Assert.That(decimalCellValue.IsNumber).IsFalse();
            await Assert.That(decimalCellValue.IsBlank).IsTrue();
        }
    }

    [Test]
    public async Task NullableNumber_WithNumberValue_AreConvertedToNumber()
    {
        {
            sbyte? sbyteNumber = 5;
            XLCellValue sbyteCellValue = sbyteNumber;
            await Assert.That(sbyteCellValue.IsNumber).IsTrue();
            await Assert.That(sbyteCellValue.GetNumber()).IsEqualTo(5d);
        }
        {
            byte? byteNumber = 6;
            XLCellValue byteCellValue = byteNumber;
            await Assert.That(byteCellValue.IsNumber).IsTrue();
            await Assert.That(byteCellValue.GetNumber()).IsEqualTo(6d);
        }
        {
            short? shortNumber = 7;
            XLCellValue shortCellValue = shortNumber;
            await Assert.That(shortCellValue.IsNumber).IsTrue();
            await Assert.That(shortCellValue.GetNumber()).IsEqualTo(7d);
        }
        {
            ushort? ushortNumber = 8;
            XLCellValue ushortCellValue = ushortNumber;
            await Assert.That(ushortCellValue.IsNumber).IsTrue();
            await Assert.That(ushortCellValue.GetNumber()).IsEqualTo(8d);
        }
        {
            int? intNumber = 9;
            XLCellValue intCellValue = intNumber;
            await Assert.That(intCellValue.IsNumber).IsTrue();
            await Assert.That(intCellValue.GetNumber()).IsEqualTo(9d);
        }
        {
            uint? uintNumber = 9;
            XLCellValue uintCellValue = uintNumber;
            await Assert.That(uintCellValue.IsNumber).IsTrue();
            await Assert.That(uintCellValue.GetNumber()).IsEqualTo(9d);
        }
        {
            long? longNumber = 10;
            XLCellValue longCellValue = longNumber;
            await Assert.That(longCellValue.IsNumber).IsTrue();
            await Assert.That(longCellValue.GetNumber()).IsEqualTo(10d);
        }
        {
            ulong? ulongNumber = 11;
            XLCellValue ulongCellValue = ulongNumber;
            await Assert.That(ulongCellValue.IsNumber).IsTrue();
            await Assert.That(ulongCellValue.GetNumber()).IsEqualTo(11d);
        }
        {
            float? floatNumber = 12.875f;
            XLCellValue floatCellValue = floatNumber;
            await Assert.That(floatCellValue.IsNumber).IsTrue();
            await Assert.That(floatCellValue.GetNumber()).IsEqualTo(12.875d);
        }
        {
            double? doubleNumber = 13.875d;
            XLCellValue doubleCellValue = doubleNumber;
            await Assert.That(doubleCellValue.IsNumber).IsTrue();
            await Assert.That(doubleCellValue.GetNumber()).IsEqualTo(13.875d);
        }
        {
            decimal? decimalNumber = 14.875m;
            XLCellValue decimalCellValue = decimalNumber;
            await Assert.That(decimalCellValue.IsNumber).IsTrue();
            await Assert.That(decimalCellValue.GetNumber()).IsEqualTo(14.875d);
        }
    }

    [Test]
    [SuppressMessage("ReSharper", "ExpressionIsAlwaysNull")]
    public async Task NullableDateTime_WithNullValue_IsConvertedToBlank()
    {
        DateTime? dateTimeNull = null;
        XLCellValue dateTimeCellValue = dateTimeNull;
        await Assert.That(dateTimeCellValue.IsDateTime).IsFalse();
        await Assert.That(dateTimeCellValue.IsBlank).IsTrue();
    }

    [Test]
    public async Task NullableDateTime_WithDateValue_IsConvertedToDateTime()
    {
        DateTime? dateTime = new DateTime(2020, 5, 14, 8, 14, 30, DateTimeKind.Unspecified);
        XLCellValue dateTimeCellValue = dateTime;
        await Assert.That(dateTimeCellValue.IsDateTime).IsTrue();
        await Assert.That(dateTimeCellValue.GetDateTime()).IsEqualTo(dateTime.Value);
    }

    [Test]
    [SuppressMessage("ReSharper", "ExpressionIsAlwaysNull")]
    public async Task NullableTimeSpan_WithNullValue_IsConvertedToBlank()
    {
        TimeSpan? timeSpanNull = null;
        XLCellValue timeSpanCellValue = timeSpanNull;
        await Assert.That(timeSpanCellValue.IsTimeSpan).IsFalse();
        await Assert.That(timeSpanCellValue.IsBlank).IsTrue();
    }

    [Test]
    public async Task NullableTimeSpan_WithTimeSpanValue_IsConvertedToTimeSpan()
    {
        TimeSpan? timeSpan = new TimeSpan(48, 12, 45, 30);
        XLCellValue timeSpanCellValue = timeSpan;
        await Assert.That(timeSpanCellValue.IsTimeSpan).IsTrue();
        await Assert.That(timeSpanCellValue.GetTimeSpan()).IsEqualTo(timeSpan.Value);
    }

    [Test]
    public async Task UnifiedNumber_IsFormOf_Number_DateTime_And_TimeSpan()
    {
        XLCellValue value = Blank.Value;
        await Assert.That(value.IsUnifiedNumber).IsFalse();

        value = true;
        await Assert.That(value.IsUnifiedNumber).IsFalse();

        value = 14;
        await Assert.That(value.IsUnifiedNumber).IsTrue();
        await Assert.That(value.GetUnifiedNumber()).IsEqualTo(14.0);

        value = new DateTime(1900, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
        await Assert.That(value.IsUnifiedNumber).IsTrue();
        await Assert.That(value.GetUnifiedNumber()).IsEqualTo(1.0);

        value = new TimeSpan(2, 12, 0, 0);
        await Assert.That(value.IsUnifiedNumber).IsTrue();
        await Assert.That(value.GetUnifiedNumber()).IsEqualTo(2.5);

        value = "Text";
        await Assert.That(value.IsUnifiedNumber).IsFalse();

        value = XLError.CellReference;
        await Assert.That(value.IsUnifiedNumber).IsFalse();
    }

    [Test]
    [Arguments("1900-01-01", 1)]
    [Arguments("1900-01-02", 2)]
    [Arguments("1900-02-01", 32)]
    [Arguments("1900-02-28", 59)] // Excel assumes 1900 was a leap year and 29.1.1900 existed
    [Arguments("1900-03-01", 61)]
    [Arguments("2017-01-01", 42736)]
    public async Task SerialDateTime(string dateString, double expectedSerial)
    {
        XLCellValue date = DateTime.Parse(dateString);
        await Assert.That(date.GetUnifiedNumber()).IsEqualTo(expectedSerial);
    }

    [Test]
    [SetCulture("cs-CZ")]
    public async Task ToString_RespectsCulture()
    {
        XLCellValue v = Blank.Value;
        await Assert.That(v.ToString()).IsEqualTo(string.Empty);

        v = true;
        await Assert.That(v.ToString()).IsEqualTo("TRUE");

        v = 25.4;
        await Assert.That(v.ToString()).IsEqualTo("25,4");

        v = "Hello";
        await Assert.That(v.ToString()).IsEqualTo("Hello");

        v = XLError.IncompatibleValue;
        await Assert.That(v.ToString()).IsEqualTo("#VALUE!");

        v = new DateTime(1900, 1, 2, 0, 0, 0, DateTimeKind.Unspecified);
        await Assert.That(v.ToString()).IsEqualTo("02.01.1900 0:00:00");

        v = new DateTime(1900, 3, 1, 4, 10, 5, DateTimeKind.Unspecified);
        await Assert.That(v.ToString()).IsEqualTo("01.03.1900 4:10:05");

        v = new TimeSpan(4, 5, 6, 7, 82);
        await Assert.That(v.ToString()).IsEqualTo("101:06:07,082");
    }

    [Test]
    public async Task TryConvert_Blank()
    {
        XLCellValue value = Blank.Value;
        await Assert.That(value.TryConvert(out Blank blank)).IsTrue();
        await Assert.That(blank).IsEqualTo(Blank.Value);

        value = string.Empty;
        await Assert.That(value.TryConvert(out blank)).IsTrue();
        await Assert.That(blank).IsEqualTo(Blank.Value);
    }

    [Test]
    public async Task TryConvert_Boolean()
    {
        XLCellValue value = true;
        await Assert.That(value.TryConvert(out bool boolean)).IsTrue();
        await Assert.That(boolean).IsTrue();

        value = "True";
        await Assert.That(value.TryConvert(out boolean)).IsTrue();
        await Assert.That(boolean).IsTrue();

        value = "False";
        await Assert.That(value.TryConvert(out boolean)).IsTrue();
        await Assert.That(boolean).IsFalse();

        value = 0;
        await Assert.That(value.TryConvert(out boolean)).IsTrue();
        await Assert.That(boolean).IsFalse();

        value = 0.001;
        await Assert.That(value.TryConvert(out boolean)).IsTrue();
        await Assert.That(boolean).IsTrue();
    }

    [Test]
    public async Task TryConvert_Number()
    {
        var c = CultureInfo.GetCultureInfo("cs-CZ");
        XLCellValue value = 5;
        await Assert.That(value.TryConvert(out double number, c)).IsTrue();
        await Assert.That(number).IsEqualTo(5.0);

        value = "1,5";
        await Assert.That(value.TryConvert(out number, c)).IsTrue();
        await Assert.That(number).IsEqualTo(1.5);

        value = "1 1/4";
        await Assert.That(value.TryConvert(out number, c)).IsTrue();
        await Assert.That(number).IsEqualTo(1.25);

        value = "3.1.1900";
        await Assert.That(value.TryConvert(out number, c)).IsTrue();
        await Assert.That(number).IsEqualTo(3);

        value = true;
        await Assert.That(value.TryConvert(out number, c)).IsTrue();
        await Assert.That(number).IsEqualTo(1.0);

        value = false;
        await Assert.That(value.TryConvert(out number, c)).IsTrue();
        await Assert.That(number).IsEqualTo(0.0);

        value = new DateTime(2020, 4, 5, 10, 14, 5, DateTimeKind.Unspecified);
        await Assert.That(value.TryConvert(out number, c)).IsTrue();
        await Assert.That(number).IsEqualTo(43926.42644675926);

        value = new TimeSpan(18, 0, 0);
        await Assert.That(value.TryConvert(out number, c)).IsTrue();
        await Assert.That(number).IsEqualTo(0.75);
    }

    [Test]
    public async Task TryConvert_DateTime()
    {
        XLCellValue v = new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
        await Assert.That(v.TryConvert(out DateTime dt)).IsTrue();
        await Assert.That(dt).IsEqualTo(new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Unspecified));

        var lastSerialDate = 2958465;
        v = lastSerialDate;
        await Assert.That(v.TryConvert(out dt)).IsTrue();
        await Assert.That(dt).IsEqualTo(new DateTime(9999, 12, 31, 0, 0, 0, DateTimeKind.Unspecified));

        v = lastSerialDate + 1;
        await Assert.That(v.TryConvert(out dt)).IsFalse();

        v = new TimeSpan(14, 0, 0, 0);
        await Assert.That(v.TryConvert(out dt)).IsTrue();
        await Assert.That(dt).IsEqualTo(new DateTime(1900, 1, 14, 0, 0, 0, DateTimeKind.Unspecified));
    }

    [Test]
    public async Task TryConvert_TimeSpan()
    {
        var c = CultureInfo.GetCultureInfo("cs-CZ");
        XLCellValue v = new TimeSpan(10, 15, 30);
        await Assert.That(v.TryConvert(out TimeSpan ts, c)).IsTrue();
        await Assert.That(ts).IsEqualTo(new TimeSpan(10, 15, 30));

        v = "26:15:30,5";
        await Assert.That(v.TryConvert(out ts, c)).IsTrue();
        await Assert.That(ts).IsEqualTo(new TimeSpan(1, 2, 15, 30, 500));

        v = 0.75;
        await Assert.That(v.TryConvert(out ts, c)).IsTrue();
        await Assert.That(ts).IsEqualTo(new TimeSpan(18, 0, 0));
    }

    [Test]
    [Arguments(1)]
    [Arguments(10)] // microsecond
    [Arguments(3000000001)] // 5 min 1 tick
    public async Task TimeSpan_can_have_sub_millisecond_precision(long ticks)
    {
        var subMsTimeSpan = TimeSpan.FromTicks(ticks);
        XLCellValue value = subMsTimeSpan;
        await Assert.That(value.GetTimeSpan()).IsEqualTo(subMsTimeSpan);
    }

    [Test]
    [Arguments(1)]
    [Arguments(10)] // microsecond
    [Arguments(3000000001)] // 5 min 1 tick
    public async Task TimeSpan_with_sub_millisecond_precision_is_written_and_loaded_correctly(long ticks)
    {
        // NetFx converts double to string using G15. Core changed it to G17, but XLibur still use G15.
        var subMsTimeSpan = TimeSpan.FromTicks(ticks);
        await TestHelper.CreateSaveLoadAssert(
            (_, ws) =>
            {
                ws.Cell("A1").Value = subMsTimeSpan;
            },
            async (_, ws) =>
            {
                var cellValue = ws.Cell("A1").CachedValue;
                await Assert.That(cellValue.GetTimeSpan()).IsEqualTo(subMsTimeSpan);
            });
    }

    [Test]
    [Arguments(long.MaxValue / (double)TimeSpan.TicksPerDay + 0.01)]
    [Arguments(long.MinValue / (double)TimeSpan.TicksPerDay - 0.01)]
    public async Task TimeSpan_throws_when_not_representable(double serialDateTime)
    {
        var value = XLCellValue.FromSerialTimeSpan(serialDateTime);
        var ex = await Assert.That(() => value.GetTimeSpan()).Throws<OverflowException>()!;
        await Assert.That(ex.Message).IsEqualTo("The serial date time value is too large to be represented in a TimeSpan.");
    }
}
