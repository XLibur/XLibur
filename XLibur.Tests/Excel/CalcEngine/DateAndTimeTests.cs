using XLibur.Excel;
using System;
using XLibur.Extensions;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;

[SetCulture("en-US")]
public class DateAndTimeTests
{
    [Test]
    [Arguments(2008, 1, 1, 39448)]
    [Arguments(2008, 15, 1, 39873)]
    [Arguments(2008, -50, 1, 37895)]
    [Arguments(2008, 5, 63, 39631)]
    [Arguments(2008, 13, 63, 39876)]
    [Arguments(2008, 15, -120, 39752)]
    [Arguments(1900, 2, 29, 60)] // Loveable 29th feb 1900
    [Arguments(1900, 2, 28, 59)]
    [Arguments(1900, 1, 1, 1)]
    [Arguments(1900, 1, 0, 0)] // Excel formats it as 1900-01-00, but more like 1899-12-31
    [Arguments(1899, 1, 1, 693598)] // If year < 1900, add 1900 to it
    public async Task Date_returns_serial_date(int year, int month, int day, double expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"DATE({year},{month},{day})").GetNumber()).IsEqualTo(expected);
    }

    [Test]
    [Arguments(1900, 1, -1)] // Serial date -1, below 0
    [Arguments(9999, 12, 32)]
    public async Task Date_returns_error_when_result_outside_base_date_to_max_date_of_calendar_system(int year, int month, int day)
    {
        var actual = XLWorkbook.EvaluateExpr($"DATE({year},{month},{day})");
        await Assert.That(actual).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(-1, 32000, 1, 973586)]  // Year -1.1 behaves as -2
    [Arguments(-1.1, 32000, 1, 973221)]
    [Arguments(-2, 32000, 1, 973221)]
    [Arguments(2000, -5, 1, 36342)] // Month -5.1 behaves as -6
    [Arguments(2000, -5.1, 1, 36312)]
    [Arguments(2000, -6, 1, 36312)]
    [Arguments(2000, 2, -10, 36546)] // Day -10.1 behaves as -11
    [Arguments(2000, 2, -10.1, 36545)]
    [Arguments(2000, 2, -11, 36545)]
    public async Task Date_floors_arguments(double year, double month, double day, double expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"DATE({year},{month},{day})").GetNumber()).IsEqualTo(expected);
    }

    [Test]
    [Arguments(10000, -32767, 3, "7269-05-03")] // Month can be [-32767..32767)
    [Arguments(10000, -32767.1, 3, XLError.NumberInvalid)]
    [Arguments(2000, 32766.9, 1, "4730-06-01")]
    [Arguments(2000, 32767, 1, XLError.NumberInvalid)]
    [Arguments(2000, 1, 32767.9, "2089-09-16")] // Day is clamped to at most 32767
    [Arguments(2000, 1, 32768, "2089-09-16")]
    [Arguments(2000, 1, 1E+100, "2089-09-16")]
    [Arguments(2000, 1, -32768, "1910-04-14")] // When day is < -32768, day uses 32767 value instead
    [Arguments(2000, 1, -32768.1, "2089-09-16")]
    [Arguments(2000, 1, -1E+100, "2089-09-16")]
    [Arguments(10000, -32000, 1, "7333-04-01")] // Year is clamped to 10000
    [Arguments(10001, -32000, 1, "7333-04-01")]
    [Arguments(1E+100, -32000, 1, "7333-04-01")]
    [Arguments(-1E+100, 1, 1, XLError.NumberInvalid)] // Even if year is less than int.MinValue, there is no error
    public async Task Date_matches_excel_behavior_for_out_of_range_arguments(double year, double month, double day, object expectedResult)
    {
        if (expectedResult is string iso8601)
            expectedResult = DateTime.Parse(iso8601).ToSerialDateTime();

        var actual = XLWorkbook.EvaluateExpr($"DATE({year},{month},{day})");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(expectedResult));
    }

    [Test]
    [Arguments("1/1/2006", "12/12/2010", "Y", 4)]
    [Arguments("1/1/2006", "12/12/2010", "M", 59)]
    [Arguments("1/1/2006", "12/12/2010", "D", 1806)]
    [Arguments("1/1/2006", "12/12/2010", "MD", 11)]
    [Arguments("1/1/2006", "12/12/2010", "YM", 11)]
    [Arguments("1/1/2006", "12/12/2010", "YD", 345)]
    [Arguments(38718, 40524, "Y", 4)]
    [Arguments(38718, 40524, "M", 59)]
    [Arguments(38718, 40524, "D", 1806)]
    [Arguments(38718, 40524, "MD", 11)]
    [Arguments("2020-01-31", "2024-03-01", "MD", -1)] // Pathological case. Start is shifted to 2024-02-31, thus 2024-03-02 is one day before the end
    [Arguments("1990-01-20", "2002-12-15", "YM", 10)] // YM across many years
    [Arguments(38718, 40524, "YM", 11)]
    [Arguments(38718, 40524, "YD", 345)]
    [Arguments("2001-12-31", "2002-4-15", "YM", 3)] // YM counts only full months - the last month is not full
    [Arguments("2001-12-10", "2002-4-15", "YM", 4)] // YM counts only full months - the last month is full
    [Arguments("2001-12-15", "2002-4-15", "YM", 4)] // YM counts only full months - the last month exactly full
    [Arguments("1900-01-12", "1901-03-04", "YD", 51)] // YD has plus +1 error with start dates in jan/feb 1900 and end in march of subsequent years
    [Arguments("2001-12-31", "2002-4-15", "YD", 105)] // YD ignores year, baseline
    [Arguments("2001-12-31", "2003-4-15", "YD", 105)] // YD ignores year, different year
    [Arguments("2000-02-20", "2100-02-10", "YD", 356)] // YD uses start year, not end year. Start has feb29, baseline
    [Arguments("2001-02-20", "2100-02-10", "YD", 355)] // YD uses start year, not end year. Start doesn't have feb29 => it's one less than the baseline
    [Arguments("2002-01-31", "2002-4-15", "YD", 74)]
    [Arguments("2001-12-02", "2001-12-15", "Y", 0)]
    [Arguments("2001-12-02", "2002-12-02", "Y", 1)]
    [Arguments("2006-01-15", "2006-03-14", "M", 1)]
    [Arguments("2020-11-22", "2020-11-23 9:00", "D", 1)]
    public async Task DateDif(object startDate, object endDate, string unit, double expected)
    {
        if (startDate is string s1) startDate = $"\"{s1}\"";
        if (endDate is string s2) endDate = $"\"{s2}\"";
        await Assert.That((double)XLWorkbook.EvaluateExpr($"DATEDIF({startDate},{endDate},\"{unit}\")")).IsEqualTo(expected);
    }

    [Test]
    [Arguments("N")]
    public async Task DateDif_returns_number_error_on_unexpected_unit(string unit)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"DATEDIF(10,100,\"{unit}\")")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task DateDif_end_date_cant_be_after_start_date()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("DATEDIF(40524,38718,\"D\")")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(-0.1, 100)]
    [Arguments(1, 2958466)]
    public async Task DateDif_returns_number_error_on_date_out_of_date_system(decimal startDate, decimal endDate)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"DATEDIF({startDate},{endDate},\"D\")")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments("8/22/2008", 39682)]
    [Arguments("2/1/2006", 38749)]
    [Arguments("2006-2-1", 38749)]
    [Arguments("February 1, 2006 17:45", 38749)]
    public async Task DateValue_returns_truncated_serial_date_extracted_from_text(string date, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExprCurrent($"DATEVALUE(\"{date}\")")).IsEqualTo(expected);
    }

    [Test]
    [Arguments("\"100\"")]
    [Arguments("\"0\"")]
    public async Task DateValue_doesnt_coerce_number_in_a_text_to_a_date(string arg)
    {
        await Assert.That(XLWorkbook.EvaluateExprCurrent($"DATEVALUE({arg})")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments("TRUE")]
    [Arguments("FALSE")]
    [Arguments("1000")]
    [Arguments("DATE(2006,1,5)")]
    public async Task DateValue_returns_coercion_error_on_non_text(string arg)
    {
        await Assert.That(XLWorkbook.EvaluateExprCurrent($"DATEVALUE({arg})")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task DateValue_propagates_error()
    {
        await Assert.That(XLWorkbook.EvaluateExprCurrent("DATEVALUE(#DIV/0!)")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    [Arguments(0, 0)]
    [Arguments(1, 1)]
    [Arguments(31, 31)]
    [Arguments(32, 1)]
    [Arguments(59, 28)]
    [Arguments(60, 29)]
    [Arguments(61, 1)]
    [Arguments(30000, 18)]
    [Arguments(45718, 2)]
    public async Task Day_returns_day_of_a_month_for_serial_culture(double serialDate, double expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"DAY({serialDate})").GetNumber()).IsEqualTo(expected);
    }

    [Test]
    public async Task Day_only_accepts_serial_date_from_0_to_upper_limit_of_calendar_system()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("DAY(-0.1)")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr("DAY(DATE(9999,12,31)+1)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [SetCulture("eu-ES")]
    [Arguments("\"2006/1/2 10:45 AM\"", 2)]
    [Arguments("DATE(2006,1,2)", 2)]
    [Arguments("DATE(2006,0,2)", 2)]
    [Arguments("DATE(2013,9,0)", 31)]
    public async Task Day_examples(string date, double expected)
    {
        await Assert.That(XLWorkbook.EvaluateExprCurrent($"DAY({date})").GetNumber()).IsEqualTo(expected);
    }

    [Test]
    [Arguments(2016, 10, 1, 1992, 2, 29, 8981)]
    [Arguments(1901, 3, 10, 1900, 1, 26, 409)]
    public async Task Days_calculate_difference_between_two_dates(double endYear, double endMonth, double endDay, double startYear, double startMonth, double startDay, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"DAYS(DATE({endYear},{endMonth},{endDay}),DATE({startYear},{startMonth},{startDay}))")).IsEqualTo(expected);
    }

    [Test]
    [Arguments("2016-10-01", "1992-02-29", 8981)]
    [Arguments("1901-03-10", "1900-01-26", 409)]
    [Arguments("1900-01-26", "1901-03-10", -409)]
    public async Task Days_coerces_dates_to_number(string endDate, string startDate, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"DAYS(\"{endDate}\",\"{startDate}\")")).IsEqualTo(expected);
    }

    [Test]
    public async Task Days_truncates_passed_arguments()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("DAYS(10.6,1.9)")).IsEqualTo(9);
    }

    [Test]
    public async Task Days_arguments_must_be_in_date_range()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("DAYS(-0.1,1)")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr("DAYS(2958466,1)")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr("DAYS(1,-0.1)")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr("DAYS(1,2958466)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task Days360_uses_US_method_by_default()
    {
        const string formulaFormat = "DAYS360(DATE(2002,2,3),DATE(2005,5,31){0})";
        var defaultResult = XLWorkbook.EvaluateExpr(string.Format(formulaFormat, string.Empty));
        var usResult = XLWorkbook.EvaluateExpr(string.Format(formulaFormat, ",FALSE"));
        var euResult = XLWorkbook.EvaluateExpr(string.Format(formulaFormat, ",TRUE"));
        await Assert.That(defaultResult).IsEqualTo(1198);
        await Assert.That(defaultResult).IsEqualTo(usResult);
        await Assert.That(defaultResult).IsNotEqualTo(euResult);
    }

    [Test]
    public async Task Days360_Europe1()
    {
        var actual = XLWorkbook.EvaluateExpr("DAYS360(\"1/1/2008\", \"3/31/2008\",TRUE)");
        await Assert.That(actual).IsEqualTo(89);
    }

    [Test]
    public async Task Days360_Europe2()
    {
        var actual = XLWorkbook.EvaluateExpr("DAYS360(\"3/31/2008\", \"1/1/2008\",TRUE)");
        await Assert.That(actual).IsEqualTo(-89);
    }

    [Test]
    [Arguments(2002, 2, 3, 2005, 5, 31, 1198)]
    [Arguments(2005, 5, 31, 2002, 2, 3, -1197)]
    [Arguments(2008, 1, 1, 2008, 3, 31, 90)]
    [Arguments(2008, 3, 31, 2008, 1, 1, -89)]
    [Arguments(2020, 2, 29, 2021, 2, 28, 358)]
    [Arguments(2020, 5, 29, 2020, 4, 1, -58)]
    [Arguments(2020, 5, 29, 2020, 3, 31, -58)]
    [Arguments(2020, 5, 30, 2020, 4, 1, -59)]
    [Arguments(2020, 5, 30, 2020, 3, 31, -60)]
    [Arguments(2020, 5, 30, 2020, 3, 30, -60)]
    public async Task Days360_US_method(int startYear, int startMonth, int startDay, int endYear, int endMonth, int endDay, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"DAYS360(DATE({startYear},{startMonth},{startDay}),DATE({endYear},{endMonth},{endDay}),FALSE)")).IsEqualTo(expected);
    }

    [Test]
    [Arguments(1900, 2, 27, 1900, 2, 27, 0)]
    [Arguments(1900, 2, 27, 1900, 2, 28, 1)]
    [Arguments(1900, 2, 27, 1900, 2, 29, 2)]
    [Arguments(1900, 2, 27, 1900, 3, 1, 4)]
    [Arguments(1900, 2, 28, 1900, 2, 27, -1)]
    [Arguments(1900, 2, 28, 1900, 2, 28, 0)]
    [Arguments(1900, 2, 28, 1900, 2, 29, 1)]
    [Arguments(1900, 2, 28, 1900, 3, 1, 3)]
    [Arguments(1900, 2, 29, 1900, 2, 27, -3)]
    [Arguments(1900, 2, 29, 1900, 2, 28, -2)]
    [Arguments(1900, 2, 29, 1900, 2, 29, -1)]
    [Arguments(1900, 2, 29, 1900, 3, 1, 1)]
    [Arguments(1900, 3, 1, 1900, 2, 27, -4)]
    [Arguments(1900, 3, 1, 1900, 2, 28, -3)]
    [Arguments(1900, 3, 1, 1900, 2, 29, -2)]
    [Arguments(1900, 3, 1, 1900, 3, 1, 0)]
    public async Task Days360_US_method_for_feb_29_1900(int startYear, int startMonth, int startDay, int endYear, int endMonth, int endDay, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"DAYS360(DATE({startYear},{startMonth},{startDay}),DATE({endYear},{endMonth},{endDay}),FALSE)")).IsEqualTo(expected);
    }

    [Test]
    [Arguments("2008-03-01", -1, "2008-02-01")]
    [Arguments("2008-03-31", -1, "2008-02-29")]
    [Arguments("2008-03-01", 1, "2008-04-01")]
    [Arguments("2008-03-31", 1, "2008-04-30")]
    [Arguments("2008-03-01", -1, "2008-02-01")]
    [Arguments("2008-03-31", 1, "2008-04-30")]
    [Arguments("1900-01-31", 1, "1900-02-28")] // Uses correct FEB28
    [Arguments("1900-01-31", 2, "1900-03-31")]
    [Arguments("1983-07-31", -77, "1977-02-28")]
    [Arguments("2021-05-14", 35, "2024-04-14")]
    public async Task EDate_returns_end_date_from_start_date_and_month_offset(string startDate, double monthOffset, string expectedEndDate)
    {
        var actual = XLWorkbook.EvaluateExpr($"EDATE(\"{startDate}\",{monthOffset})");
        await Assert.That(actual).IsEqualTo(DateTime.Parse(expectedEndDate).ToSerialDateTime());
    }

    [Test]
    public async Task EDate_returns_number_error_for_non_date_values()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("EDATE(-0.1,0)")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr("EDATE(2958466,0)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments("1900-01-01", -1)]
    [Arguments("9999-07-10", 6)]
    [Arguments("9999-07-10", 1E+100)]
    public async Task EDate_returns_number_error_when_end_date_is_out_of_date_system(string startDate, double monthOffset)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"EDATE(\"{startDate}\",{monthOffset})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(1900, 1, 0, 0, 31)]
    [Arguments(1900, 1, 1, 0, 31)]
    [Arguments(1900, 1, 31, 0, 31)]
    [Arguments(1900, 2, 20, 0, 59)]
    [Arguments(1900, 2, 29, 0, 59)]
    [Arguments(1900, 2, 29, 1, 91)]
    [Arguments(1900, 2, 29, 1, 91)]
    [Arguments(1900, 3, 1, -1, 59)]
    [Arguments(1985, 4, 15, 9, 31443)]
    [Arguments(2006, 1, 31, 5, 38898)] // Spec examples
    [Arguments(2004, 2, 29, 12, 38411)]
    [Arguments(2004, 2, 28, 12, 38411)]
    [Arguments(2004, 1, 15, -23, 37315)]
    public async Task Eomonth_returns_end_of_month_from_start_date_plus_month_offset(int year, int month, int day, int months, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"EOMONTH(DATE({year},{month},{day}),{months})")).IsEqualTo(expected);
    }

    [Test]
    public async Task Eomonth_truncates_arguments()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("EOMONTH(60.1,0.9)")).IsEqualTo(59);
    }

    [Test]
    public async Task Eomonth_start_date_must_be_in_date_values()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("EOMONTH(-0.1,0)")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr("EOMONTH(DATE(9999,12,31)+1,0)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments("1900-01-01", -1)]
    [Arguments("9999-12-10", 1)]
    public async Task Eomonth_returns_number_error_when_end_date_is_out_of_date_system(string startDate, double monthOffset)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"EOMONTH(\"{startDate}\",{monthOffset})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments("0", 0)]
    [Arguments("0.25", 6)]
    [Arguments("0.5", 12)]
    [Arguments("0.75", 18)]
    [Arguments("1", 0)]
    [Arguments("1.75", 18)]
    [Arguments("\"7/18/2011 7:45\"", 7)]
    [Arguments("\"4/21/2012\"", 0)]
    [Arguments("\"12:00:00\"", 12)]
    [Arguments("\"8/22/2008 3:30:45 PM\"", 15, Skip = "We don't parse seconds")]
    [Arguments("\"8/22/2008 3:30 PM\"", 15)]
    [Arguments("DATE(2006,2,26)+TIME(2,10,20)", 2)]
    [Arguments("TIME(22,56,34)", 22)]
    [Arguments("\"22-Oct-2001 10:53:12\"", 10, Skip = "We don't parse seconds plus culture is wrong")]
    [Arguments("\"October 22, 2001 10:53\"", 10)]
    [Arguments("\"10:53:12 pm\"", 22)]
    [Arguments("\"22:53:12\"", 22)]
    public async Task Hour_returns_hour_of_serial_date(string dateArg, double expected)
    {
        await Assert.That(XLWorkbook.EvaluateExprCurrent($"HOUR({dateArg})").GetNumber()).IsEqualTo(expected);
    }

    [Test]
    public async Task Hour_accepts_only_serial_time_between_zero_and_upper_limit_of_date_system()
    {
        await Assert.That(XLWorkbook.EvaluateExprCurrent("HOUR(0)")).IsEqualTo(0);
        await Assert.That(XLWorkbook.EvaluateExprCurrent("HOUR(-0.1)")).IsEqualTo(XLError.NumberInvalid);

        await Assert.That(XLWorkbook.EvaluateExprCurrent("HOUR(DATE(9999,12,31)+0.9)")).IsEqualTo(21);
        await Assert.That(XLWorkbook.EvaluateExprCurrent("HOUR(DATE(9999,12,31)+1)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments("0", 0)]
    [Arguments("0.5", 0)]
    [Arguments("0.68", 19)]
    [Arguments("0.69", 33)]
    [Arguments("0.85", 24)]
    [Arguments("10.85", 24)]
    [Arguments("\"14:47:20\"", 47)]
    [Arguments("\"8/22/2008 3:30 AM\"", 30)]
    public async Task Minute_returns_minute_of_serial_date(string dateArg, double expected)
    {
        await Assert.That(XLWorkbook.EvaluateExprCurrent($"MINUTE({dateArg})").GetNumber()).IsEqualTo(expected);
    }

    [Test]
    public async Task Minute_accepts_only_serial_time_between_zero_and_upper_limit_of_date_system()
    {
        await Assert.That(XLWorkbook.EvaluateExprCurrent("MINUTE(0)")).IsEqualTo(0);
        await Assert.That(XLWorkbook.EvaluateExprCurrent("MINUTE(-0.1)")).IsEqualTo(XLError.NumberInvalid);

        await Assert.That(XLWorkbook.EvaluateExprCurrent("MINUTE(DATE(9999,12,31)+0.9)")).IsEqualTo(36);
        await Assert.That(XLWorkbook.EvaluateExprCurrent("MINUTE(DATE(9999,12,31)+1)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [SetCulture("eu-ES")]
    [Arguments(0, 1)] // 1900-01-00
    [Arguments(31, 1)] // 1900-01-31
    [Arguments(32, 2)] // 1900-02-01
    [Arguments(59, 2)] // 1900-02-28
    [Arguments(60, 2)] // 1900-02-29
    [Arguments(61, 3)] // 1900-03-01
    [Arguments("DATE(2006,1,2)", 1)]
    [Arguments("DATE(2006,0,2)", 12)]
    [Arguments("\"2006/1/2 10:45 AM\"", 1)]
    [Arguments("30000", 2)]
    [Arguments("45596", 10)]
    [Arguments("45596.9", 10)]
    [Arguments("45597", 11)]
    public async Task Month_returns_month_of_serial_date(object argument, double expected)
    {
        await Assert.That(XLWorkbook.EvaluateExprCurrent($"MONTH({argument})").GetNumber()).IsEqualTo(expected);
    }

    [Test]
    public async Task Month_serial_date_must_be_between_zero_and_upper_limit_of_date_system()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("MONTH(-0.1)")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr("MONTH(DATE(9999,12,31) + 0.9)")).IsEqualTo(12);
        await Assert.That(XLWorkbook.EvaluateExpr("MONTH(DATE(9999,12,31) + 1)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(1900, 1, 0, 52)]
    [Arguments(1900, 1, 1, 52)]
    [Arguments(1900, 1, 2, 1)]
    [Arguments(1900, 2, 28, 9)]
    [Arguments(1900, 2, 29, 9)]
    [Arguments(1900, 3, 1, 9)]
    [Arguments(2012, 1, 2, 1)]
    [Arguments(2012, 12, 31, 1)]
    [Arguments(2012, 3, 9, 10)]
    [Arguments(2014, 12, 12, 50)]
    [Arguments(9999, 12, 31, 52)]
    public async Task IsoWeekNum(int year, int month, int day, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"ISOWEEKNUM(DATE({year},{month},{day}))")).IsEqualTo(expected);
    }

    [Test]
    public async Task NetWorkDays_with_holidays()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.FirstCell().SetValue("Date")
            .CellBelow().SetValue(new DateTime(2008, 10, 1, 0, 0, 0, DateTimeKind.Unspecified))
            .CellBelow().SetValue(new DateTime(2009, 3, 1, 0, 0, 0, DateTimeKind.Unspecified))
            .CellBelow().SetValue(new DateTime(2008, 11, 26, 0, 0, 0, DateTimeKind.Unspecified))
            .CellBelow().SetValue(new DateTime(2008, 12, 4, 0, 0, 0, DateTimeKind.Unspecified))
            .CellBelow().SetValue(new DateTime(2009, 1, 21, 0, 0, 0, DateTimeKind.Unspecified))
            .CellBelow().SetValue(new DateTime(2009, 1, 4, 0, 0, 0, DateTimeKind.Unspecified)) // Holiday is on Sunday - do not count twice
            .CellBelow().SetValue(new DateTime(2009, 1, 6, 0, 0, 0, DateTimeKind.Unspecified))  // Workweek holiday is specified twice, shouldn't be counted twice
            .CellBelow().SetValue(new DateTime(2009, 1, 6, 0, 0, 0, DateTimeKind.Unspecified))
            .CellBelow().SetValue(new DateTime(2008, 9, 30, 0, 0, 0, DateTimeKind.Unspecified)) // Tuesday holiday just before the first date, shouldn't be counted
            .CellBelow().SetValue(new DateTime(2009, 3, 2, 0, 0, 0, DateTimeKind.Unspecified)) // Monday holiday just after the last date, shouldn't be counted
            ;
        var actual = ws.Evaluate("NETWORKDAYS(A2, A3, A4:A11)");
        await Assert.That(actual).IsEqualTo(104);
    }

    [Test]
    [Arguments("2024-10-01", "2024-10-01", 1)] // Tue-Tue
    [Arguments("2024-10-01", "2024-10-02", 2)] // Tue-Wed
    [Arguments("2024-10-01", "2024-10-03", 3)] // Tue-Thu
    [Arguments("2024-10-01", "2024-10-04", 4)] // Tue-Fri
    [Arguments("2024-10-01", "2024-10-05", 4)] // Tue-Sat
    [Arguments("2024-10-01", "2024-10-06", 4)] // Tue-Sun
    [Arguments("2024-10-01", "2024-10-07", 5)] // Tue-Mon
    [Arguments("2024-09-29", "2024-10-12", 10)] // Sun-Sat
    [Arguments("2024-09-29", "2024-10-13", 10)] // Sun-Sun
    [Arguments("2024-09-29", "2024-10-14", 11)] // Sun-Mon
    [Arguments("2024-09-29", "2024-10-15", 12)] // Sun-Tue
    [Arguments("2024-09-29", "2024-10-16", 13)] // Sun-Wed
    [Arguments("2024-09-29", "2024-10-17", 14)] // Sun-Thu
    [Arguments("2024-09-29", "2024-10-18", 15)] // Sun-Fri
    [Arguments("2024-09-29", "2024-10-19", 15)] // Sun-Sat
    public async Task NetWorkDays_non_full_weeks_are_counted_correctly(string startDate, string endDate, int expected)
    {
        var actual = XLWorkbook.EvaluateExpr($"NETWORKDAYS(\"{startDate}\", \"{endDate}\")");
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    [Culture("en-US")]
    public async Task NetWorkDays_with_end_date_earlier_than_start_date()
    {
        var actual = XLWorkbook.EvaluateExpr("NETWORKDAYS(\"3/01/2009\", \"10/01/2008\")");
        await Assert.That(actual).IsEqualTo(-108);

        actual = XLWorkbook.EvaluateExpr("NETWORKDAYS(\"2016-01-01\", \"2015-12-23\")");
        await Assert.That(actual).IsEqualTo(-8);
    }

    [Test]
    [Culture("en-US")]
    public async Task NetWorkDays_behavior()
    {
        using var wb = new XLWorkbook();
        var actual = wb.Evaluate("NETWORKDAYS(\"10/01/2008\", \"3/01/2009\", \"11/26/2008\")");
        await Assert.That(actual).IsEqualTo(107);

        // Example from specification. Except spec wrong. The value is 1 off from Excel value.
        await Assert.That(wb.Evaluate("NETWORKDAYS(DATE(2006, 1, 1), DATE(2006, 1, 31))")).IsEqualTo(22);
        await Assert.That(wb.Evaluate("NETWORKDAYS(DATE(2006, 1, 31), DATE(2006, 1, 1))")).IsEqualTo(-22);
        await Assert.That(wb.Evaluate("NETWORKDAYS(DATE(2006, 1, 1), DATE(2006, 2, 1), { \"2006-01-02\", \"2006-01-16\" })")).IsEqualTo(21);

        // Scalar number is accepted for holidays
        await Assert.That(wb.Evaluate("NETWORKDAYS(1, 10, 2)")).IsEqualTo(6);

        // Scalar logical causes conversion error
        await Assert.That(wb.Evaluate("NETWORKDAYS(TRUE, 10)")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(wb.Evaluate("NETWORKDAYS(0, TRUE)")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(wb.Evaluate("NETWORKDAYS(1, 10, TRUE)")).IsEqualTo(XLError.IncompatibleValue);

        // Scalar text is converted
        await Assert.That(wb.Evaluate("NETWORKDAYS(\"1\", \"10\", \"2\")")).IsEqualTo(6);
        await Assert.That(wb.Evaluate("NETWORKDAYS(1, 10, \"0 4/2\")")).IsEqualTo(6);
        await Assert.That(wb.Evaluate("NETWORKDAYS(1, 10, \"1900-01-02\")")).IsEqualTo(6);
        await Assert.That(wb.Evaluate("NETWORKDAYS(\"Text\", 10)")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(wb.Evaluate("NETWORKDAYS(1, \"Text\")")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(wb.Evaluate("NETWORKDAYS(1, 10, \"Text\")")).IsEqualTo(XLError.IncompatibleValue);

        // Array accepts numbers and converts text
        await Assert.That(wb.Evaluate("NETWORKDAYS(1, 10, {\"2\", 3})")).IsEqualTo(5);
        await Assert.That(wb.Evaluate("NETWORKDAYS(1, 10, {\"Text\"})")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(wb.Evaluate("NETWORKDAYS(1, 10, {TRUE})")).IsEqualTo(XLError.IncompatibleValue);

        // Same conversion logic applies to reference values
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = Blank.Value; // Ignored
        ws.Cell("A2").Value = false; // Causes conversion error
        ws.Cell("A3").Value = true; // Causes conversion error
        ws.Cell("A4").Value = 37147; // 2001-09-13
        ws.Cell("A5").Value = "2001-09-12"; // Monday
        ws.Cell("A6").Value = XLError.NoValueAvailable;

        await Assert.That(ws.Evaluate("NETWORKDAYS(\"2001-05-01\", \"2001-12-31\", A1)")).IsEqualTo(175);
        await Assert.That(ws.Evaluate("NETWORKDAYS(\"2001-05-01\", \"2001-12-31\", A1:A3)")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(ws.Evaluate("NETWORKDAYS(\"2001-05-01\",\"2001-12-31\", A4:A5)")).IsEqualTo(173);

        // Errors are propagated
        await Assert.That(wb.Evaluate("NETWORKDAYS(#N/A, 10)")).IsEqualTo(XLError.NoValueAvailable);
        await Assert.That(wb.Evaluate("NETWORKDAYS(1, #N/A)")).IsEqualTo(XLError.NoValueAvailable);
        await Assert.That(wb.Evaluate("NETWORKDAYS(1, 10, {#N/A})")).IsEqualTo(XLError.NoValueAvailable);
        await Assert.That(ws.Evaluate("NETWORKDAYS(1, 10, A6)")).IsEqualTo(XLError.NoValueAvailable);
    }

    [Test]
    [Arguments("0", 0)]
    [Arguments("\"3:30:45\"", 45)]
    public async Task Second_returns_minute_of_serial_date(string dateArg, double expected)
    {
        await Assert.That(XLWorkbook.EvaluateExprCurrent($"SECOND({dateArg})").GetNumber()).IsEqualTo(expected);
    }

    [Test]
    public async Task Second_accepts_only_serial_time_between_zero_and_upper_limit_of_date_system()
    {
        await Assert.That(XLWorkbook.EvaluateExprCurrent("SECOND(0)")).IsEqualTo(0);
        await Assert.That(XLWorkbook.EvaluateExprCurrent("SECOND(-0.1)")).IsEqualTo(XLError.NumberInvalid);

        await Assert.That(XLWorkbook.EvaluateExprCurrent("SECOND(DATE(9999,12,31)+0.9999)")).IsEqualTo(51);
        await Assert.That(XLWorkbook.EvaluateExprCurrent("SECOND(DATE(9999,12,31)+1)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(0, 0, 0, 0)]
    [Arguments(0, 0, 1, 0.0000115740740741)]
    [Arguments(0, 0, 2, 0.0000231481481481)]
    [Arguments(0, 0, 20, 0.0002314814814815)]
    [Arguments(2, 3, 20, 0.0856481481481481)]
    [Arguments(12, 0, 0, 0.5000000000000000)]
    [Arguments(23, 59, 59, 0.9999884259259260)]
    [Arguments(26, 120, 240, 0.1694444444444450)]
    [Arguments(1, 2, 3, 0.043090277777778)]
    [Arguments(1.9, 2.9, 3.9, 0.043090277777778)]
    [Arguments(24, 0, 0, 0)]
    [Arguments(0, 42 * 60, 0, 0.75)]
    [Arguments(0, 0, 60 * 60 * 3, 0.125)]
    [Arguments(120, 240, 347, 0.170682870370)]
    public async Task Time_returns_serial_date_time(double hour, double minute, double second, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"TIME({hour},{minute},{second})")).IsEqualTo(expected).Within(XLHelper.Epsilon);
    }

    [Test]
    [Arguments(-0.1, 0, 0)]
    [Arguments(32768, 0, 0)]
    [Arguments(0, -0.1, 0)]
    [Arguments(0, 32768, 0)]
    [Arguments(0, 0, -0.1)]
    [Arguments(0, 0, 32768)]
    public async Task Time_components_must_be_in_zero_to_32767_interval(double hour, double minute, double second)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"TIME({hour},{minute},{second})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments("2:24 AM", 0.1)]
    [Arguments("August 22, 2008 6:35 AM", 0.27430555555555558)]
    public async Task TimeValue_returns_time_component_of_serial_date_extracted_from_text(string time, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExprCurrent($"TIMEVALUE(\"{time}\")")).IsEqualTo(expected).Within(XLHelper.Epsilon);
    }

    [Test]
    [Arguments("\"10.5\"")]
    [Arguments("\"0\"")]
    public async Task TimeValue_doesnt_coerce_number_in_a_text_to_a_time(string numberText)
    {
        await Assert.That(XLWorkbook.EvaluateExprCurrent($"TIMEVALUE({numberText})")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments("TRUE")]
    [Arguments("FALSE")]
    [Arguments("0.25")]
    [Arguments("TIME(18,25,48)")]
    public async Task TimeValue_returns_coercion_error_on_non_text(string nonText)
    {
        await Assert.That(XLWorkbook.EvaluateExprCurrent($"TIMEVALUE({nonText})")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task TimeValue_propagates_error()
    {
        await Assert.That(XLWorkbook.EvaluateExprCurrent("TIMEVALUE(#DIV/0!)")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    public async Task Today()
    {
        var actual = (double)XLWorkbook.EvaluateExpr("TODAY()");
        await Assert.That(actual).IsEqualTo(DateTime.Today.ToSerialDateTime());
    }

    [Test]
    [Arguments("\"2/14/2008\"", 1, 5)]
    [Arguments("\"2/14/2008\"", 2, 4)]
    [Arguments("\"2/14/2008\"", 3, 3)]
    [Arguments("\"2/14/2008\"", 11, 4)]
    [Arguments("\"2/14/2008\"", 12, 3)]
    [Arguments("\"2/14/2008\"", 13, 2)]
    [Arguments("\"2/14/2008\"", 14, 1)]
    [Arguments("\"2/14/2008\"", 15, 7)]
    [Arguments("\"2/14/2008\"", 16, 6)]
    [Arguments("\"2/14/2008\"", 17, 5)]
    public async Task Weekday_calculates_week_day(string value, int flag, int expected)
    {
        var actual = XLWorkbook.EvaluateExpr($"WEEKDAY({value}, {flag})");
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    public async Task Weekday_without_flag()
    {
        var actual = XLWorkbook.EvaluateExpr("WEEKDAY(\"2/14/2008\")");
        await Assert.That(actual).IsEqualTo(5);
    }

    [Test]
    public async Task Weekday_behavior()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        ws.Cell("A1").Value = 45577;
        await Assert.That(ws.Evaluate("WEEKDAY(A1)")).IsEqualTo(7);

        // Time of the day doesn't matter, serial date is truncated
        await Assert.That(XLWorkbook.EvaluateExpr("WEEKDAY(45577.9, 1.9)")).IsEqualTo(7);

        await Assert.That(XLWorkbook.EvaluateExpr("WEEKDAY(0)")).IsEqualTo(7);
        await Assert.That(XLWorkbook.EvaluateExpr("WEEKDAY(-1)")).IsEqualTo(XLError.NumberInvalid);

        // Year 10k
        await Assert.That(XLWorkbook.EvaluateExpr("WEEKDAY(2958465)")).IsEqualTo(6);
        await Assert.That(XLWorkbook.EvaluateExpr("WEEKDAY(2958466)")).IsEqualTo(XLError.NumberInvalid);

        // Convert from logical/text to number
        await Assert.That(XLWorkbook.EvaluateExpr("WEEKDAY(TRUE)")).IsEqualTo(1);
        await Assert.That(XLWorkbook.EvaluateExpr("WEEKDAY(\"0 2/2\")")).IsEqualTo(1);
        await Assert.That(XLWorkbook.EvaluateExpr("WEEKDAY(1, TRUE)")).IsEqualTo(1);
        await Assert.That(XLWorkbook.EvaluateExpr("WEEKDAY(1, \"1 0/2\")")).IsEqualTo(1);
        await Assert.That(XLWorkbook.EvaluateExpr("WEEKDAY(\"text\")")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(XLWorkbook.EvaluateExpr("WEEKDAY(1, \"text\")")).IsEqualTo(XLError.IncompatibleValue);

        // Flag can only have some values
        await Assert.That(XLWorkbook.EvaluateExpr("WEEKDAY(1, 0)")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr("WEEKDAY(1, 4)")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr("WEEKDAY(1, 10)")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr("WEEKDAY(1, 18)")).IsEqualTo(XLError.NumberInvalid);

        // Error is propagated
        await Assert.That(XLWorkbook.EvaluateExpr("WEEKDAY(#N/A)")).IsEqualTo(XLError.NoValueAvailable);
        await Assert.That(XLWorkbook.EvaluateExpr("WEEKDAY(5, #N/A)")).IsEqualTo(XLError.NoValueAvailable);
    }

    [Test]
    [Arguments(1, 1986, 12, 27, 52)]
    [Arguments(1, 1986, 12, 28, 53)]
    [Arguments(1, 1986, 12, 31, 53)]
    [Arguments(1, 1987, 1, 1, 1)]
    [Arguments(1, 1987, 1, 3, 1)]
    [Arguments(1, 1987, 1, 4, 2)]
    [Arguments(1, 2000, 3, 9, 11)]
    [Arguments(1, 2002, 3, 9, 10)]
    [Arguments(1, 2003, 3, 9, 11)]
    [Arguments(1, 2004, 3, 9, 11)]
    [Arguments(1, 2005, 3, 9, 11)]
    [Arguments(1, 2006, 3, 9, 10)]
    [Arguments(1, 2007, 3, 9, 10)]
    [Arguments(1, 2008, 3, 9, 11)]
    [Arguments(1, 2009, 3, 9, 11)]
    [Arguments(2, 1988, 12, 25, 52)]
    [Arguments(2, 1988, 12, 26, 53)]
    [Arguments(2, 1988, 12, 31, 53)]
    [Arguments(2, 1989, 1, 1, 1)]
    [Arguments(2, 1989, 1, 2, 2)]
    [Arguments(2, 2000, 3, 9, 11)]
    [Arguments(2, 2001, 3, 9, 10)]
    [Arguments(2, 2002, 3, 9, 10)]
    [Arguments(2, 2003, 3, 9, 10)]
    [Arguments(2, 2004, 3, 9, 11)]
    [Arguments(2, 2005, 3, 9, 11)]
    [Arguments(2, 2006, 3, 9, 11)]
    [Arguments(2, 2007, 3, 9, 10)]
    [Arguments(2, 2008, 3, 9, 10)]
    [Arguments(2, 2009, 3, 9, 11)]
    [Arguments(11, 1990, 12, 23, 51)]
    [Arguments(11, 1990, 12, 24, 52)]
    [Arguments(11, 1990, 12, 30, 52)]
    [Arguments(11, 1990, 12, 31, 53)]
    [Arguments(11, 1991, 1, 1, 1)]
    [Arguments(11, 1991, 1, 6, 1)]
    [Arguments(11, 1991, 1, 7, 2)]
    [Arguments(12, 1992, 12, 28, 52)]
    [Arguments(12, 1992, 12, 29, 53)]
    [Arguments(12, 1992, 12, 31, 53)]
    [Arguments(12, 1993, 1, 1, 1)]
    [Arguments(12, 1993, 1, 4, 1)]
    [Arguments(12, 1993, 1, 5, 2)]
    [Arguments(13, 1994, 12, 27, 52)]
    [Arguments(13, 1994, 12, 28, 53)]
    [Arguments(13, 1994, 12, 31, 53)]
    [Arguments(13, 1995, 1, 1, 1)]
    [Arguments(13, 1995, 1, 3, 1)]
    [Arguments(13, 1995, 1, 4, 2)]
    [Arguments(14, 1999, 12, 29, 52)]
    [Arguments(14, 1999, 12, 30, 53)]
    [Arguments(14, 1999, 12, 31, 53)]
    [Arguments(14, 2000, 1, 1, 1)]
    [Arguments(14, 2000, 1, 5, 1)]
    [Arguments(14, 2000, 1, 6, 2)]
    [Arguments(15, 2004, 12, 24, 53)]
    [Arguments(15, 2004, 12, 30, 53)]
    [Arguments(15, 2004, 12, 31, 54)]
    [Arguments(15, 2005, 1, 1, 1)]
    [Arguments(15, 2005, 1, 6, 1)]
    [Arguments(15, 2005, 1, 7, 2)]
    [Arguments(16, 2008, 12, 26, 52)]
    [Arguments(16, 2008, 12, 27, 53)]
    [Arguments(16, 2008, 12, 31, 53)]
    [Arguments(16, 2009, 1, 1, 1)]
    [Arguments(16, 2009, 1, 2, 1)]
    [Arguments(16, 2009, 1, 3, 2)]
    [Arguments(16, 2009, 1, 9, 2)]
    [Arguments(17, 1929, 12, 21, 51)]
    [Arguments(17, 1929, 12, 22, 52)]
    [Arguments(17, 1929, 12, 28, 52)]
    [Arguments(17, 1929, 12, 29, 53)]
    [Arguments(17, 1929, 12, 31, 53)]
    [Arguments(17, 1930, 1, 1, 1)]
    [Arguments(17, 1930, 1, 4, 1)]
    [Arguments(17, 1930, 1, 5, 2)]
    [Arguments(17, 1930, 1, 11, 2)]
    [Arguments(21, 1964, 12, 27, 52)]
    [Arguments(21, 1964, 12, 28, 53)]
    [Arguments(21, 1964, 12, 31, 53)]
    [Arguments(21, 1965, 1, 1, 53)]
    [Arguments(21, 1965, 1, 3, 53)]
    [Arguments(21, 1965, 1, 4, 1)]
    [Arguments(21, 1968, 12, 29, 52)]
    [Arguments(21, 1968, 12, 30, 1)]
    [Arguments(21, 1968, 12, 31, 1)]
    [Arguments(21, 1969, 1, 1, 1)]
    [Arguments(21, 1969, 1, 5, 1)]
    [Arguments(21, 1969, 1, 6, 2)]
    public async Task Weeknum_returns_week_number_for_date(double weekStartFlag, double year, double month, double day, double expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"WEEKNUM(DATE({year},{month},{day}),{weekStartFlag})").GetNumber()).IsEqualTo(expected);
    }

    [Test]
    public async Task Weeknum_default_week_starts_on_sunday()
    {
        for (var day = 14; day <= 20; day++)
        {
            var defaultValue = XLWorkbook.EvaluateExpr($"WEEKNUM(DATE(1967,5,{day}))");
            var sundayValue = XLWorkbook.EvaluateExpr($"WEEKNUM(DATE(1967,5,{day}),1)");
            await Assert.That(defaultValue).IsEqualTo(sundayValue);
        }
    }

    [Test]
    [Arguments]
    public async Task Weeknum_match_excel_behavior_and_returns_zero_for_serial_date_zero_when_week_starts_on_sunday()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("WEEKNUM(0,1)")).IsEqualTo(0);
        await Assert.That(XLWorkbook.EvaluateExpr("WEEKNUM(0,17)")).IsEqualTo(0);
    }

    [Test]
    [Arguments]
    public async Task Weeknum_returns_number_invalid_error_on_non_serial_dates()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("WEEKNUM(-0.1)")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr("WEEKNUM(DATE(9999,12,31)+1)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(-5)]
    [Arguments(0)]
    [Arguments(3)]
    [Arguments(10)]
    [Arguments(18)]
    [Arguments(20)]
    [Arguments(22)]
    [Arguments(100)]
    public async Task Weeknum_returns_number_invalid_error_on_non_specified_flags(double flag)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"WEEKNUM(DATE(200,1,1),{flag})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task Workdays_MultipleHolidaysGiven()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue("Date")
            .CellBelow().SetValue(new DateTime(2008, 10, 1, 0, 0, 0, DateTimeKind.Unspecified))
            .CellBelow().SetValue(151)
            .CellBelow().SetValue(new DateTime(2008, 11, 26, 0, 0, 0, DateTimeKind.Unspecified))
            .CellBelow().SetValue(new DateTime(2008, 12, 4, 0, 0, 0, DateTimeKind.Unspecified))
            .CellBelow().SetValue(new DateTime(2009, 1, 21, 0, 0, 0, DateTimeKind.Unspecified));
        var actual = ws.Evaluate("Workday(A2,A3,A4:A6)");
        await Assert.That(actual).IsEqualTo(new DateTime(2009, 5, 5, 0, 0, 0, DateTimeKind.Unspecified).ToSerialDateTime());
    }

    [Test]
    public async Task Workdays_NoHolidaysGiven()
    {
        var actual = XLWorkbook.EvaluateExpr("Workday(\"10/01/2008\", 151)");
        await Assert.That(actual).IsEqualTo(new DateTime(2009, 4, 30, 0, 0, 0, DateTimeKind.Unspecified).ToSerialDateTime());

        actual = XLWorkbook.EvaluateExpr("Workday(\"2016-01-01\", -10)");
        await Assert.That(actual).IsEqualTo(new DateTime(2015, 12, 18, 0, 0, 0, DateTimeKind.Unspecified).ToSerialDateTime());
    }

    [Test]
    public async Task Workdays_OneHolidaysGiven()
    {
        var actual = XLWorkbook.EvaluateExpr("Workday(\"10/01/2008\", 152, \"11/26/2008\")");
        await Assert.That(actual).IsEqualTo(new DateTime(2009, 5, 4, 0, 0, 0, DateTimeKind.Unspecified).ToSerialDateTime());
    }

    [Test]
    [Arguments(0, 0, 0)]
    [Arguments(0, 1, 2)]
    [Arguments(1, 1, 2)]
    [Arguments(2, 1, 3)]
    [Arguments(0, 5, 6)]
    [Arguments(2, 8, 12)]
    [Arguments(10, -1, 9)]
    [Arguments(10, -2, 6)]
    [Arguments(10, -3, 5)]
    [Arguments(9, -1, 6)]
    [Arguments(9, -2, 5)]
    [Arguments(8, -1, 6)]
    [Arguments(7, -1, 6)]
    [Arguments(6, -1, 5)]
    public async Task Workdays(int startDate, int dayOffset, int expected)
    {
        var actual = XLWorkbook.EvaluateExpr($"WORKDAY({startDate}, {dayOffset})");
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    [Arguments(0, 1, new[] { 1 }, 2)]
    [Arguments(0, 1, new[] { 2 }, 3)]
    [Arguments(0, 5, new[] { 2, 4 }, 10)]
    [Arguments(0, 4, new[] { 2, 4, 6 }, 10)]
    [Arguments(0, 3, new[] { 2, 3, 4, 6 }, 10)]
    [Arguments(0, 2, new[] { 2, 3, 4, 5, 6 }, 10)]
    [Arguments(0, 1, new[] { 2, 3, 5 }, 4)]
    [Arguments(0, 2, new[] { 2, 3, 5 }, 6)]
    [Arguments(2, 1, new[] { 2 }, 3)]
    [Arguments(15, -1, new[] { 13 }, 12)] // 15 = Sunday
    [Arguments(100, -5, new[] { 82, 93, 94, 95, 94, 100 }, 88)]
    [Arguments(98, -2, new[] { 97 }, 95)]
    public async Task Workdays_with_holiday(int startDate, int dayOffset, int[] holidays, int expected)
    {
        var actual = XLWorkbook.EvaluateExpr($"WORKDAY({startDate}, {dayOffset}, {{{string.Join(",", holidays)}}})");
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    [Arguments("\"8/22/2008\"", 2008)]
    [Arguments("\"1/2/2006 10:45 AM\"", 2006)]
    [Arguments("0", 1900)]
    [Arguments("0.5", 1900)]
    [Arguments("1", 1900)]
    [Arguments("59", 1900)]
    [Arguments("60", 1900)]
    [Arguments("366", 1900)]
    [Arguments("367", 1901)]
    [Arguments("DATE(9999,12,31)+0.9", 9999)]
    [Arguments("DATE(9999,12,31)+1", XLError.NumberInvalid)]
    [Arguments("-1", XLError.NumberInvalid)]
    [Arguments("\"test\"", XLError.IncompatibleValue)]
    [Arguments("IF(TRUE,)", 1900)] // Blank
    [Arguments("TRUE", 1900)]
    [Arguments("FALSE", 1900)]
    [Arguments("#DIV/0!", XLError.DivisionByZero)]
    public async Task Year(string value, object expected)
    {
        var actual = XLWorkbook.EvaluateExpr($"YEAR({value})");
        await Assert.That(actual).IsEqualTo(XLCellValue.FromObject(expected));
    }

    [Test]
    public async Task Year_BlankValue()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = Blank.Value;
        ws.Cell("A2").FormulaA1 = "YEAR(A1)";
        var valueA2 = ws.Cell("A2").Value;
        await Assert.That(valueA2).IsEqualTo(1900);
    }

    [Test]
    [Arguments(0, 2008, 1, 1, 2008, 3, 31, 0.25)]
    [Arguments(0, 2008, 1, 1, 2013, 3, 31, 5.25)]
    [Arguments(1, 2008, 1, 1, 2008, 3, 31, 0.24590163934426229)]
    [Arguments(1, 2008, 1, 1, 2013, 3, 31, 5.24452554744526)]
    [Arguments(1, 1900, 1, 10, 2024, 2, 29, 124.137572279657)]
    [Arguments(1, 1924, 6, 25, 2025, 2, 28, 100.67763581705)]
    [Arguments(2, 2008, 1, 1, 2008, 3, 31, 0.25)]
    [Arguments(2, 2008, 1, 1, 2013, 3, 31, 5.32222222222222)]
    [Arguments(3, 2008, 1, 1, 2008, 3, 31, 0.24657534246575341)]
    [Arguments(3, 2008, 1, 1, 2013, 3, 31, 5.24931506849315)]
    [Arguments(4, 2008, 1, 1, 2008, 3, 31, 0.24722222222222223)]
    [Arguments(4, 2008, 1, 1, 2013, 3, 31, 5.24722222222222)]
    [Arguments(0, 2006, 1, 1, 2006, 3, 26, 0.23611111111)]
    [Arguments(0, 2006, 3, 26, 2006, 1, 1, 0.23611111111)]
    [Arguments(0, 2006, 1, 1, 2006, 7, 1, 0.5)]
    [Arguments(0, 2006, 1, 1, 2007, 9, 1, 1.6666666666)]
    [Arguments(1, 2006, 1, 1, 2006, 7, 1, 0.495890411)]
    [Arguments(2, 2006, 1, 1, 2006, 7, 1, 0.5027777778)]
    [Arguments(3, 2006, 1, 1, 2006, 7, 1, 0.495890411)]
    [Arguments(4, 2006, 1, 1, 2006, 7, 1, 0.5)]
    [Arguments(1, 2004, 3, 1, 2006, 3, 1, 1.9981751825)]
    public async Task YearFrac_calculates_fraction_of_a_year(double basis, double startYear, double startMonth, double startDay, double endYear, double endMonth, double endDay, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"YEARFRAC(DATE({startYear},{startMonth},{startDay}),DATE({endYear},{endMonth},{endDay}),{basis})")).IsEqualTo(expected).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task YearFrac_dates_must_fit_in_date_system_range()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("YEARFRAC(-0.1,10)")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr("YEARFRAC(0,-0.1)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task YearFrac_basis_must_be_between_0_and_4()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("YEARFRAC(0,10,-0.1)")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr("YEARFRAC(0,10,5)")).IsEqualTo(XLError.NumberInvalid);
    }
}
