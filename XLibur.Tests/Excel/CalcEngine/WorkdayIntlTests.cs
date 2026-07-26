using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// NETWORKDAYS.INTL and WORKDAY.INTL, which take a weekend as either one of Excel's numbered codes
/// or a seven-character Monday-to-Sunday mask. Expected values are the worked examples from
/// Microsoft's per-function documentation, or counted from a calendar and shown in the comment.
/// </summary>
[SetCulture("en-US")]
public class WorkdayIntlTests
{
    private static XLWorksheet NewSheet(out XLWorkbook wb)
    {
        wb = new XLWorkbook();
        return (XLWorksheet)wb.AddWorksheet("Sheet1");
    }

    [Test]
    // 2006-01-01 is a Sunday. To 2006-01-31 there are 22 weekdays with a Sat+Sun weekend.
    [Arguments("NETWORKDAYS.INTL(DATE(2006,1,1), DATE(2006,1,31))", 22d)]
    [Arguments("NETWORKDAYS.INTL(DATE(2006,1,1), DATE(2006,1,31), 1)", 22d)] // Code 1 is the default.
    [Arguments("NETWORKDAYS.INTL(DATE(2006,1,1), DATE(2006,1,31), 7)", 23d)] // Fri+Sat: 31 days less 4 Fridays and 4 Saturdays.
    [Arguments("NETWORKDAYS.INTL(DATE(2006,1,1), DATE(2006,1,31), 11)", 26d)] // Sunday only: 31 days less 5 Sundays.
    public async Task NetWorkDaysIntl_ReferenceExamplesFromExcelDocumentation(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected);
    }

    [Test]
    // A single calendar week, 2024-01-01 (Monday) to 2024-01-07 (Sunday), so each weekend code
    // simply removes its own days from seven.
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,1,7), 1)", 5d)] // Sat, Sun.
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,1,7), 2)", 5d)] // Sun, Mon.
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,1,7), 3)", 5d)] // Mon, Tue.
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,1,7), 11)", 6d)] // Sun.
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,1,7), 12)", 6d)] // Mon.
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,1,7), 17)", 6d)] // Sat.
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,1,7), \"0000011\")", 5d)] // Sat+Sun as a mask.
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,1,7), \"1000000\")", 6d)] // Monday only.
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,1,7), \"1111110\")", 1d)] // Only Sunday works.
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,1,7), \"0000000\")", 7d)] // No weekend at all.
    public async Task NetWorkDaysIntl_EveryWeekendCodeAndMask(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected);
    }

    [Test]
    public async Task NetWorkDaysIntl_MatchesNetWorkDaysForTheDefaultWeekend()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "NETWORKDAYS(DATE(2024,1,1), DATE(2024,12,31))";
            ws.Cell("A2").FormulaA1 = "NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,12,31))";
            ws.Cell("A3").FormulaA1 = "NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,12,31), \"0000011\")";

            await Assert.That((double)ws.Cell("A2").Value).IsEqualTo((double)ws.Cell("A1").Value);
            await Assert.That((double)ws.Cell("A3").Value).IsEqualTo((double)ws.Cell("A1").Value);
        }
    }

    [Test]
    public async Task NetWorkDaysIntl_CountsBackwardsAsANegativeNumber()
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr("NETWORKDAYS.INTL(DATE(2024,1,7), DATE(2024,1,1))")).IsEqualTo(-5d);
    }

    [Test]
    public async Task NetWorkDaysIntl_SubtractsHolidaysThatFallOnAWorkingDay()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "DATE(2024,1,3)"; // A Wednesday, so it counts.
            ws.Cell("A2").FormulaA1 = "DATE(2024,1,6)"; // A Saturday, already a weekend.
            ws.Cell("A3").FormulaA1 = "DATE(2024,1,3)"; // A duplicate of the first, only counted once.

            ws.Cell("C1").FormulaA1 = "NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,1,7), 1, A1:A3)";
            ws.Cell("C2").FormulaA1 = "NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,1,7), 1, A2)";

            await Assert.That((double)ws.Cell("C1").Value).IsEqualTo(4d);
            await Assert.That((double)ws.Cell("C2").Value).IsEqualTo(5d);
        }
    }

    [Test]
    public async Task NetWorkDaysIntl_SameDayCountsAsOneOrZero()
    {
        // 2024-01-03 is a Wednesday; 2024-01-06 is a Saturday.
        await Assert.That((double)XLWorkbook.EvaluateExpr("NETWORKDAYS.INTL(DATE(2024,1,3), DATE(2024,1,3))")).IsEqualTo(1d);
        await Assert.That((double)XLWorkbook.EvaluateExpr("NETWORKDAYS.INTL(DATE(2024,1,6), DATE(2024,1,6))")).IsEqualTo(0d);
    }

    [Test]
    // From 2012-01-01, a Sunday, with a Sunday-and-Monday weekend, five days work each week. Thirty
    // working days is therefore six whole weeks: the first ends Saturday 7 January, so the thirtieth
    // is the Saturday five weeks later, 11 February.
    [Arguments("WORKDAY.INTL(DATE(2012,1,1), 30, 2)", "2012-02-11")]
    // Weekend code 11 leaves Sunday alone, so six days work each week and ninety of them is fifteen
    // whole weeks. The first ends Saturday 7 January; the fifteenth ends fourteen weeks later.
    [Arguments("WORKDAY.INTL(DATE(2012,1,1), 90, 11)", "2012-04-14")]
    public async Task WorkdayIntl_LandsOnTheNthWorkingDay(string formula, string expected)
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = formula;
            ws.Cell("A2").FormulaA1 = "TEXT(A1, \"yyyy-mm-dd\")";

            await Assert.That(ws.Cell("A2").Value).IsEqualTo(expected);
        }
    }

    [Test]
    // From Monday 2024-01-01 with a Saturday+Sunday weekend, one working day on is Tuesday the 2nd
    // and five working days on is the following Monday, the 8th.
    [Arguments("WORKDAY.INTL(DATE(2024,1,1), 1)", "2024-01-02")]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1), 5)", "2024-01-08")]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1), 0)", "2024-01-01")] // Zero leaves the date alone.
    [Arguments("WORKDAY.INTL(DATE(2024,1,8), -5)", "2024-01-01")] // And a negative offset walks back.
    [Arguments("WORKDAY.INTL(DATE(2024,1,1), 1, 11)", "2024-01-02")] // Sunday-only weekend.
    [Arguments("WORKDAY.INTL(DATE(2024,1,5), 1, 11)", "2024-01-06")] // Saturday is a working day here.
    [Arguments("WORKDAY.INTL(DATE(2024,1,1), 1, \"1000000\")", "2024-01-02")] // Monday-only weekend.
    [Arguments("WORKDAY.INTL(DATE(2024,1,6), 0)", "2024-01-06")] // A weekend start is returned as is.
    public async Task WorkdayIntl_AdvancesByWorkingDays(string formula, string expected)
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = formula;
            ws.Cell("A2").FormulaA1 = "TEXT(A1, \"yyyy-mm-dd\")";

            await Assert.That(ws.Cell("A2").Value).IsEqualTo(expected);
        }
    }

    [Test]
    public async Task WorkdayIntl_SkipsHolidays()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "DATE(2024,1,2)"; // The Tuesday that would otherwise be the answer.
            ws.Cell("B1").FormulaA1 = "WORKDAY.INTL(DATE(2024,1,1), 1, 1, A1)";
            ws.Cell("B2").FormulaA1 = "TEXT(B1, \"yyyy-mm-dd\")";

            await Assert.That(ws.Cell("B2").Value).IsEqualTo("2024-01-03");
        }
    }

    [Test]
    public async Task WorkdayIntl_AndNetWorkDaysIntl_AgreeWithEachOther()
    {
        // Counting the working days up to the date WORKDAY.INTL lands on has to give that same
        // number back, plus the start day itself.
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "WORKDAY.INTL(DATE(2024,3,4), 17, 3)";
            ws.Cell("A2").FormulaA1 = "NETWORKDAYS.INTL(DATE(2024,3,4), A1, 3)";

            // 2024-03-04 is a Monday, which weekend code 3 (Mon+Tue) makes a weekend day, so the
            // range covers 17 working days and no more.
            await Assert.That((double)ws.Cell("A2").Value).IsEqualTo(17d);
        }
    }

    [Test]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,1,7), 8)")] // 8, 9 and 10 are not codes.
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,1,7), 0)")]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,1,7), 18)")]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,1,7), \"1111111\")")] // No working day left.
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,1,7), \"000001\")")] // Too short.
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,1,7), \"000001x\")")] // Not 0 or 1.
    [Arguments("WORKDAY.INTL(DATE(2024,1,1), 5, 8)")]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1), 5, \"1111111\")")]
    public async Task WeekendArgument_OutOfRangeReturnsNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task IntlFunctions_ReadTheirArgumentsFromCells()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "DATE(2024,1,1)";
            ws.Cell("A2").FormulaA1 = "DATE(2024,1,7)";
            ws.Cell("A3").Value = 11;

            ws.Cell("B1").FormulaA1 = "NETWORKDAYS.INTL(A1, A2, A3)";
            ws.Cell("B2").FormulaA1 = "WORKDAY.INTL(A1, 3, A3)";
            ws.Cell("B3").FormulaA1 = "TEXT(B2, \"yyyy-mm-dd\")";

            await Assert.That((double)ws.Cell("B1").Value).IsEqualTo(6d);
            await Assert.That(ws.Cell("B3").Value).IsEqualTo("2024-01-04");
        }
    }
}
