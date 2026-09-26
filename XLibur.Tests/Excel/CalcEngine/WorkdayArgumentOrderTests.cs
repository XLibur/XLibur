using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// The order in which WORKDAY, WORKDAY.INTL, NETWORKDAYS and NETWORKDAYS.INTL check their
/// arguments, and so which error wins when more than one argument is bad. Every expected value was
/// read from Excel 16 over COM, with the formula in E10 of a sheet where A1 is <c>#N/A</c>, B1 is
/// text, B2 is 45300, C1 is -5 and D1 is empty. DATE(2024,1,1) is 45292, a Monday.
/// </summary>
[SetCulture("en-US")]
public class WorkdayArgumentOrderTests
{
    [Test]
    // A zero offset returns the start date before the weekend or holidays are looked at.
    [Arguments("WORKDAY(DATE(2024,1,1),0,#N/A)", 45292d)]
    [Arguments("WORKDAY(DATE(2024,1,1),0,A1:A2)", 45292d)]
    [Arguments("WORKDAY(DATE(2024,1,1),0,B1:B2)", 45292d)]
    [Arguments("WORKDAY(DATE(2024,1,1),0,C1)", 45292d)]
    [Arguments("WORKDAY(DATE(2024,1,1),0,\"abc\")", 45292d)]
    [Arguments("WORKDAY(DATE(2024,1,6),0)", 45297d)] // A Saturday.
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),0,1,#N/A)", 45292d)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),0,1,A1:A2)", 45292d)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),0,1,B1:B2)", 45292d)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),0,1,C1)", 45292d)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),0,\"0000000\")", 45292d)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),0,\"1111111\")", 45292d)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),0,\"abc\")", 45292d)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),0,99)", 45292d)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),0,#DIV/0!)", 45292d)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),0,99,C1)", 45292d)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),0.5,99)", 45292d)]
    // The offset is rounded down, so a negative fraction is a whole working day back.
    [Arguments("WORKDAY(DATE(2024,1,1),-0.5)", 45289d)]
    [Arguments("WORKDAY(DATE(2024,1,1),-0.0001)", 45289d)]
    [Arguments("WORKDAY(DATE(2024,1,1),-1.5)", 45288d)]
    [Arguments("WORKDAY(DATE(2024,1,1),1.9)", 45293d)]
    [Arguments("WORKDAY(DATE(2024,1,6),-0.5)", 45296d)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),-0.5)", 45289d)]
    // Weekend spellings that work.
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,TRUE)", 45293d)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,1.9)", 45293d)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,)", 45293d)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,{1,2})", 45293d)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,\"0000000\")", 45293d)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),TRUE)", 5d)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),)", 5d)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),{1,2})", 5d)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),1.9)", 5d)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),\"0000000\")", 5d)]
    // NETWORKDAYS.INTL accepts a weekend of all seven days; WORKDAY.INTL does not.
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),\"1111111\")", 0d)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,5),DATE(2024,1,1),\"1111111\")", 0d)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,14),\"1111111\")", 0d)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),\"1111111\",45300)", 0d)]
    // A holiday written as numeric text is read as a date.
    [Arguments("WORKDAY(DATE(2024,1,1),1,\"45293\")", 45294d)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,1,\"45293\")", 45294d)]
    [Arguments("NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,5),\"45293\")", 4d)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),1,\"45293\")", 4d)]
    public async Task ReturnsNumber(string formula, double expected)
    {
        await Assert.That(Evaluate(formula)).IsEqualTo((XLCellValue)expected);
    }

    [Test]
    // WORKDAY: start, offset, zero offset, then each holiday in order.
    [Arguments("WORKDAY(DATE(2024,1,1),1,#N/A)", XLError.NoValueAvailable)]
    [Arguments("WORKDAY(DATE(2024,1,1),1,B1:B2)", XLError.IncompatibleValue)]
    [Arguments("WORKDAY(DATE(2024,1,1),1,C1)", XLError.NumberInvalid)]
    [Arguments("WORKDAY(DATE(2024,1,1),1,TRUE)", XLError.IncompatibleValue)]
    [Arguments("WORKDAY(DATE(2024,1,1),1,{\"x\",#N/A})", XLError.IncompatibleValue)]
    [Arguments("WORKDAY(DATE(2024,1,1),1,{-5,#N/A})", XLError.NumberInvalid)]
    [Arguments("WORKDAY(DATE(2024,1,1),-0.5,#N/A)", XLError.NoValueAvailable)]
    [Arguments("WORKDAY(DATE(2024,1,1),\"x\",#DIV/0!)", XLError.IncompatibleValue)]
    [Arguments("WORKDAY(-1,0)", XLError.NumberInvalid)]
    // WORKDAY.INTL: an error given as the weekend becomes #NUM!.
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,#N/A)", XLError.NumberInvalid)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,#DIV/0!)", XLError.NumberInvalid)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,A1)", XLError.NumberInvalid)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,A1:A2)", XLError.NumberInvalid)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,D1)", XLError.NumberInvalid)] // An empty cell is not an omitted weekend.
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,0)", XLError.NumberInvalid)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,99)", XLError.NumberInvalid)]
    // A string that is not a usable mask is #VALUE!.
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,\"abc\")", XLError.IncompatibleValue)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,\"1111111\")", XLError.IncompatibleValue)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,\"1\")", XLError.IncompatibleValue)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,\"\")", XLError.IncompatibleValue)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,\"2000000\")", XLError.IncompatibleValue)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,B1)", XLError.IncompatibleValue)]
    // Holidays given as one value are converted before the weekend is read...
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,1,#N/A)", XLError.NoValueAvailable)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,99,#DIV/0!)", XLError.DivisionByZero)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,\"abc\",#DIV/0!)", XLError.DivisionByZero)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,\"1111111\",#DIV/0!)", XLError.DivisionByZero)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,99,A1)", XLError.NoValueAvailable)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,99,B1)", XLError.IncompatibleValue)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,99,\"abc\")", XLError.IncompatibleValue)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,99,TRUE)", XLError.IncompatibleValue)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,1,TRUE)", XLError.IncompatibleValue)]
    // ...but whether it is a valid date, and every value of a range or array, only after.
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,\"1111111\",C1)", XLError.IncompatibleValue)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,\"1111111\",-5)", XLError.IncompatibleValue)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,\"1111111\",45300)", XLError.IncompatibleValue)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,99,\"45300\")", XLError.NumberInvalid)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,99,B1:B2)", XLError.NumberInvalid)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,99,{\"x\"})", XLError.NumberInvalid)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,99,{#N/A})", XLError.NumberInvalid)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,99,{\"x\",#N/A})", XLError.NumberInvalid)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,\"1111111\",B1:B2)", XLError.IncompatibleValue)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,1,B1:B2)", XLError.IncompatibleValue)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,1,{\"x\",#N/A})", XLError.IncompatibleValue)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,1,{#N/A,\"x\"})", XLError.NoValueAvailable)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),1,1,{-5,#N/A})", XLError.NumberInvalid)]
    // The start date and offset come before everything else.
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),#N/A,#DIV/0!)", XLError.NoValueAvailable)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),\"x\",1,#DIV/0!)", XLError.IncompatibleValue)]
    [Arguments("WORKDAY.INTL(#N/A,\"x\",1,#DIV/0!)", XLError.NoValueAvailable)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),-0.5,99)", XLError.NumberInvalid)]
    [Arguments("WORKDAY.INTL(DATE(2024,1,1),-0.5,1,#N/A)", XLError.NoValueAvailable)]
    [Arguments("WORKDAY.INTL(-1,0,99)", XLError.NumberInvalid)]
    // NETWORKDAYS.INTL reads the weekend first, and an error given as the weekend becomes #VALUE!.
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),#N/A)", XLError.IncompatibleValue)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),#DIV/0!)", XLError.IncompatibleValue)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),A1)", XLError.IncompatibleValue)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),A1:A2)", XLError.IncompatibleValue)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),D1)", XLError.NumberInvalid)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),0)", XLError.NumberInvalid)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),99)", XLError.NumberInvalid)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),\"abc\")", XLError.IncompatibleValue)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),\"1\")", XLError.IncompatibleValue)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),\"\")", XLError.IncompatibleValue)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),B1)", XLError.IncompatibleValue)]
    [Arguments("NETWORKDAYS.INTL(#N/A,DATE(2024,1,5),\"abc\")", XLError.IncompatibleValue)]
    [Arguments("NETWORKDAYS.INTL(#N/A,DATE(2024,1,5),#DIV/0!)", XLError.IncompatibleValue)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),#N/A,#DIV/0!)", XLError.IncompatibleValue)]
    [Arguments("NETWORKDAYS.INTL(\"x\",DATE(2024,1,5),99)", XLError.NumberInvalid)]
    [Arguments("NETWORKDAYS.INTL(#N/A,DATE(2024,1,5),99,#DIV/0!)", XLError.NumberInvalid)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),#DIV/0!,99,#N/A)", XLError.NumberInvalid)]
    // Then the start date, the end date and the holidays.
    [Arguments("NETWORKDAYS.INTL(#N/A,DATE(2024,1,5),1)", XLError.NoValueAvailable)]
    [Arguments("NETWORKDAYS.INTL(#N/A,DATE(2024,1,5),\"1111111\")", XLError.NoValueAvailable)]
    [Arguments("NETWORKDAYS.INTL(#N/A,DATE(2024,1,5),1,#DIV/0!)", XLError.NoValueAvailable)]
    [Arguments("NETWORKDAYS.INTL(#N/A,#DIV/0!)", XLError.NoValueAvailable)]
    [Arguments("NETWORKDAYS.INTL(-1,#DIV/0!)", XLError.NumberInvalid)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),99,#DIV/0!)", XLError.NumberInvalid)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),\"abc\",#DIV/0!)", XLError.IncompatibleValue)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),99,\"abc\")", XLError.NumberInvalid)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),99,A1)", XLError.NumberInvalid)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),99,B1:B2)", XLError.NumberInvalid)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),99,C1)", XLError.NumberInvalid)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),\"1111111\",#DIV/0!)", XLError.DivisionByZero)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),\"1111111\",C1)", XLError.NumberInvalid)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),\"1111111\",#N/A)", XLError.NoValueAvailable)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),\"1111111\",B1:B2)", XLError.IncompatibleValue)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),\"0000000\",B1:B2)", XLError.IncompatibleValue)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),1,{\"x\",#N/A})", XLError.IncompatibleValue)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),1,{-5,#N/A})", XLError.NumberInvalid)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),1,TRUE)", XLError.IncompatibleValue)]
    [Arguments("NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,1),1,#N/A)", XLError.NoValueAvailable)]
    // NETWORKDAYS: start, end, then the holidays.
    [Arguments("NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,1),#N/A)", XLError.NoValueAvailable)]
    [Arguments("NETWORKDAYS(DATE(2024,1,1),#DIV/0!,#N/A)", XLError.DivisionByZero)]
    [Arguments("NETWORKDAYS(#N/A,#DIV/0!)", XLError.NoValueAvailable)]
    [Arguments("NETWORKDAYS(-1,#DIV/0!)", XLError.NumberInvalid)]
    [Arguments("NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,5),B1:B2)", XLError.IncompatibleValue)]
    [Arguments("NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,5),C1)", XLError.NumberInvalid)]
    [Arguments("NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,5),{\"x\",#N/A})", XLError.IncompatibleValue)]
    [Arguments("NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,5),{-5,#N/A})", XLError.NumberInvalid)]
    public async Task ReturnsError(string formula, XLError expected)
    {
        await Assert.That(Evaluate(formula)).IsEqualTo((XLCellValue)expected);
    }

    /// <summary>
    /// Evaluate the formula in E10, as Excel did, so a two-row range such as A1:A2 has no implicit
    /// intersection with the formula's row.
    /// </summary>
    private static XLCellValue Evaluate(string formula)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = "NA()";
        ws.Cell("B1").Value = "text";
        ws.Cell("B2").Value = 45300;
        ws.Cell("C1").Value = -5;
        ws.Cell("E10").FormulaA1 = formula;
        return ws.Cell("E10").Value;
    }
}
