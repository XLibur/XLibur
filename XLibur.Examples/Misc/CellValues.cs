using XLibur.Excel;
using System;

namespace XLibur.Examples.Misc;

public class CellValues : IXLExample
{
    public void Create(string filePath)
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Cell Values");

        // Set the titles
        ws.Cell(2, 2).Value = "Initial Value";
        ws.Cell(2, 3).Value = "Casting";
        ws.Cell(2, 4).Value = "Using Get...()";
        ws.Cell(2, 5).Value = "Using GetValue<T>()";
        ws.Cell(2, 6).Value = "GetString()";
        ws.Cell(2, 7).Value = "GetFormattedString()";

        //////////////////////////////////////////////////////////////////
        // DateTime

        // Fill a cell with a date
        var cellDateTime = ws.Cell(3, 2);
        cellDateTime.Value = new DateTime(2010, 9, 2, 0, 0, 0, DateTimeKind.Unspecified);
        cellDateTime.Style.DateFormat.Format = "yyyy-MMM-dd";

        // Extract the date in different ways
        DateTime dateTime1 = cellDateTime.Value;
        var dateTime2 = cellDateTime.GetDateTime();
        var dateTime3 = cellDateTime.GetValue<DateTime>();
        var dateTimeString = cellDateTime.GetString();
        var dateTimeFormattedString = cellDateTime.GetFormattedString();

        // Set the values back to cells
        // The apostrophe is to force XLibur to treat the date as a string
        ws.Cell(3, 3).Value = dateTime1;
        ws.Cell(3, 4).Value = dateTime2;
        ws.Cell(3, 5).Value = dateTime3;
        ws.Cell(3, 6).Value = "'" + dateTimeString;
        ws.Cell(3, 7).Value = "'" + dateTimeFormattedString;

        // Boolean

        // Fill a cell with a boolean
        var cellBoolean = ws.Cell(4, 2);
        cellBoolean.Value = true;

        // Extract the boolean in different ways
        var boolean1 = (bool)cellBoolean.Value;
        var boolean2 = cellBoolean.GetBoolean();
        var boolean3 = cellBoolean.GetValue<bool>();
        var booleanString = cellBoolean.GetString();
        var booleanFormattedString = cellBoolean.GetFormattedString();

        // Set the values back to cells
        // The apostrophe is to force XLibur to treat the boolean as a string
        ws.Cell(4, 3).Value = boolean1;
        ws.Cell(4, 4).Value = boolean2;
        ws.Cell(4, 5).Value = boolean3;
        ws.Cell(4, 6).Value = "'" + booleanString;
        ws.Cell(4, 7).Value = "'" + booleanFormattedString;

        // Double

        // Fill a cell with a double
        var cellDouble = ws.Cell(5, 2);
        cellDouble.Value = 1234.567;
        cellDouble.Style.NumberFormat.Format = "#,##0.00";

        // Extract the double in different ways
        var double1 = (double)cellDouble.Value;
        var double2 = cellDouble.GetDouble();
        var double3 = cellDouble.GetValue<double>();
        var doubleString = cellDouble.GetString();
        var doubleFormattedString = cellDouble.GetFormattedString();

        // Set the values back to cells
        // The apostrophe is to force XLibur to treat the double as a string
        ws.Cell(5, 3).Value = double1;
        ws.Cell(5, 4).Value = double2;
        ws.Cell(5, 5).Value = double3;
        ws.Cell(5, 6).Value = "'" + doubleString;
        ws.Cell(5, 7).Value = "'" + doubleFormattedString;

        // String

        // Fill a cell with a string
        var cellString = ws.Cell(6, 2);
        cellString.Value = "Test Case";

        // Extract the string in different ways
        var string1 = (string)cellString.Value;
        var string2 = cellString.GetText();
        var string3 = cellString.GetValue<string>();
        var stringString = cellString.GetString();
        var stringFormattedString = cellString.GetFormattedString();

        // Set the values back to cells
        ws.Cell(6, 3).Value = string1;
        ws.Cell(6, 4).Value = string2;
        ws.Cell(6, 5).Value = string3;
        ws.Cell(6, 6).Value = stringString;
        ws.Cell(6, 7).Value = stringFormattedString;

        // TimeSpan

        // Fill a cell with a timeSpan
        var cellTimeSpan = ws.Cell(7, 2);
        cellTimeSpan.Value = new TimeSpan(1, 2, 31, 45);

        // Extract the timeSpan in different ways
        TimeSpan timeSpan1 = cellTimeSpan.Value;
        var timeSpan2 = cellTimeSpan.GetTimeSpan();
        var timeSpan3 = cellTimeSpan.GetValue<TimeSpan>();
        var timeSpanString = "'" + cellTimeSpan.GetString();
        var timeSpanFormattedString = "'" + cellTimeSpan.GetFormattedString();

        // Set the values back to cells
        ws.Cell(7, 3).Value = timeSpan1;
        ws.Cell(7, 4).Value = timeSpan2;
        ws.Cell(7, 5).Value = timeSpan3;
        ws.Cell(7, 6).Value = timeSpanString;
        ws.Cell(7, 7).Value = timeSpanFormattedString;

        // XLError

        var cellError = ws.Cell(8, 2);
        cellError.Value = XLError.DivisionByZero;

        // Extract the error in different ways
        var error1 = (XLError)cellError.Value;
        var error2 = cellError.GetError();
        var error3 = cellError.GetValue<XLError>();
        var errorString = "'" + cellError.GetString();
        var errorFormattedString = "'" + cellError.GetFormattedString();

        // Set the values back to cells
        ws.Cell(8, 3).Value = error1;
        ws.Cell(8, 4).Value = error2;
        ws.Cell(8, 5).Value = error3;
        ws.Cell(8, 6).Value = errorString;
        ws.Cell(8, 7).Value = errorFormattedString;

        // Do some formatting
        ws.Columns("B:G").Width = 20;
        var rngTitle = ws.Range("B2:G2");
        rngTitle.Style.Font.Bold = true;
        rngTitle.Style.Fill.BackgroundColor = XLColor.Cyan;

        ws.Columns().AdjustToContents();

        ws = workbook.AddWorksheet("Test Whitespace");
        ws.FirstCell().Value = "'    ";

        ws = workbook.AddWorksheet("Errors");
        ws.Cell(2, 2).Value = "Error value";
        ws.Cell(2, 3).Value = "Formula error";

        ws.Cell(3, 2).Value = XLError.CellReference;
        ws.Cell(3, 3).FormulaA1 = "#REF!+1";

        ws.Cell(4, 2).Value = XLError.IncompatibleValue;
        ws.Cell(4, 3).FormulaA1 = "\"TRUE\"*1";

        ws.Cell(5, 2).Value = XLError.DivisionByZero;
        ws.Cell(5, 3).FormulaA1 = "1/0";

        ws.Cell(6, 2).Value = XLError.NameNotRecognized;
        ws.Cell(6, 3).FormulaA1 = "NONEXISTENT.FUNCTION()";

        ws.Cell(7, 2).Value = XLError.NoValueAvailable;
        ws.Cell(7, 3).FormulaA1 = "NA()";

        ws.Cell(8, 2).Value = XLError.NullValue;
        ws.Cell(8, 3).FormulaA1 = "#NULL!+1";

        ws.Cell(9, 2).Value = XLError.NumberInvalid;
        ws.Cell(9, 3).FormulaA1 = "#NUM!+1";

        workbook.SaveAs(filePath, true, true);
    }
}
