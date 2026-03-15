using XLibur.Extensions;
using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace XLibur.Excel;

internal static partial class XLCellFormulaShifter
{
    private static readonly Regex A1SimpleRegex = A1SimpleRegexGenerated();

    private static readonly Regex A1RowRegex = A1RowRegexGenerated();

    private static readonly Regex A1ColumnRegex = A1ColumnRegexGenerated();

    internal static string ShiftFormulaRows(string formulaA1, XLWorksheet worksheetInAction, XLRange shiftedRange,
        int rowsShifted)
    {
        if (string.IsNullOrWhiteSpace(formulaA1)) return string.Empty;

        var value = formulaA1;

        var regex = A1SimpleRegex;

        var sb = new StringBuilder();
        var lastIndex = 0;

        var shiftedWsName = shiftedRange.Worksheet.Name;
        foreach (var match in regex.Matches(value).Cast<Match>())
        {
            var matchString = match.Value;
            var matchIndex = match.Index;
            if (value.Substring(0, matchIndex).CharCount('"') % 2 == 0)
            {
                // Check that the match is not between quotes
                sb.Append(value.Substring(lastIndex, matchIndex - lastIndex));
                string sheetName;
                var useSheetName = false;
                if (matchString.Contains('!'))
                {
                    sheetName = matchString.Substring(0, matchString.IndexOf('!'));
                    if (sheetName[0] == '\'')
                        sheetName = sheetName.Substring(1, sheetName.Length - 2);
                    useSheetName = true;
                }
                else
                    sheetName = worksheetInAction.Name;

                if (String.Compare(sheetName, shiftedWsName, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    var rangeAddress = matchString.Substring(matchString.IndexOf('!') + 1);
                    if (!A1ColumnRegex.IsMatch(rangeAddress))
                    {
                        var matchRange = worksheetInAction.Workbook.Worksheet(sheetName).Range(rangeAddress);
                        if (shiftedRange.RangeAddress.FirstAddress.RowNumber <=
                            matchRange.RangeAddress.LastAddress.RowNumber
                            && shiftedRange.RangeAddress.FirstAddress.ColumnNumber <=
                            matchRange.RangeAddress.FirstAddress.ColumnNumber
                            && shiftedRange.RangeAddress.LastAddress.ColumnNumber >=
                            matchRange.RangeAddress.LastAddress.ColumnNumber)
                        {
                            if (useSheetName)
                            {
                                sb.Append(sheetName.EscapeSheetName());
                                sb.Append('!');
                            }

                            if (A1RowRegex.IsMatch(rangeAddress))
                            {
                                var rows = rangeAddress.Split(':');
                                var row1String = rows[0];
                                var row2String = rows[1];
                                string row1;
                                if (row1String[0] == '$')
                                {
                                    row1 = "$" +
                                           (XLHelper.TrimRowNumber(int.Parse(row1String.Substring(1)) + rowsShifted))
                                           .ToInvariantString();
                                }
                                else
                                    row1 = (XLHelper.TrimRowNumber(int.Parse(row1String) + rowsShifted))
                                        .ToInvariantString();

                                string row2;
                                if (row2String[0] == '$')
                                {
                                    row2 = "$" +
                                           (XLHelper.TrimRowNumber(int.Parse(row2String.Substring(1)) + rowsShifted))
                                           .ToInvariantString();
                                }
                                else
                                    row2 = (XLHelper.TrimRowNumber(int.Parse(row2String) + rowsShifted))
                                        .ToInvariantString();

                                sb.Append(row1);
                                sb.Append(':');
                                sb.Append(row2);
                            }
                            else if (shiftedRange.RangeAddress.FirstAddress.RowNumber <=
                                     matchRange.RangeAddress.FirstAddress.RowNumber)
                            {
                                if (rangeAddress.Contains(':'))
                                {
                                    sb.Append(
                                        new XLAddress(
                                            worksheetInAction,
                                            XLHelper.TrimRowNumber(matchRange.RangeAddress.FirstAddress.RowNumber +
                                                                   rowsShifted),
                                            matchRange.RangeAddress.FirstAddress.ColumnLetter,
                                            matchRange.RangeAddress.FirstAddress.FixedRow,
                                            matchRange.RangeAddress.FirstAddress.FixedColumn));
                                    sb.Append(':');
                                    sb.Append(
                                        new XLAddress(
                                            worksheetInAction,
                                            XLHelper.TrimRowNumber(matchRange.RangeAddress.LastAddress.RowNumber +
                                                                   rowsShifted),
                                            matchRange.RangeAddress.LastAddress.ColumnLetter,
                                            matchRange.RangeAddress.LastAddress.FixedRow,
                                            matchRange.RangeAddress.LastAddress.FixedColumn));
                                }
                                else
                                {
                                    sb.Append(
                                        new XLAddress(
                                            worksheetInAction,
                                            XLHelper.TrimRowNumber(matchRange.RangeAddress.FirstAddress.RowNumber +
                                                                   rowsShifted),
                                            matchRange.RangeAddress.FirstAddress.ColumnLetter,
                                            matchRange.RangeAddress.FirstAddress.FixedRow,
                                            matchRange.RangeAddress.FirstAddress.FixedColumn));
                                }
                            }
                            else
                            {
                                sb.Append(matchRange.RangeAddress.FirstAddress);
                                sb.Append(':');
                                sb.Append(
                                    new XLAddress(
                                        worksheetInAction,
                                        XLHelper.TrimRowNumber(matchRange.RangeAddress.LastAddress.RowNumber +
                                                               rowsShifted),
                                        matchRange.RangeAddress.LastAddress.ColumnLetter,
                                        matchRange.RangeAddress.LastAddress.FixedRow,
                                        matchRange.RangeAddress.LastAddress.FixedColumn));
                            }
                        }
                        else
                            sb.Append(matchString);
                    }
                    else
                        sb.Append(matchString);
                }
                else
                    sb.Append(matchString);
            }
            else
                sb.Append(value.Substring(lastIndex, matchIndex - lastIndex + matchString.Length));

            lastIndex = matchIndex + matchString.Length;
        }

        if (lastIndex < value.Length)
            sb.Append(value.Substring(lastIndex));

        return sb.ToString();
    }

    internal static string ShiftFormulaColumns(string formulaA1, XLWorksheet worksheetInAction, XLRange shiftedRange,
        int columnsShifted)
    {
        if (string.IsNullOrWhiteSpace(formulaA1)) return string.Empty;

        var value = formulaA1;

        var regex = A1SimpleRegex;

        var sb = new StringBuilder();
        var lastIndex = 0;

        foreach (var match in regex.Matches(value).Cast<Match>())
        {
            var matchString = match.Value;
            var matchIndex = match.Index;
            if (value.Substring(0, matchIndex).CharCount('"') % 2 == 0)
            {
                // Check that the match is not between quotes
                sb.Append(value.Substring(lastIndex, matchIndex - lastIndex));
                string sheetName;
                var useSheetName = false;
                if (matchString.Contains('!'))
                {
                    sheetName = matchString.Substring(0, matchString.IndexOf('!'));
                    if (sheetName[0] == '\'')
                        sheetName = sheetName.Substring(1, sheetName.Length - 2);
                    useSheetName = true;
                }
                else
                    sheetName = worksheetInAction.Name;

                if (String.Compare(sheetName, shiftedRange.Worksheet.Name, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    var rangeAddress = matchString[(matchString.IndexOf('!') + 1)..];
                    if (!A1RowRegex.IsMatch(rangeAddress))
                    {
                        var matchRange = worksheetInAction.Workbook.Worksheet(sheetName).Range(rangeAddress);

                        if (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <=
                            matchRange.RangeAddress.LastAddress.ColumnNumber
                            &&
                            shiftedRange.RangeAddress.FirstAddress.RowNumber <=
                            matchRange.RangeAddress.FirstAddress.RowNumber
                            &&
                            shiftedRange.RangeAddress.LastAddress.RowNumber >=
                            matchRange.RangeAddress.LastAddress.RowNumber)
                        {
                            if (useSheetName)
                            {
                                sb.Append(sheetName.EscapeSheetName());
                                sb.Append('!');
                            }

                            if (A1ColumnRegex.IsMatch(rangeAddress))
                            {
                                var columns = rangeAddress.Split(':');
                                var column1String = columns[0];
                                var column2String = columns[1];
                                string column1;
                                if (column1String[0] == '$')
                                {
                                    column1 = "$" +
                                              XLHelper.GetColumnLetterFromNumber(
                                                  XLHelper.GetColumnNumberFromLetter(
                                                      column1String.Substring(1)) + columnsShifted, true);
                                }
                                else
                                {
                                    column1 =
                                        XLHelper.GetColumnLetterFromNumber(
                                            XLHelper.GetColumnNumberFromLetter(column1String) +
                                            columnsShifted, true);
                                }

                                string column2;
                                if (column2String[0] == '$')
                                {
                                    column2 = "$" +
                                              XLHelper.GetColumnLetterFromNumber(
                                                  XLHelper.GetColumnNumberFromLetter(
                                                      column2String.Substring(1)) + columnsShifted, true);
                                }
                                else
                                {
                                    column2 =
                                        XLHelper.GetColumnLetterFromNumber(
                                            XLHelper.GetColumnNumberFromLetter(column2String) +
                                            columnsShifted, true);
                                }

                                sb.Append(column1);
                                sb.Append(':');
                                sb.Append(column2);
                            }
                            else if (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <=
                                     matchRange.RangeAddress.FirstAddress.ColumnNumber)
                            {
                                if (rangeAddress.Contains(':'))
                                {
                                    sb.Append(
                                        new XLAddress(
                                            worksheetInAction,
                                            matchRange.RangeAddress.FirstAddress.RowNumber,
                                            XLHelper.TrimColumnNumber(
                                                matchRange.RangeAddress.FirstAddress.ColumnNumber + columnsShifted),
                                            matchRange.RangeAddress.FirstAddress.FixedRow,
                                            matchRange.RangeAddress.FirstAddress.FixedColumn));
                                    sb.Append(':');
                                    sb.Append(
                                        new XLAddress(
                                            worksheetInAction,
                                            matchRange.RangeAddress.LastAddress.RowNumber,
                                            XLHelper.TrimColumnNumber(matchRange.RangeAddress.LastAddress.ColumnNumber +
                                                                      columnsShifted),
                                            matchRange.RangeAddress.LastAddress.FixedRow,
                                            matchRange.RangeAddress.LastAddress.FixedColumn));
                                }
                                else
                                {
                                    sb.Append(
                                        new XLAddress(
                                            worksheetInAction,
                                            matchRange.RangeAddress.FirstAddress.RowNumber,
                                            XLHelper.TrimColumnNumber(
                                                matchRange.RangeAddress.FirstAddress.ColumnNumber + columnsShifted),
                                            matchRange.RangeAddress.FirstAddress.FixedRow,
                                            matchRange.RangeAddress.FirstAddress.FixedColumn));
                                }
                            }
                            else
                            {
                                sb.Append(matchRange.RangeAddress.FirstAddress);
                                sb.Append(':');
                                sb.Append(
                                    new XLAddress(
                                        worksheetInAction,
                                        matchRange.RangeAddress.LastAddress.RowNumber,
                                        XLHelper.TrimColumnNumber(matchRange.RangeAddress.LastAddress.ColumnNumber +
                                                                  columnsShifted),
                                        matchRange.RangeAddress.LastAddress.FixedRow,
                                        matchRange.RangeAddress.LastAddress.FixedColumn));
                            }
                        }
                        else
                            sb.Append(matchString);
                    }
                    else
                        sb.Append(matchString);
                }
                else
                    sb.Append(matchString);
            }
            else
                sb.Append(value.Substring(lastIndex, matchIndex - lastIndex + matchString.Length));

            lastIndex = matchIndex + matchString.Length;
        }

        if (lastIndex < value.Length)
            sb.Append(value.Substring(lastIndex));

        return sb.ToString();
    }

    [GeneratedRegex(@"(\$?\d{1,7}:\$?\d{1,7})" // 1:1
        , RegexOptions.Compiled)]
    private static partial Regex A1RowRegexGenerated();

    [GeneratedRegex(@"(\$?[a-zA-Z]{1,3}:\$?[a-zA-Z]{1,3})" // A:A
        , RegexOptions.Compiled)]
    private static partial Regex A1ColumnRegexGenerated();

    [GeneratedRegex(
        @"(?<Reference>(?<Sheet>(\'([^\[\]\*/\\\?:\']+|\'\')\'|\'?\w+\'?)!)?(?<Range>(?<![\w\d])\$?[a-zA-Z]{1,3}\$?\d{1,7}(?<RangeEnd>:\$?[a-zA-Z]{1,3}\$?\d{1,7})?(?![\w\d])|(?<ColumnNumbers>\$?\d{1,7}:\$?\d{1,7})|(?<ColumnLetters>\$?[a-zA-Z]{1,3}:\$?[a-zA-Z]{1,3})))",
        RegexOptions.Compiled)]
    private static partial Regex A1SimpleRegexGenerated();
}
