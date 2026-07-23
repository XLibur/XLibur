using System;
using System.Text;
using System.Text.RegularExpressions;
using XLibur.Excel.Coordinates;
using XLibur.Extensions;

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
        var sb = new StringBuilder();
        var lastIndex = 0;
        var shiftedWsName = shiftedRange.Worksheet.Name;

        foreach (Match match in A1SimpleRegex.Matches(value))
        {
            var matchString = match.Value;
            var matchIndex = match.Index;
            if (value.AsSpan(0, matchIndex).Count('"') % 2 == 0)
            {
                sb.Append(value.AsSpan(lastIndex, matchIndex - lastIndex));
                var (sheetName, useSheetName) = ExtractSheetName(matchString, worksheetInAction);

                if (String.Compare(sheetName, shiftedWsName, StringComparison.OrdinalIgnoreCase) == 0)
                    AppendShiftedRowMatch(sb, matchString, sheetName, useSheetName, worksheetInAction, shiftedRange, rowsShifted);
                else
                    sb.Append(matchString);
            }
            else
                sb.Append(value.AsSpan(lastIndex, matchIndex - lastIndex + matchString.Length));

            lastIndex = matchIndex + matchString.Length;
        }

        if (lastIndex < value.Length)
            sb.Append(value.AsSpan(lastIndex));

        return sb.ToString();
    }

    private static (string sheetName, bool useSheetName) ExtractSheetName(string matchString, XLWorksheet worksheetInAction)
    {
        if (matchString.Contains('!'))
        {
            var sheetName = matchString.Substring(0, matchString.IndexOf('!'));
            if (sheetName[0] == '\'')
                sheetName = sheetName.Substring(1, sheetName.Length - 2).Replace("''", "'");
            return (sheetName, true);
        }

        return (worksheetInAction.Name, false);
    }

    private static void AppendShiftedRowMatch(StringBuilder sb, string matchString, string sheetName, bool useSheetName,
        XLWorksheet worksheetInAction, XLRange shiftedRange, int rowsShifted)
    {
        var rangeAddress = matchString.Substring(matchString.IndexOf('!') + 1);
        if (A1ColumnRegex.IsMatch(rangeAddress))
        {
            sb.Append(matchString);
            return;
        }

        var matchRange = worksheetInAction.Workbook.Worksheet(sheetName).Range(rangeAddress);
        if (!IsRowRangeWithinShiftedRange(shiftedRange, matchRange))
        {
            sb.Append(matchString);
            return;
        }

        if (useSheetName)
        {
            sb.Append(sheetName.EscapeSheetName());
            sb.Append('!');
        }

        if (A1RowRegex.IsMatch(rangeAddress))
            AppendShiftedRowOnlyRange(sb, rangeAddress, rowsShifted);
        else if (shiftedRange.RangeAddress.FirstAddress.RowNumber <= matchRange.RangeAddress.FirstAddress.RowNumber)
        {
            if (IsTopBoundaryDeletion(shiftedRange, matchRange, rowsShifted))
                AppendClampedTopRowShift(sb, worksheetInAction, shiftedRange, matchRange, rowsShifted);
            else
                AppendShiftedRowCellRange(sb, worksheetInAction, matchRange, rangeAddress, rowsShifted);
        }
        else
            AppendPartialRowShift(sb, worksheetInAction, matchRange, rowsShifted);
    }

    /// <summary>
    /// True when a row deletion removes the top boundary of <paramref name="matchRange"/> while some
    /// rows below the deletion survive. Excel keeps the range's top row fixed at the deletion start and
    /// shifts only the bottom up (shrink + shift), e.g. deleting row 3 turns A3:A4 into A3:A3. Shifting
    /// both endpoints (as for a range that is entirely below the deletion) would instead expand the range
    /// upward to A2:A3. See issue #2866.
    /// </summary>
    private static bool IsTopBoundaryDeletion(XLRange shiftedRange, IXLRange matchRange, int rowsShifted)
    {
        return rowsShifted < 0
            && matchRange.RangeAddress.FirstAddress.RowNumber <= shiftedRange.RangeAddress.LastAddress.RowNumber
            && matchRange.RangeAddress.LastAddress.RowNumber > shiftedRange.RangeAddress.LastAddress.RowNumber;
    }

    private static void AppendClampedTopRowShift(StringBuilder sb, XLWorksheet ws, XLRange shiftedRange,
        IXLRange matchRange, int rowsShifted)
    {
        sb.Append(new XLAddress(ws,
            XLHelper.TrimRowNumber(shiftedRange.RangeAddress.FirstAddress.RowNumber),
            matchRange.RangeAddress.FirstAddress.ColumnLetter,
            matchRange.RangeAddress.FirstAddress.FixedRow,
            matchRange.RangeAddress.FirstAddress.FixedColumn));
        sb.Append(':');
        sb.Append(new XLAddress(ws,
            XLHelper.TrimRowNumber(matchRange.RangeAddress.LastAddress.RowNumber + rowsShifted),
            matchRange.RangeAddress.LastAddress.ColumnLetter,
            matchRange.RangeAddress.LastAddress.FixedRow,
            matchRange.RangeAddress.LastAddress.FixedColumn));
    }

    private static bool IsRowRangeWithinShiftedRange(XLRange shiftedRange, IXLRange matchRange)
    {
        return shiftedRange.RangeAddress.FirstAddress.RowNumber <= matchRange.RangeAddress.LastAddress.RowNumber
            && shiftedRange.RangeAddress.FirstAddress.ColumnNumber <= matchRange.RangeAddress.FirstAddress.ColumnNumber
            && shiftedRange.RangeAddress.LastAddress.ColumnNumber >= matchRange.RangeAddress.LastAddress.ColumnNumber;
    }

    private static string ShiftRowString(string rowString, int rowsShifted)
    {
        if (rowString[0] == '$')
            return "$" + XLHelper.TrimRowNumber(int.Parse(rowString.Substring(1)) + rowsShifted).ToInvariantString();

        return XLHelper.TrimRowNumber(int.Parse(rowString) + rowsShifted).ToInvariantString();
    }

    private static void AppendShiftedRowOnlyRange(StringBuilder sb, string rangeAddress, int rowsShifted)
    {
        var rows = rangeAddress.Split(':');
        sb.Append(ShiftRowString(rows[0], rowsShifted));
        sb.Append(':');
        sb.Append(ShiftRowString(rows[1], rowsShifted));
    }

    private static void AppendShiftedRowCellRange(StringBuilder sb, XLWorksheet ws, IXLRange matchRange,
        string rangeAddress, int rowsShifted)
    {
        sb.Append(new XLAddress(ws,
            XLHelper.TrimRowNumber(matchRange.RangeAddress.FirstAddress.RowNumber + rowsShifted),
            matchRange.RangeAddress.FirstAddress.ColumnLetter,
            matchRange.RangeAddress.FirstAddress.FixedRow,
            matchRange.RangeAddress.FirstAddress.FixedColumn));

        if (rangeAddress.Contains(':'))
        {
            sb.Append(':');
            sb.Append(new XLAddress(ws,
                XLHelper.TrimRowNumber(matchRange.RangeAddress.LastAddress.RowNumber + rowsShifted),
                matchRange.RangeAddress.LastAddress.ColumnLetter,
                matchRange.RangeAddress.LastAddress.FixedRow,
                matchRange.RangeAddress.LastAddress.FixedColumn));
        }
    }

    private static void AppendPartialRowShift(StringBuilder sb, XLWorksheet ws, IXLRange matchRange, int rowsShifted)
    {
        sb.Append(matchRange.RangeAddress.FirstAddress);
        sb.Append(':');
        sb.Append(new XLAddress(ws,
            XLHelper.TrimRowNumber(matchRange.RangeAddress.LastAddress.RowNumber + rowsShifted),
            matchRange.RangeAddress.LastAddress.ColumnLetter,
            matchRange.RangeAddress.LastAddress.FixedRow,
            matchRange.RangeAddress.LastAddress.FixedColumn));
    }

    internal static string ShiftFormulaColumns(string formulaA1, XLWorksheet worksheetInAction, XLRange shiftedRange,
        int columnsShifted)
    {
        if (string.IsNullOrWhiteSpace(formulaA1)) return string.Empty;

        var value = formulaA1;
        var sb = new StringBuilder();
        var lastIndex = 0;

        foreach (Match match in A1SimpleRegex.Matches(value))
        {
            var matchString = match.Value;
            var matchIndex = match.Index;
            if (value.AsSpan(0, matchIndex).Count('"') % 2 == 0)
            {
                sb.Append(value.AsSpan(lastIndex, matchIndex - lastIndex));
                var (sheetName, useSheetName) = ExtractSheetName(matchString, worksheetInAction);

                if (String.Compare(sheetName, shiftedRange.Worksheet.Name, StringComparison.OrdinalIgnoreCase) == 0)
                    AppendShiftedColumnMatch(sb, matchString, sheetName, useSheetName, worksheetInAction, shiftedRange, columnsShifted);
                else
                    sb.Append(matchString);
            }
            else
                sb.Append(value.AsSpan(lastIndex, matchIndex - lastIndex + matchString.Length));

            lastIndex = matchIndex + matchString.Length;
        }

        if (lastIndex < value.Length)
            sb.Append(value.AsSpan(lastIndex));

        return sb.ToString();
    }

    private static void AppendShiftedColumnMatch(StringBuilder sb, string matchString, string sheetName, bool useSheetName,
        XLWorksheet worksheetInAction, XLRange shiftedRange, int columnsShifted)
    {
        var rangeAddress = matchString[(matchString.IndexOf('!') + 1)..];
        if (A1RowRegex.IsMatch(rangeAddress))
        {
            sb.Append(matchString);
            return;
        }

        var matchRange = worksheetInAction.Workbook.Worksheet(sheetName).Range(rangeAddress);
        if (!IsColumnRangeWithinShiftedRange(shiftedRange, matchRange))
        {
            sb.Append(matchString);
            return;
        }

        if (useSheetName)
        {
            sb.Append(sheetName.EscapeSheetName());
            sb.Append('!');
        }

        if (A1ColumnRegex.IsMatch(rangeAddress))
            AppendShiftedColumnOnlyRange(sb, rangeAddress, columnsShifted);
        else if (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <= matchRange.RangeAddress.FirstAddress.ColumnNumber)
        {
            if (IsLeftBoundaryDeletion(shiftedRange, matchRange, columnsShifted))
                AppendClampedLeftColumnShift(sb, worksheetInAction, shiftedRange, matchRange, columnsShifted);
            else
                AppendShiftedColumnCellRange(sb, worksheetInAction, matchRange, rangeAddress, columnsShifted);
        }
        else
            AppendPartialColumnShift(sb, worksheetInAction, matchRange, columnsShifted);
    }

    /// <summary>
    /// Column-wise counterpart of <see cref="IsTopBoundaryDeletion"/>: true when a column deletion removes
    /// the left boundary of <paramref name="matchRange"/> while some columns to the right survive. Excel
    /// keeps the range's left column fixed at the deletion start and shifts only the right edge left, e.g.
    /// deleting column C turns C1:D1 into C1:C1 rather than expanding it to B1:C1. See issue #2866.
    /// </summary>
    private static bool IsLeftBoundaryDeletion(XLRange shiftedRange, IXLRange matchRange, int columnsShifted)
    {
        return columnsShifted < 0
            && matchRange.RangeAddress.FirstAddress.ColumnNumber <= shiftedRange.RangeAddress.LastAddress.ColumnNumber
            && matchRange.RangeAddress.LastAddress.ColumnNumber > shiftedRange.RangeAddress.LastAddress.ColumnNumber;
    }

    private static void AppendClampedLeftColumnShift(StringBuilder sb, XLWorksheet ws, XLRange shiftedRange,
        IXLRange matchRange, int columnsShifted)
    {
        sb.Append(new XLAddress(ws,
            matchRange.RangeAddress.FirstAddress.RowNumber,
            XLHelper.TrimColumnNumber(shiftedRange.RangeAddress.FirstAddress.ColumnNumber),
            matchRange.RangeAddress.FirstAddress.FixedRow,
            matchRange.RangeAddress.FirstAddress.FixedColumn));
        sb.Append(':');
        sb.Append(new XLAddress(ws,
            matchRange.RangeAddress.LastAddress.RowNumber,
            XLHelper.TrimColumnNumber(matchRange.RangeAddress.LastAddress.ColumnNumber + columnsShifted),
            matchRange.RangeAddress.LastAddress.FixedRow,
            matchRange.RangeAddress.LastAddress.FixedColumn));
    }

    private static bool IsColumnRangeWithinShiftedRange(XLRange shiftedRange, IXLRange matchRange)
    {
        return shiftedRange.RangeAddress.FirstAddress.ColumnNumber <= matchRange.RangeAddress.LastAddress.ColumnNumber
            && shiftedRange.RangeAddress.FirstAddress.RowNumber <= matchRange.RangeAddress.FirstAddress.RowNumber
            && shiftedRange.RangeAddress.LastAddress.RowNumber >= matchRange.RangeAddress.LastAddress.RowNumber;
    }

    private static string ShiftColumnString(string columnString, int columnsShifted)
    {
        if (columnString[0] == '$')
            return "$" + XLHelper.GetColumnLetterFromNumber(
                XLHelper.GetColumnNumberFromLetter(columnString.Substring(1)) + columnsShifted, true);

        return XLHelper.GetColumnLetterFromNumber(
            XLHelper.GetColumnNumberFromLetter(columnString) + columnsShifted, true);
    }

    private static void AppendShiftedColumnOnlyRange(StringBuilder sb, string rangeAddress, int columnsShifted)
    {
        var columns = rangeAddress.Split(':');
        sb.Append(ShiftColumnString(columns[0], columnsShifted));
        sb.Append(':');
        sb.Append(ShiftColumnString(columns[1], columnsShifted));
    }

    private static void AppendShiftedColumnCellRange(StringBuilder sb, XLWorksheet ws, IXLRange matchRange,
        string rangeAddress, int columnsShifted)
    {
        sb.Append(new XLAddress(ws,
            matchRange.RangeAddress.FirstAddress.RowNumber,
            XLHelper.TrimColumnNumber(matchRange.RangeAddress.FirstAddress.ColumnNumber + columnsShifted),
            matchRange.RangeAddress.FirstAddress.FixedRow,
            matchRange.RangeAddress.FirstAddress.FixedColumn));

        if (rangeAddress.Contains(':'))
        {
            sb.Append(':');
            sb.Append(new XLAddress(ws,
                matchRange.RangeAddress.LastAddress.RowNumber,
                XLHelper.TrimColumnNumber(matchRange.RangeAddress.LastAddress.ColumnNumber + columnsShifted),
                matchRange.RangeAddress.LastAddress.FixedRow,
                matchRange.RangeAddress.LastAddress.FixedColumn));
        }
    }

    private static void AppendPartialColumnShift(StringBuilder sb, XLWorksheet ws, IXLRange matchRange, int columnsShifted)
    {
        sb.Append(matchRange.RangeAddress.FirstAddress);
        sb.Append(':');
        sb.Append(new XLAddress(ws,
            matchRange.RangeAddress.LastAddress.RowNumber,
            XLHelper.TrimColumnNumber(matchRange.RangeAddress.LastAddress.ColumnNumber + columnsShifted),
            matchRange.RangeAddress.LastAddress.FixedRow,
            matchRange.RangeAddress.LastAddress.FixedColumn));
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
