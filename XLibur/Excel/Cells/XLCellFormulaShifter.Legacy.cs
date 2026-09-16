using System;
using System.Text;
using System.Text.RegularExpressions;
using XLibur.Excel.Coordinates;
using XLibur.Extensions;

namespace XLibur.Excel;

/// <summary>
/// The original regex-based reference shifter, kept as the fallback for formulas
/// <see cref="XLCellFormulaShifter"/>'s parser path cannot parse (external workbook references such
/// as <c>'[file.xlsx]Sheet'!A1</c> are the known case). It is not the primary path and is not where
/// behaviour is defined — see <c>XLCellFormulaShifter.cs</c> for that, including the boundary cases
/// this implementation gets wrong.
/// </summary>
internal static partial class XLCellFormulaShifter
{
    private const string RefError = "#REF!";

    private static readonly Regex A1SimpleRegex = A1SimpleRegexGenerated();

    private static readonly Regex A1RowRegex = A1RowRegexGenerated();

    private static readonly Regex A1ColumnRegex = A1ColumnRegexGenerated();

    internal static string ShiftFormulaRowsLegacy(string formulaA1, XLWorksheet worksheetInAction, XLRange shiftedRange,
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

                if (sheetName is not null && string.Equals(sheetName, shiftedWsName, StringComparison.OrdinalIgnoreCase))
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

    /// <summary>
    /// The sheet a matched reference names: the name written before its <c>!</c>, unquoted and
    /// unescaped, or the sheet the formula lives on when the reference names none.
    /// </summary>
    /// <remarks>
    /// A name opened with an apostrophe has to close with one, and this used to assume it did. The
    /// reference regex's unquoted alternative admits a name written with one apostrophe instead of two
    /// (see <see cref="A1SimpleRegexGenerated"/>), and unescaping <c>'Data</c> as though it were
    /// quoted took <c>Substring(1, 3)</c> of five characters and read the name as <c>Dat</c> — so a
    /// workbook that really had a sheet named <c>Dat</c> got the malformed reference shifted as one of
    /// its references (#576). Such a name is reported as no sheet at all, which is the outcome the
    /// truncation was hiding: an unrecognised sheet name leaves its reference verbatim.
    /// <para>
    /// The separator between the sheet part and the range is found by
    /// <see cref="FindSheetSeparatorIndex"/>, which looks past a quoted name's closing apostrophe
    /// (accounting for doubled apostrophes inside it) rather than stopping at the first <c>!</c> in the
    /// match. A quoted name may itself contain <c>!</c>, a legal sheet-name character
    /// (<see cref="XLHelper.TryValidateSheetName"/> excludes <c>: \ / ? * [ ]</c> and not this one), so
    /// <c>'Q1!Sales'!A5</c> now reads the sheet as <c>Q1!Sales</c> instead of splitting inside the name
    /// (#584). <see cref="AppendShiftedRowMatch"/> and <see cref="AppendShiftedColumnMatch"/> use the
    /// same helper for their range extraction, so all three sites agree on where a reference splits.
    /// </para>
    /// </remarks>
    /// <returns>
    /// The sheet name, or <c>null</c> when no sheet name can be read out of the match, and whether
    /// the reference wrote a sheet name at all — a rewritten reference has to write it back.
    /// </returns>
    private static (string? sheetName, bool useSheetName) ExtractSheetName(string matchString, XLWorksheet worksheetInAction)
    {
        var separatorIndex = FindSheetSeparatorIndex(matchString);
        if (separatorIndex < 0)
            return (worksheetInAction.Name, false);

        var sheetName = matchString.Substring(0, separatorIndex);
        if (!sheetName.StartsWith('\''))
            return (sheetName, true);

        // A quoted name is an opening apostrophe, at least one character of name and a closing one, so
        // fewer than three characters is malformed however they are arranged: ' has no closing
        // apostrophe and '' has no name between the two it has.
        if (sheetName.Length < 3 || !sheetName.EndsWith('\''))
            return (null, true);

        return (sheetName.Substring(1, sheetName.Length - 2).Replace("''", "'"), true);
    }

    /// <summary>
    /// The index of the <c>!</c> that separates a matched reference's sheet part from its range part —
    /// the shared helper <see cref="ExtractSheetName"/>, <see cref="AppendShiftedRowMatch"/> and
    /// <see cref="AppendShiftedColumnMatch"/> all split on, so the three sites cannot disagree about
    /// where a reference splits.
    /// </summary>
    /// <remarks>
    /// An unquoted name cannot contain <c>!</c> or an apostrophe, so the first <c>!</c> in the match is
    /// still the right separator for one, and also for a match with no sheet name at all. A quoted name
    /// can contain <c>!</c>, so for one the separator is the <c>!</c> immediately after its closing
    /// apostrophe, skipping over any doubled apostrophes (<c>''</c>, an escaped apostrophe) inside the
    /// name. A match that opens with an apostrophe but never closes it — the malformed, one-apostrophe
    /// form <see cref="ExtractSheetName"/>'s remarks describe — falls back to the first <c>!</c> too,
    /// which is what lets that method keep reporting such a name as unrecognised rather than reading a
    /// range address out of it.
    /// </remarks>
    private static int FindSheetSeparatorIndex(string matchString)
    {
        if (matchString.Length == 0 || matchString[0] != '\'')
            return matchString.IndexOf('!');

        var i = 1;
        while (i < matchString.Length)
        {
            if (matchString[i] != '\'')
            {
                i++;
                continue;
            }

            if (i + 1 < matchString.Length && matchString[i + 1] == '\'')
            {
                i += 2;
                continue;
            }

            // matchString[i] is the closing apostrophe of the quoted name; the separator is the '!' that
            // must follow it directly. If it does not, the name never really closed and the caller falls
            // back to the first '!' in the match, same as an unquoted, unterminated name.
            return i + 1 < matchString.Length && matchString[i + 1] == '!' ? i + 1 : matchString.IndexOf('!');
        }

        return matchString.IndexOf('!');
    }

    /// <summary>
    /// The sheet a matched reference names, which is always the sheet being shifted: both callers reach
    /// their append only once the reference's sheet name has matched that sheet's name.
    /// </summary>
    /// <remarks>
    /// Taking the sheet from the shifted range rather than looking the name up again is what makes a
    /// name holding a doubled apostrophe work. <see cref="XLWorkbook.Worksheet(string)"/> undoubles the
    /// apostrophes of the name it is given, and <see cref="ExtractSheetName"/> has already done so, so
    /// a sheet legally named <c>Ann''s</c> was looked for as <c>Ann's</c> — a sheet that does not
    /// exist. Nothing reached that until the sheet-name pattern was fixed, because such a reference was
    /// never recognised as naming the shifted sheet in the first place (#570).
    /// </remarks>
    /// <returns>
    /// The sheet as <see cref="IXLWorksheet"/>, so that <c>Range(string)</c> keeps raising its own
    /// exception for an address it cannot read rather than handing back null.
    /// </returns>
    private static IXLWorksheet ReferencedSheet(XLRange shiftedRange) => shiftedRange.Worksheet;

    private static void AppendShiftedRowMatch(StringBuilder sb, string matchString, string sheetName, bool useSheetName,
        XLWorksheet worksheetInAction, XLRange shiftedRange, int rowsShifted)
    {
        var rangeAddress = matchString.Substring(FindSheetSeparatorIndex(matchString) + 1);
        if (A1ColumnRegex.IsMatch(rangeAddress))
        {
            sb.Append(matchString);
            return;
        }

        var matchRange = ReferencedSheet(shiftedRange).Range(rangeAddress);
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

        if (IsDeletedEntirelyByRowShift(shiftedRange, matchRange, rowsShifted))
            sb.Append(RefError);
        else if (A1RowRegex.IsMatch(rangeAddress))
            AppendShiftedRowOnlyRange(sb, rangeAddress, shiftedRange, matchRange, rowsShifted);
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
    /// True when a row deletion removes every row of <paramref name="matchRange"/>. Excel replaces such
    /// a reference with <c>#REF!</c> rather than repointing it, so deleting rows 1-5 turns <c>A1:B2</c>
    /// into <c>#REF!</c>. Without this the shifted endpoints would be negative and
    /// <see cref="XLHelper.TrimRowNumber"/> would clamp them back to row 1, leaving the reference
    /// pointing at a surviving row it never covered (<c>A1:B1</c>). See ClosedXML/ClosedXML#880.
    /// </summary>
    private static bool IsDeletedEntirelyByRowShift(XLRange shiftedRange, IXLRange matchRange, int rowsShifted)
    {
        return rowsShifted < 0
            && matchRange.RangeAddress.FirstAddress.RowNumber >= shiftedRange.RangeAddress.FirstAddress.RowNumber
            && matchRange.RangeAddress.LastAddress.RowNumber + rowsShifted <
               shiftedRange.RangeAddress.FirstAddress.RowNumber;
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

    /// <summary>
    /// A row-only reference (<c>3:5</c>) obeys the same rules as a cell range but renders as bare row
    /// numbers. Both boundaries are decided here rather than shifted blindly: the top only moves when the
    /// shift starts at or above it, and a deletion that removes the top boundary leaves the top pinned at
    /// the deletion start; the bottom always shifts.
    /// <para>
    /// Shifting both endpoints regardless (as this did) walked the reference onto rows it never covered.
    /// Deleting row 4 turned 3:5 into 2:4, and it kept the reference from ever shrinking to nothing, so
    /// <see cref="IsDeletedEntirelyByRowShift"/> never saw the fully deleted case. The same applied to
    /// insertions: inserting two rows at row 4 turned 3:5 into 5:7 instead of expanding it to 3:7.
    /// </para>
    /// </summary>
    private static void AppendShiftedRowOnlyRange(StringBuilder sb, string rangeAddress, XLRange shiftedRange,
        IXLRange matchRange, int rowsShifted)
    {
        var firstRow = matchRange.RangeAddress.FirstAddress.RowNumber;
        var lastRow = matchRange.RangeAddress.LastAddress.RowNumber;
        var shiftStart = shiftedRange.RangeAddress.FirstAddress.RowNumber;

        var newFirstRow = firstRow;
        if (shiftStart <= firstRow)
        {
            newFirstRow = IsTopBoundaryDeletion(shiftedRange, matchRange, rowsShifted)
                ? shiftStart
                : firstRow + rowsShifted;
        }

        var rows = rangeAddress.Split(':');
        AppendRowBoundary(sb, rows[0], newFirstRow);
        sb.Append(':');
        AppendRowBoundary(sb, rows[1], lastRow + rowsShifted);
    }

    private static void AppendRowBoundary(StringBuilder sb, string rowToken, int rowNumber)
    {
        if (rowToken[0] == '$')
            sb.Append('$');

        sb.Append(XLHelper.TrimRowNumber(rowNumber).ToInvariantString());
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

    internal static string ShiftFormulaColumnsLegacy(string formulaA1, XLWorksheet worksheetInAction, XLRange shiftedRange,
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

                if (sheetName is not null && string.Equals(sheetName, shiftedRange.Worksheet.Name, StringComparison.OrdinalIgnoreCase))
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
        var rangeAddress = matchString[(FindSheetSeparatorIndex(matchString) + 1)..];
        if (A1RowRegex.IsMatch(rangeAddress))
        {
            sb.Append(matchString);
            return;
        }

        var matchRange = ReferencedSheet(shiftedRange).Range(rangeAddress);
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

        if (IsDeletedEntirelyByColumnShift(shiftedRange, matchRange, columnsShifted))
            sb.Append(RefError);
        else if (A1ColumnRegex.IsMatch(rangeAddress))
            AppendShiftedColumnOnlyRange(sb, rangeAddress, shiftedRange, matchRange, columnsShifted);
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
    /// Column-wise counterpart of <see cref="IsDeletedEntirelyByRowShift"/>: true when a column deletion
    /// removes every column of <paramref name="matchRange"/>, so the reference becomes <c>#REF!</c> instead
    /// of being clamped back to column A by <see cref="XLHelper.TrimColumnNumber"/>.
    /// </summary>
    private static bool IsDeletedEntirelyByColumnShift(XLRange shiftedRange, IXLRange matchRange, int columnsShifted)
    {
        return columnsShifted < 0
            && matchRange.RangeAddress.FirstAddress.ColumnNumber >=
               shiftedRange.RangeAddress.FirstAddress.ColumnNumber
            && matchRange.RangeAddress.LastAddress.ColumnNumber + columnsShifted <
               shiftedRange.RangeAddress.FirstAddress.ColumnNumber;
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

    /// <summary>
    /// Column-wise counterpart of <see cref="AppendShiftedRowOnlyRange"/>: a column-only reference
    /// (<c>C:E</c>) keeps its left boundary unless the shift starts at or to the left of it, and pins that
    /// boundary at the deletion start when a deletion removes it.
    /// </summary>
    private static void AppendShiftedColumnOnlyRange(StringBuilder sb, string rangeAddress, XLRange shiftedRange,
        IXLRange matchRange, int columnsShifted)
    {
        var firstColumn = matchRange.RangeAddress.FirstAddress.ColumnNumber;
        var lastColumn = matchRange.RangeAddress.LastAddress.ColumnNumber;
        var shiftStart = shiftedRange.RangeAddress.FirstAddress.ColumnNumber;

        var newFirstColumn = firstColumn;
        if (shiftStart <= firstColumn)
        {
            newFirstColumn = IsLeftBoundaryDeletion(shiftedRange, matchRange, columnsShifted)
                ? shiftStart
                : firstColumn + columnsShifted;
        }

        var columns = rangeAddress.Split(':');
        AppendColumnBoundary(sb, columns[0], newFirstColumn);
        sb.Append(':');
        AppendColumnBoundary(sb, columns[1], lastColumn + columnsShifted);
    }

    private static void AppendColumnBoundary(StringBuilder sb, string columnToken, int columnNumber)
    {
        if (columnToken[0] == '$')
            sb.Append('$');

        sb.Append(XLHelper.GetColumnLetterFromNumber(columnNumber, true));
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

    /// <summary>
    /// The reference finder. Unlike <see cref="XLHelper.A1SimpleRegex"/>, which is anchored and asks
    /// whether a whole string is one range address, this one scans a formula for every reference in it,
    /// so it is unanchored and guards the bare A1 form with word boundaries instead. The two cannot be
    /// one regex for that reason; the quoted sheet-name alternative they do share is
    /// <see cref="XLHelper.QuotedSheetNamePattern"/>, defined once (#570, and #524 before it).
    /// </summary>
    /// <remarks>
    /// The unquoted alternative keeps its optional apostrophes, where
    /// <see cref="XLHelper.A1SimpleRegex"/> dropped them for #560. Here they are what consumes the
    /// trailing apostrophe of an external workbook reference: <c>'[file.xlsx]Sheet'!A1</c> matches as
    /// <c>Sheet'!A1</c>, whose sheet name matches no real sheet, so the reference is left alone.
    /// Narrowing this to <c>\w+</c> would leave a bare <c>A1</c> to match on its own, which the shifter
    /// would then read as a reference to the sheet the formula lives on and move.
    /// <para>
    /// Keeping them does mean this alternative still admits a name with one apostrophe instead of two,
    /// as in <c>'Data!A1</c>. That is the price of the branch above and not an endorsement of the form:
    /// matching such a name is not the same as reading it, and <see cref="ExtractSheetName"/> reports
    /// no sheet for it rather than unescaping it into a shorter, real one (#576).
    /// </para>
    /// </remarks>
    [GeneratedRegex(
        @"(?<Reference>(?<Sheet>(" + XLHelper.QuotedSheetNamePattern + @"|\'?\w+\'?)!)?(?<Range>(?<![\w\d])\$?[a-zA-Z]{1,3}\$?\d{1,7}(?<RangeEnd>:\$?[a-zA-Z]{1,3}\$?\d{1,7})?(?![\w\d])|(?<ColumnNumbers>\$?\d{1,7}:\$?\d{1,7})|(?<ColumnLetters>\$?[a-zA-Z]{1,3}:\$?[a-zA-Z]{1,3})))",
        RegexOptions.Compiled)]
    private static partial Regex A1SimpleRegexGenerated();
}
