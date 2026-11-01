using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using XLibur.Excel;

namespace XLibur.Report.Tests.Infrastructure;

/// <summary>
/// Compares two workbooks by meaning rather than by bytes, reporting every difference it finds.
/// </summary>
/// <remarks>
/// A generated .xlsx is not byte-reproducible — zip ordering, timestamps and part ids all vary —
/// so golden-file testing has to compare the model instead. Reporting all differences rather than
/// failing on the first keeps a broken fixture to one debugging round-trip.
/// </remarks>
public static class WorkbookComparer
{
    /// <summary>
    /// Compares <paramref name="actual"/> against <paramref name="expected"/> and returns a
    /// human-readable description of each difference. An empty list means they match.
    /// </summary>
    public static IReadOnlyList<string> Compare(
        IXLWorkbook expected,
        IXLWorkbook actual,
        WorkbookComparisonOptions? options = null)
    {
        options ??= WorkbookComparisonOptions.Default;
        var differences = new List<string>();

        var expectedSheets = expected.Worksheets.Select(w => w.Name).ToList();
        var actualSheets = actual.Worksheets.Select(w => w.Name).ToList();

        if (!expectedSheets.SequenceEqual(actualSheets, StringComparer.Ordinal))
        {
            differences.Add(
                $"Worksheets: expected [{string.Join(", ", expectedSheets)}] but was [{string.Join(", ", actualSheets)}]");
            return differences;
        }

        foreach (var name in expectedSheets)
        {
            CompareWorksheet(expected.Worksheet(name), actual.Worksheet(name), options, differences);

            if (differences.Count >= options.MaxDifferences)
            {
                differences.Add($"... stopped after {options.MaxDifferences} differences.");
                break;
            }
        }

        return differences;
    }

    private static void CompareWorksheet(
        IXLWorksheet expected,
        IXLWorksheet actual,
        WorkbookComparisonOptions options,
        List<string> differences)
    {
        var sheet = expected.Name;

        CompareCells(expected, actual, options, differences, sheet);

        if (options.MergedRanges)
        {
            CompareSets(
                expected.MergedRanges.Select(r => r.RangeAddress.ToStringRelative()),
                actual.MergedRanges.Select(r => r.RangeAddress.ToStringRelative()),
                $"{sheet}: merged ranges",
                differences);
        }

        if (options.ConditionalFormats)
        {
            CompareConditionalFormats(expected, actual, differences, sheet);
        }

        if (options.Dimensions)
        {
            CompareDimensions(expected, actual, differences, sheet);
        }

        if (options.PageSetup)
        {
            ComparePageSetup(expected, actual, differences, sheet);
        }
    }

    private static void CompareCells(
        IXLWorksheet expected,
        IXLWorksheet actual,
        WorkbookComparisonOptions options,
        List<string> differences,
        string sheet)
    {
        var (lastRow, lastColumn) = UsedExtent(expected, actual);

        for (var row = 1; row <= lastRow; row++)
        {
            for (var column = 1; column <= lastColumn; column++)
            {
                if (differences.Count >= options.MaxDifferences)
                {
                    return;
                }

                CompareCell(expected.Cell(row, column), actual.Cell(row, column), options, differences, sheet);
            }
        }
    }

    private static void CompareCell(
        IXLCell expectedCell,
        IXLCell actualCell,
        WorkbookComparisonOptions options,
        List<string> differences,
        string sheet)
    {
        var location = $"{sheet}!{expectedCell.Address.ToStringRelative()}";

        if (options.Values)
        {
            CompareValue(expectedCell, actualCell, differences, location);
        }

        if (options.Formulas)
        {
            CompareFormula(expectedCell, actualCell, differences, location);
        }

        if (options.Styles)
        {
            CompareStyle(expectedCell, actualCell, differences, location);
        }

        if (options.Comments)
        {
            CompareComment(expectedCell, actualCell, differences, location);
        }

        if (options.Hyperlinks)
        {
            CompareHyperlink(expectedCell, actualCell, differences, location);
        }
    }

    private static void CompareValue(IXLCell expectedCell, IXLCell actualCell, List<string> differences, string location)
    {
        var expectedValue = Describe(expectedCell.Value);
        var actualValue = Describe(actualCell.Value);
        if (!string.Equals(expectedValue, actualValue, StringComparison.Ordinal))
        {
            differences.Add($"{location}: value expected {expectedValue} but was {actualValue}");
        }
    }

    private static void CompareFormula(IXLCell expectedCell, IXLCell actualCell, List<string> differences, string location)
    {
        var expectedFormula = DescribeFormula(expectedCell);
        var actualFormula = DescribeFormula(actualCell);
        if (!string.Equals(expectedFormula, actualFormula, StringComparison.Ordinal))
        {
            differences.Add($"{location}: formula expected '{expectedFormula}' but was '{actualFormula}'");
        }
    }

    private static void CompareStyle(IXLCell expectedCell, IXLCell actualCell, List<string> differences, string location)
    {
        var expectedStyle = expectedCell.Style.ToString();
        var actualStyle = actualCell.Style.ToString();
        if (!string.Equals(expectedStyle, actualStyle, StringComparison.Ordinal))
        {
            differences.Add($"{location}: style differs\n  expected: {expectedStyle}\n  actual:   {actualStyle}");
        }
    }

    private static void CompareComment(IXLCell expectedCell, IXLCell actualCell, List<string> differences, string location)
    {
        var expectedComment = DescribeComment(expectedCell);
        var actualComment = DescribeComment(actualCell);
        if (!string.Equals(expectedComment, actualComment, StringComparison.Ordinal))
        {
            differences.Add($"{location}: comment expected '{expectedComment}' but was '{actualComment}'");
        }
    }

    private static void CompareHyperlink(IXLCell expectedCell, IXLCell actualCell, List<string> differences, string location)
    {
        var expectedLink = DescribeHyperlink(expectedCell);
        var actualLink = DescribeHyperlink(actualCell);
        if (!string.Equals(expectedLink, actualLink, StringComparison.Ordinal))
        {
            differences.Add($"{location}: hyperlink expected '{expectedLink}' but was '{actualLink}'");
        }
    }

    private static string DescribeFormula(IXLCell cell) => cell.HasFormula ? cell.FormulaA1 : string.Empty;

    private static string DescribeComment(IXLCell cell) => cell.HasComment ? cell.GetComment().Text : string.Empty;

    private static void CompareConditionalFormats(
        IXLWorksheet expected,
        IXLWorksheet actual,
        List<string> differences,
        string sheet)
    {
        var expectedFormats = expected.ConditionalFormats.Select(DescribeConditionalFormat).ToList();
        var actualFormats = actual.ConditionalFormats.Select(DescribeConditionalFormat).ToList();

        if (expectedFormats.Count != actualFormats.Count)
        {
            differences.Add(
                $"{sheet}: conditional format count expected {expectedFormats.Count} but was {actualFormats.Count}");
        }

        CompareSets(expectedFormats, actualFormats, $"{sheet}: conditional formats", differences);
    }

    private static void CompareDimensions(
        IXLWorksheet expected,
        IXLWorksheet actual,
        List<string> differences,
        string sheet)
    {
        var (lastRow, lastColumn) = UsedExtent(expected, actual);

        for (var row = 1; row <= lastRow; row++)
        {
            CompareRow(expected.Row(row), actual.Row(row), row, differences, sheet);
        }

        for (var column = 1; column <= lastColumn; column++)
        {
            CompareColumn(expected.Column(column), actual.Column(column), column, differences, sheet);
        }
    }

    private static void CompareRow(
        IXLRow expectedRow,
        IXLRow actualRow,
        int row,
        List<string> differences,
        string sheet)
    {
        if (Math.Abs(expectedRow.Height - actualRow.Height) > 0.001)
        {
            differences.Add(
                $"{sheet}: row {row} height expected {Format(expectedRow.Height)} but was {Format(actualRow.Height)}");
        }

        if (expectedRow.OutlineLevel != actualRow.OutlineLevel)
        {
            differences.Add(
                $"{sheet}: row {row} outline level expected {expectedRow.OutlineLevel} but was {actualRow.OutlineLevel}");
        }

        if (expectedRow.IsHidden != actualRow.IsHidden)
        {
            differences.Add($"{sheet}: row {row} hidden expected {expectedRow.IsHidden} but was {actualRow.IsHidden}");
        }
    }

    private static void CompareColumn(
        IXLColumn expectedColumn,
        IXLColumn actualColumn,
        int column,
        List<string> differences,
        string sheet)
    {
        if (Math.Abs(expectedColumn.Width - actualColumn.Width) > 0.001)
        {
            differences.Add(
                $"{sheet}: column {column} width expected {Format(expectedColumn.Width)} but was {Format(actualColumn.Width)}");
        }

        if (expectedColumn.OutlineLevel != actualColumn.OutlineLevel)
        {
            differences.Add(
                $"{sheet}: column {column} outline level expected {expectedColumn.OutlineLevel} but was {actualColumn.OutlineLevel}");
        }
    }

    private static void ComparePageSetup(
        IXLWorksheet expected,
        IXLWorksheet actual,
        List<string> differences,
        string sheet)
    {
        CompareSets(
            expected.PageSetup.RowBreaks.Select(b => b.ToString(CultureInfo.InvariantCulture)),
            actual.PageSetup.RowBreaks.Select(b => b.ToString(CultureInfo.InvariantCulture)),
            $"{sheet}: row page breaks",
            differences);

        CompareSets(
            expected.PageSetup.ColumnBreaks.Select(b => b.ToString(CultureInfo.InvariantCulture)),
            actual.PageSetup.ColumnBreaks.Select(b => b.ToString(CultureInfo.InvariantCulture)),
            $"{sheet}: column page breaks",
            differences);

        if (expected.AutoFilter.IsEnabled != actual.AutoFilter.IsEnabled)
        {
            differences.Add(
                $"{sheet}: autofilter expected {expected.AutoFilter.IsEnabled} but was {actual.AutoFilter.IsEnabled}");
        }
    }

    private static void CompareSets(
        IEnumerable<string> expected,
        IEnumerable<string> actual,
        string label,
        List<string> differences)
    {
        var expectedSet = expected.OrderBy(x => x, StringComparer.Ordinal).ToList();
        var actualSet = actual.OrderBy(x => x, StringComparer.Ordinal).ToList();

        if (expectedSet.SequenceEqual(actualSet, StringComparer.Ordinal))
        {
            return;
        }

        var missing = expectedSet.Except(actualSet, StringComparer.Ordinal).ToList();
        var unexpected = actualSet.Except(expectedSet, StringComparer.Ordinal).ToList();

        if (missing.Count > 0)
        {
            differences.Add($"{label}: missing [{string.Join(", ", missing)}]");
        }

        if (unexpected.Count > 0)
        {
            differences.Add($"{label}: unexpected [{string.Join(", ", unexpected)}]");
        }
    }

    private static (int LastRow, int LastColumn) UsedExtent(IXLWorksheet expected, IXLWorksheet actual)
    {
        // Deliberately XLCellsUsedOptions.All rather than contents alone: a report can differ in
        // a cell that holds no value — banded fill on a generated row, a hyperlink, a merge — and
        // a contents-only extent would place those outside the compared region entirely.
        var expectedRange = expected.RangeUsed(XLCellsUsedOptions.All);
        var actualRange = actual.RangeUsed(XLCellsUsedOptions.All);

        var lastRow = Math.Max(
            expectedRange?.RangeAddress.LastAddress.RowNumber ?? 0,
            actualRange?.RangeAddress.LastAddress.RowNumber ?? 0);
        var lastColumn = Math.Max(
            expectedRange?.RangeAddress.LastAddress.ColumnNumber ?? 0,
            actualRange?.RangeAddress.LastAddress.ColumnNumber ?? 0);

        return (lastRow, lastColumn);
    }

    private static string DescribeConditionalFormat(IXLConditionalFormat format)
    {
        var ranges = string.Join(",", format.Ranges.Select(r => r.RangeAddress.ToStringRelative()).OrderBy(r => r, StringComparer.Ordinal));
        return $"{format.ConditionalFormatType}[{ranges}]";
    }

    private static string DescribeHyperlink(IXLCell cell)
    {
        if (!cell.HasHyperlink)
        {
            return string.Empty;
        }

        var hyperlink = cell.GetHyperlink();
        return hyperlink.IsExternal
            ? hyperlink.ExternalAddress?.ToString() ?? string.Empty
            : hyperlink.InternalAddress ?? string.Empty;
    }

    private static string Describe(XLCellValue value) => value.Type switch
    {
        XLDataType.Blank => "Blank",
        XLDataType.Boolean => $"Boolean:{value.GetBoolean()}",
        XLDataType.Number => $"Number:{value.GetNumber().ToString("R", CultureInfo.InvariantCulture)}",
        XLDataType.Text => $"Text:{value.GetText()}",
        XLDataType.Error => $"Error:{value.GetError()}",
        XLDataType.DateTime => $"DateTime:{value.GetDateTime().ToString("O", CultureInfo.InvariantCulture)}",
        XLDataType.TimeSpan => $"TimeSpan:{value.GetTimeSpan()}",
        _ => value.ToString() ?? string.Empty,
    };

    private static string Format(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);
}
