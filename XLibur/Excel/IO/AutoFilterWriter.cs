using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.AutoFilters;
using XLibur.Excel.ContentManagers;
using static XLibur.Excel.XLWorkbook;

namespace XLibur.Excel.IO;

internal static class AutoFilterWriter
{
    internal static void WriteAutoFilter(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet,
        SaveContext context)
    {
        worksheet.RemoveAllChildren<AutoFilter>();
        if (xlWorksheet.AutoFilter.IsEnabled)
        {
            var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.AutoFilter);
            worksheet.InsertAfter(new AutoFilter(), previousElement);

            var autoFilter = worksheet.Elements<AutoFilter>().First();
            cm.SetElement(XLWorksheetContents.AutoFilter, autoFilter);

            PopulateAutoFilter(xlWorksheet.AutoFilter, autoFilter, context);
        }
        else
        {
            cm.SetElement(XLWorksheetContents.AutoFilter, null);
        }
    }

    internal static void PopulateAutoFilter(XLAutoFilter xlAutoFilter, AutoFilter autoFilter, SaveContext context)
    {
        var filterRange = xlAutoFilter.Range;
        autoFilter.Reference = filterRange.RangeAddress.ToString();

        foreach (var (columnNumber, xlFilterColumn) in xlAutoFilter.Columns)
        {
            if (GetCriteria(xlFilterColumn, (uint)columnNumber - 1, context) is not { } criteria)
                continue;

            autoFilter.Append(FilterColumnCriteriaWriter.Write(criteria));
        }

        if (xlAutoFilter.Sorted)
            AppendSortState(autoFilter, xlAutoFilter, filterRange);
    }

    /// <summary>
    /// The criteria to write for one column, or <c>null</c> when it has nothing to say.
    /// </summary>
    /// <remarks>
    /// A column that was loaded and not since changed is written from the criteria it was loaded
    /// with, so the parts the runtime state cannot hold — an <c>iconFilter</c>, the button
    /// attributes, <c>extLst</c> — are not lost by a load and save. Once the caller changes the
    /// column those criteria are dropped, and it is written from what the caller asked for.
    /// </remarks>
    private static XLFilterColumnCriteria? GetCriteria(XLFilterColumn xlFilterColumn, uint columnId,
        SaveContext context)
    {
        if (xlFilterColumn.SourceCriteria is { } sourceCriteria)
            return sourceCriteria;

        if (xlFilterColumn.FilterType == XLFilterType.None)
            return null;

        return new XLFilterColumnCriteria
        {
            ColumnId = columnId,
            Criteria = CreateCriteria(xlFilterColumn, context),
        };
    }

    private static XLFilterCriteria? CreateCriteria(XLFilterColumn xlFilterColumn, SaveContext context)
    {
        return xlFilterColumn.FilterType switch
        {
            XLFilterType.Custom => CreateCustomFilters(xlFilterColumn),
            XLFilterType.TopBottom => CreateTop10Filter(xlFilterColumn),
            XLFilterType.Dynamic => CreateDynamicFilter(xlFilterColumn),
            XLFilterType.Regular => CreateRegularFilters(xlFilterColumn),

            // A colour whose differential format was never registered has no dxfId to point at,
            // so the column is written without criteria rather than with a dangling reference.
            XLFilterType.Color => CreateColorFilter(xlFilterColumn, context),
            _ => throw new NotSupportedException(),
        };
    }

    private static XLCustomFiltersCriteria CreateCustomFilters(XLFilterColumn xlFilterColumn)
    {
        var filters = new List<XLCustomFilterCriterion>();
        var and = false;

        foreach (var xlFilter in xlFilterColumn)
        {
            filters.Add(new XLCustomFilterCriterion
            {
                Value = xlFilter.CustomValue.ToString(CultureInfo.InvariantCulture),
                Operator = xlFilter.Operator,
            });

            if (xlFilter.Connector == XLConnector.And)
                and = true;
        }

        return new XLCustomFiltersCriteria { And = and, Filters = filters };
    }

    private static XLTop10Criteria CreateTop10Filter(XLFilterColumn xlFilterColumn)
    {
        return new XLTop10Criteria
        {
            Value = xlFilterColumn.TopBottomValue,
            Percent = xlFilterColumn.TopBottomType == XLTopBottomType.Percent,
            Top = xlFilterColumn.TopBottomPart == XLTopBottomPart.Top,
        };
    }

    private static XLDynamicFilterCriteria CreateDynamicFilter(XLFilterColumn xlFilterColumn)
    {
        return new XLDynamicFilterCriteria
        {
            Type = new EnumValue<DynamicFilterValues>(xlFilterColumn.DynamicType.ToOpenXml()).InnerText!,
            Value = xlFilterColumn.DynamicValue,
        };
    }

    private static XLValuesFilterCriteria CreateRegularFilters(XLFilterColumn xlFilterColumn)
    {
        var values = new List<string>();
        var dateGroups = new List<XLDateGroupCriteria>();

        foreach (var filter in xlFilterColumn)
        {
            switch (filter.Value)
            {
                case string text:
                    values.Add(text);
                    break;

                case DateTime time:
                    dateGroups.Add(CreateDateGroup(filter, time));
                    break;
            }
        }

        return new XLValuesFilterCriteria { Values = values, DateGroups = dateGroups };
    }

    /// <summary>
    /// A date truncated to the filter's grouping. The parts finer than the grouping are left out
    /// rather than written as the zeroes the <see cref="DateTime"/> carries, because Excel reads
    /// a present part as one the filter matches on.
    /// </summary>
    private static XLDateGroupCriteria CreateDateGroup(XLFilter filter, DateTime time)
    {
        var grouping = filter.DateTimeGrouping;
        return new XLDateGroupCriteria
        {
            Grouping = grouping,
            Year = (ushort)time.Year,
            Month = grouping >= XLDateTimeGrouping.Month ? (ushort)time.Month : null,
            Day = grouping >= XLDateTimeGrouping.Day ? (ushort)time.Day : null,
            Hour = grouping >= XLDateTimeGrouping.Hour ? (ushort)time.Hour : null,
            Minute = grouping >= XLDateTimeGrouping.Minute ? (ushort)time.Minute : null,
            Second = grouping >= XLDateTimeGrouping.Second ? (ushort)time.Second : null,
        };
    }

    private static XLColorFilterCriteria? CreateColorFilter(XLFilterColumn xlFilterColumn, SaveContext context)
    {
        var dxfKey = (xlFilterColumn.FilterColor.Key, xlFilterColumn.FilterByCellColor);
        if (!context.ColorFilterDxfIds.TryGetValue(dxfKey, out var dxfId))
            return null;

        return new XLColorFilterCriteria
        {
            DifferentialFormatId = (uint)dxfId,
            CellColor = xlFilterColumn.FilterByCellColor,
        };
    }

    private static void AppendSortState(AutoFilter autoFilter, XLAutoFilter xlAutoFilter, IXLRange filterRange)
    {
        string reference;

        if (filterRange.FirstCell().Address.RowNumber < filterRange.LastCell().Address.RowNumber)
            reference = filterRange.Range(filterRange.FirstCell().CellBelow(), filterRange.LastCell()).RangeAddress
                .ToString()!;
        else
            reference = filterRange.RangeAddress.ToString()!;

        var sortState = new SortState
        {
            Reference = reference
        };

        var sortCondition = new SortCondition
        {
            Reference =
                filterRange.Range(1, xlAutoFilter.SortColumn, filterRange.RowCount(),
                    xlAutoFilter.SortColumn).RangeAddress.ToString()
        };
        if (xlAutoFilter.SortOrder == XLSortOrder.Descending)
            sortCondition.Descending = true;

        sortState.Append(sortCondition);
        autoFilter.Append(sortState);
    }
}
