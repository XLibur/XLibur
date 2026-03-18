using XLibur.Excel.AutoFilters;
using XLibur.Excel.ContentManagers;
using XLibur.Utils;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Globalization;
using System.Linq;
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
            var filterColumn = new FilterColumn { ColumnId = (uint)columnNumber - 1 };

            if (xlFilterColumn.FilterType == XLFilterType.None)
                continue;

            PopulateFilterColumn(filterColumn, xlFilterColumn, context);
            autoFilter.Append(filterColumn);
        }

        if (xlAutoFilter.Sorted)
            AppendSortState(autoFilter, xlAutoFilter, filterRange);
    }

    private static void PopulateFilterColumn(FilterColumn filterColumn, XLFilterColumn xlFilterColumn,
        SaveContext context)
    {
        switch (xlFilterColumn.FilterType)
        {
            case XLFilterType.Custom:
                filterColumn.Append(CreateCustomFilters(xlFilterColumn));
                break;

            case XLFilterType.TopBottom:
                filterColumn.Append(CreateTop10Filter(xlFilterColumn));
                break;

            case XLFilterType.Dynamic:
                filterColumn.Append(CreateDynamicFilter(xlFilterColumn));
                break;

            case XLFilterType.Regular:
                filterColumn.Append(CreateRegularFilters(xlFilterColumn));
                break;

            case XLFilterType.Color:
                AppendColorFilter(filterColumn, xlFilterColumn, context);
                break;

            default:
                throw new NotSupportedException();
        }
    }

    private static CustomFilters CreateCustomFilters(XLFilterColumn xlFilterColumn)
    {
        var customFilters = new CustomFilters();
        foreach (var xlFilter in xlFilterColumn)
        {
            var filterValue = xlFilter.CustomValue.ToString(CultureInfo.InvariantCulture);
            var customFilter = new CustomFilter { Val = filterValue };

            if (xlFilter.Operator != XLFilterOperator.Equal)
                customFilter.Operator = xlFilter.Operator.ToOpenXml();

            if (xlFilter.Connector == XLConnector.And)
                customFilters.And = true;

            customFilters.Append(customFilter);
        }

        return customFilters;
    }

    private static Top10 CreateTop10Filter(XLFilterColumn xlFilterColumn)
    {
        return new Top10
        {
            Val = xlFilterColumn.TopBottomValue,
            Percent = OpenXmlHelper.GetBooleanValue(xlFilterColumn.TopBottomType == XLTopBottomType.Percent, false),
            Top = OpenXmlHelper.GetBooleanValue(xlFilterColumn.TopBottomPart == XLTopBottomPart.Top, true)
        };
    }

    private static DynamicFilter CreateDynamicFilter(XLFilterColumn xlFilterColumn)
    {
        return new DynamicFilter
        {
            Type = xlFilterColumn.DynamicType.ToOpenXml(),
            Val = xlFilterColumn.DynamicValue
        };
    }

    private static Filters CreateRegularFilters(XLFilterColumn xlFilterColumn)
    {
        var filters = new Filters();
        foreach (var filter in xlFilterColumn)
        {
            if (filter.Value is string s)
                filters.Append(new Filter { Val = s });
        }

        foreach (var filter in xlFilterColumn)
        {
            if (filter.Value is DateTime time)
                filters.Append(CreateDateGroupItem(filter, time));
        }

        return filters;
    }

    private static DateGroupItem CreateDateGroupItem(XLFilter filter, DateTime time)
    {
        var dgi = new DateGroupItem
        {
            Year = (ushort)time.Year,
            DateTimeGrouping = filter.DateTimeGrouping.ToOpenXml()
        };

        if (filter.DateTimeGrouping >= XLDateTimeGrouping.Month) dgi.Month = (ushort)time.Month;
        if (filter.DateTimeGrouping >= XLDateTimeGrouping.Day) dgi.Day = (ushort)time.Day;
        if (filter.DateTimeGrouping >= XLDateTimeGrouping.Hour) dgi.Hour = (ushort)time.Hour;
        if (filter.DateTimeGrouping >= XLDateTimeGrouping.Minute) dgi.Minute = (ushort)time.Minute;
        if (filter.DateTimeGrouping >= XLDateTimeGrouping.Second) dgi.Second = (ushort)time.Second;

        return dgi;
    }

    private static void AppendColorFilter(FilterColumn filterColumn, XLFilterColumn xlFilterColumn,
        SaveContext context)
    {
        var dxfKey = (xlFilterColumn.FilterColor.Key, xlFilterColumn.FilterByCellColor);
        if (context.ColorFilterDxfIds.TryGetValue(dxfKey, out var dxfId))
        {
            var colorFilter = new ColorFilter
            {
                FormatId = (uint)dxfId,
                CellColor = OpenXmlHelper.GetBooleanValue(xlFilterColumn.FilterByCellColor, true),
            };
            filterColumn.Append(colorFilter);
        }
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
