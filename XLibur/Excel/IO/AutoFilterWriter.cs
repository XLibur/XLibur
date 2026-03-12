using XLibur.Excel.ContentManagers;
using XLibur.Utils;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Globalization;
using System.Linq;

namespace XLibur.Excel.IO;

internal sealed class AutoFilterWriter
{
    internal static void WriteAutoFilter(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet)
    {
        worksheet.RemoveAllChildren<AutoFilter>();
        if (xlWorksheet.AutoFilter.IsEnabled)
        {
            var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.AutoFilter);
            worksheet.InsertAfter(new AutoFilter(), previousElement);

            var autoFilter = worksheet.Elements<AutoFilter>().First();
            cm.SetElement(XLWorksheetContents.AutoFilter, autoFilter);

            PopulateAutoFilter(xlWorksheet.AutoFilter, autoFilter);
        }
        else
        {
            cm.SetElement(XLWorksheetContents.AutoFilter, null);
        }
    }

    internal static void PopulateAutoFilter(XLAutoFilter xlAutoFilter, AutoFilter autoFilter)
    {
        var filterRange = xlAutoFilter.Range;
        autoFilter.Reference = filterRange.RangeAddress.ToString();

        foreach (var (columnNumber, xlFilterColumn) in xlAutoFilter.Columns)
        {
            var filterColumn = new FilterColumn { ColumnId = (uint)columnNumber - 1 };

            switch (xlFilterColumn.FilterType)
            {
                case XLFilterType.Custom:
                    var customFilters = new CustomFilters();
                    foreach (var xlFilter in xlFilterColumn)
                    {
                        // Since OOXML allows only string, the operand for custom filter must be serialized.
                        var filterValue = xlFilter.CustomValue.ToString(CultureInfo.InvariantCulture);
                        var customFilter = new CustomFilter { Val = filterValue };

                        if (xlFilter.Operator != XLFilterOperator.Equal)
                            customFilter.Operator = xlFilter.Operator.ToOpenXml();

                        if (xlFilter.Connector == XLConnector.And)
                            customFilters.And = true;

                        customFilters.Append(customFilter);
                    }

                    filterColumn.Append(customFilters);
                    break;

                case XLFilterType.TopBottom:
                    // Although there is a FilterValue attribute, populating it seems like more
                    // trouble than it's worth due to consistency issues. It's optional, so we
                    // can't rely on it during a load anyway.
                    var top101 = new Top10
                    {
                        Val = xlFilterColumn.TopBottomValue,
                        Percent = OpenXmlHelper.GetBooleanValue(xlFilterColumn.TopBottomType == XLTopBottomType.Percent,
                            false),
                        Top = OpenXmlHelper.GetBooleanValue(xlFilterColumn.TopBottomPart == XLTopBottomPart.Top, true)
                    };
                    filterColumn.Append(top101);
                    break;

                case XLFilterType.Dynamic:
                    var dynamicFilter = new DynamicFilter
                    {
                        Type = xlFilterColumn.DynamicType.ToOpenXml(),
                        Val = xlFilterColumn.DynamicValue
                    };
                    filterColumn.Append(dynamicFilter);
                    break;

                case XLFilterType.Regular:
                    var filters = new Filters();
                    foreach (var filter in xlFilterColumn)
                    {
                        if (filter.Value is string s)
                            filters.Append(new Filter { Val = s });
                    }

                    foreach (var filter in xlFilterColumn)
                    {
                        if (filter.Value is DateTime time)
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

                            filters.Append(dgi);
                        }
                    }

                    filterColumn.Append(filters);
                    break;

                case XLFilterType.None:
                    continue;

                default:
                    throw new NotSupportedException();
            }

            autoFilter.Append(filterColumn);
        }

        if (xlAutoFilter.Sorted)
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
}
