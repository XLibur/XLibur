using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.AutoFilters;
using XLibur.Utils;

namespace XLibur.Excel.IO;

/// <summary>
/// Reads a <c>filterColumn</c> into <see cref="XLFilterColumnCriteria"/>. The one place that
/// knows how the criteria are spelled, shared by worksheet autofilters and pivot table filters —
/// both of which reach the element through the same DOM.
/// </summary>
internal static class FilterColumnCriteriaReader
{
    internal static XLFilterColumnCriteria Read(FilterColumn filterColumn)
    {
        return new XLFilterColumnCriteria
        {
            ColumnId = filterColumn.ColumnId?.Value ?? throw PartStructureException.MissingAttribute(),
            HiddenButton = OpenXmlHelper.GetBooleanValueAsBool(filterColumn.HiddenButton, false),
            ShowButton = OpenXmlHelper.GetBooleanValueAsBool(filterColumn.ShowButton, true),
            Criteria = ReadCriteria(filterColumn),
            ExtensionListXml = filterColumn.GetFirstChild<ExtensionList>()?.OuterXml,
        };
    }

    /// <summary>
    /// The schema models the six children as a choice, so at most one is present. They are tried
    /// in no particular order; a file with several is malformed either way.
    /// </summary>
    private static XLFilterCriteria? ReadCriteria(FilterColumn filterColumn)
    {
        if (filterColumn.GetFirstChild<Filters>() is { } filters)
            return ReadValuesFilter(filters);

        if (filterColumn.GetFirstChild<Top10>() is { } top10)
            return ReadTop10(top10);

        if (filterColumn.GetFirstChild<CustomFilters>() is { } customFilters)
            return ReadCustomFilters(customFilters);

        if (filterColumn.GetFirstChild<DynamicFilter>() is { } dynamicFilter)
            return ReadDynamicFilter(dynamicFilter);

        if (filterColumn.GetFirstChild<ColorFilter>() is { } colorFilter)
            return ReadColorFilter(colorFilter);

        if (filterColumn.GetFirstChild<IconFilter>() is { } iconFilter)
            return ReadIconFilter(iconFilter);

        return null;
    }

    private static XLValuesFilterCriteria ReadValuesFilter(Filters filters)
    {
        var values = filters.Elements<Filter>()
            .Select(filter => filter.Val?.Value)
            .Where(val => val is not null)
            .Select(val => val!)
            .ToList();

        var dateGroups = filters.Elements<DateGroupItem>()
            .Select(ReadDateGroup)
            .Where(group => group is not null)
            .Select(group => group!)
            .ToList();

        return new XLValuesFilterCriteria
        {
            Blank = OpenXmlHelper.GetBooleanValueAsBool(filters.Blank, false),
            CalendarType = filters.CalendarType?.InnerText,
            Values = values,
            DateGroups = dateGroups,
        };
    }

    /// <summary>
    /// Returns <c>null</c> for an item with no grouping. The attribute is required, so its
    /// absence means the item says nothing about which dates to keep.
    /// </summary>
    private static XLDateGroupCriteria? ReadDateGroup(DateGroupItem dateGroupItem)
    {
        if (dateGroupItem.DateTimeGrouping is not { HasValue: true } grouping)
            return null;

        return new XLDateGroupCriteria
        {
            Grouping = grouping.Value.ToXLibur(),
            Year = dateGroupItem.Year?.Value,
            Month = dateGroupItem.Month?.Value,
            Day = dateGroupItem.Day?.Value,
            Hour = dateGroupItem.Hour?.Value,
            Minute = dateGroupItem.Minute?.Value,
            Second = dateGroupItem.Second?.Value,
        };
    }

    private static XLTop10Criteria ReadTop10(Top10 top10)
    {
        return new XLTop10Criteria
        {
            Top = OpenXmlHelper.GetBooleanValueAsBool(top10.Top, true),
            Percent = OpenXmlHelper.GetBooleanValueAsBool(top10.Percent, false),
            Value = top10.Val?.Value ?? throw PartStructureException.MissingAttribute(),
            FilterValue = top10.FilterValue?.Value,
        };
    }

    private static XLCustomFiltersCriteria ReadCustomFilters(CustomFilters customFilters)
    {
        var filters = customFilters.Elements<CustomFilter>()
            .Select(customFilter => new XLCustomFilterCriterion
            {
                Operator = customFilter.Operator is { HasValue: true } op
                    ? op.Value.ToXLibur()
                    : XLFilterOperator.Equal,
                Value = customFilter.Val?.Value,
            })
            .ToList();

        return new XLCustomFiltersCriteria
        {
            And = OpenXmlHelper.GetBooleanValueAsBool(customFilters.And, false),
            Filters = filters,
        };
    }

    private static XLDynamicFilterCriteria ReadDynamicFilter(DynamicFilter dynamicFilter)
    {
        return new XLDynamicFilterCriteria
        {
            // The token, not the enum: XLFilterDynamicType covers only the two average variants.
            Type = dynamicFilter.Type?.InnerText ?? throw PartStructureException.MissingAttribute(),
            Value = dynamicFilter.Val?.Value,
            MaxValue = dynamicFilter.MaxVal?.Value,
            ValueIso = dynamicFilter.ValIso?.InnerText,
            MaxValueIso = dynamicFilter.MaxValIso?.InnerText,
        };
    }

    private static XLColorFilterCriteria ReadColorFilter(ColorFilter colorFilter)
    {
        return new XLColorFilterCriteria
        {
            DifferentialFormatId = colorFilter.FormatId?.Value,
            CellColor = OpenXmlHelper.GetBooleanValueAsBool(colorFilter.CellColor, true),
        };
    }

    private static XLIconFilterCriteria ReadIconFilter(IconFilter iconFilter)
    {
        return new XLIconFilterCriteria
        {
            IconSet = iconFilter.IconSet?.InnerText ?? throw PartStructureException.MissingAttribute(),
            IconId = iconFilter.IconId?.Value,
        };
    }
}
