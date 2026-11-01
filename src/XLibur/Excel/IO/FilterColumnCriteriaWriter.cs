using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.AutoFilters;
using XLibur.Utils;

namespace XLibur.Excel.IO;

/// <summary>
/// Writes <see cref="XLFilterColumnCriteria"/> back to a <c>filterColumn</c>. The counterpart of
/// <see cref="FilterColumnCriteriaReader"/>, and the only place that spells the criteria out, so
/// the two cannot drift.
/// </summary>
internal static class FilterColumnCriteriaWriter
{
    internal static FilterColumn Write(XLFilterColumnCriteria criteria)
    {
        var filterColumn = new FilterColumn { ColumnId = criteria.ColumnId };

        // Both are written only when they differ from their schema default, so a column Excel
        // wrote without them comes back without them.
        if (criteria.HiddenButton)
            filterColumn.HiddenButton = true;

        if (!criteria.ShowButton)
            filterColumn.ShowButton = false;

        if (criteria.Criteria is { } childCriteria)
            filterColumn.Append(WriteCriteria(childCriteria));

        if (criteria.ExtensionListXml is { } extensionListXml)
            filterColumn.Append(new ExtensionList(extensionListXml));

        return filterColumn;
    }

    private static OpenXmlElement WriteCriteria(XLFilterCriteria criteria)
    {
        return criteria switch
        {
            XLValuesFilterCriteria values => WriteValuesFilter(values),
            XLTop10Criteria top10 => WriteTop10(top10),
            XLCustomFiltersCriteria customFilters => WriteCustomFilters(customFilters),
            XLDynamicFilterCriteria dynamicFilter => WriteDynamicFilter(dynamicFilter),
            XLColorFilterCriteria colorFilter => WriteColorFilter(colorFilter),
            XLIconFilterCriteria iconFilter => WriteIconFilter(iconFilter),
            _ => throw new NotSupportedException($"Unexpected filter criteria: {criteria.GetType().Name}."),
        };
    }

    private static Filters WriteValuesFilter(XLValuesFilterCriteria criteria)
    {
        var filters = new Filters();
        if (criteria.Blank)
            filters.Blank = true;

        if (criteria.CalendarType is { } calendarType)
            SetTokenAttribute(filters, "calendarType", calendarType);

        // Schema order: every filter, then every dateGroupItem.
        foreach (var value in criteria.Values)
            filters.Append(new Filter { Val = value });

        foreach (var dateGroup in criteria.DateGroups)
            filters.Append(WriteDateGroup(dateGroup));

        return filters;
    }

    private static DateGroupItem WriteDateGroup(XLDateGroupCriteria criteria)
    {
        var dateGroupItem = new DateGroupItem { DateTimeGrouping = criteria.Grouping.ToOpenXml() };

        if (criteria.Year is { } year) dateGroupItem.Year = year;
        if (criteria.Month is { } month) dateGroupItem.Month = month;
        if (criteria.Day is { } day) dateGroupItem.Day = day;
        if (criteria.Hour is { } hour) dateGroupItem.Hour = hour;
        if (criteria.Minute is { } minute) dateGroupItem.Minute = minute;
        if (criteria.Second is { } second) dateGroupItem.Second = second;

        return dateGroupItem;
    }

    private static Top10 WriteTop10(XLTop10Criteria criteria)
    {
        var top10 = new Top10 { Val = criteria.Value };

        if (!criteria.Top)
            top10.Top = false;

        if (criteria.Percent)
            top10.Percent = true;

        if (criteria.FilterValue is { } filterValue)
            top10.FilterValue = filterValue;

        return top10;
    }

    private static CustomFilters WriteCustomFilters(XLCustomFiltersCriteria criteria)
    {
        var customFilters = new CustomFilters();
        if (criteria.And)
            customFilters.And = true;

        foreach (var filter in criteria.Filters)
        {
            var customFilter = new CustomFilter();
            if (filter.Value is { } value)
                customFilter.Val = value;

            if (filter.Operator != XLFilterOperator.Equal)
                customFilter.Operator = filter.Operator.ToOpenXml();

            customFilters.Append(customFilter);
        }

        return customFilters;
    }

    private static DynamicFilter WriteDynamicFilter(XLDynamicFilterCriteria criteria)
    {
        var dynamicFilter = new DynamicFilter();
        SetTokenAttribute(dynamicFilter, "type", criteria.Type);

        if (criteria.Value is { } value)
            dynamicFilter.Val = value;

        if (criteria.MaxValue is { } maxValue)
            dynamicFilter.MaxVal = maxValue;

        // The ISO forms go back as the text they were read as: parsing them into a DateTime and
        // reformatting would change the precision Excel wrote.
        if (criteria.ValueIso is { } valueIso)
            SetTokenAttribute(dynamicFilter, "valIso", valueIso);

        if (criteria.MaxValueIso is { } maxValueIso)
            SetTokenAttribute(dynamicFilter, "maxValIso", maxValueIso);

        return dynamicFilter;
    }

    private static ColorFilter WriteColorFilter(XLColorFilterCriteria criteria)
    {
        var colorFilter = new ColorFilter();
        if (criteria.DifferentialFormatId is { } formatId)
            colorFilter.FormatId = formatId;

        if (!criteria.CellColor)
            colorFilter.CellColor = false;

        return colorFilter;
    }

    private static IconFilter WriteIconFilter(XLIconFilterCriteria criteria)
    {
        var iconFilter = new IconFilter();
        SetTokenAttribute(iconFilter, "iconSet", criteria.IconSet);

        if (criteria.IconId is { } iconId)
            iconFilter.IconId = iconId;

        return iconFilter;
    }

    /// <summary>
    /// Set an attribute the SDK types as an enum from its raw token, so a value this build has
    /// never heard of is written back as it was read rather than dropped.
    /// </summary>
    private static void SetTokenAttribute(OpenXmlElement element, string attributeName, string token)
    {
        element.SetAttribute(new OpenXmlAttribute(string.Empty, attributeName, string.Empty, token));
    }
}
