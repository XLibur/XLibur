using System.Collections.Generic;

namespace XLibur.Excel.AutoFilters;

/// <summary>
/// One of the six mutually exclusive children a <c>filterColumn</c> may carry. The schema models
/// them as a choice, so a column has exactly one or none.
/// </summary>
internal abstract class XLFilterCriteria
{
    private protected XLFilterCriteria()
    {
    }
}

/// <summary>
/// <c>filters</c> (<c>CT_Filters</c>) — an explicit list of the values to keep, the shape Excel
/// writes for a tick-box list.
/// </summary>
internal sealed class XLValuesFilterCriteria : XLFilterCriteria
{
    /// <summary>
    /// Whether blank values are kept as well. Default <c>false</c>.
    /// </summary>
    internal bool Blank { get; init; }

    /// <summary>
    /// The <c>ST_CalendarType</c> token the date groups are expressed in, e.g. <c>gregorian</c>
    /// or <c>hijri</c>. Held as the raw token: XLibur does not act on it, and there are a dozen
    /// values a newer Excel could add to. Default is <c>none</c>, written as absent.
    /// </summary>
    internal string? CalendarType { get; init; }

    /// <summary>
    /// The kept values, as written (<c>filter/@val</c>). Excel compares them against the
    /// formatted text of a cell, so they are text even for numeric columns.
    /// </summary>
    internal IReadOnlyList<string> Values { get; init; } = [];

    /// <summary>
    /// The kept date groups (<c>dateGroupItem</c>), e.g. "every March".
    /// </summary>
    internal IReadOnlyList<XLDateGroupCriteria> DateGroups { get; init; } = [];
}

/// <summary>
/// One <c>dateGroupItem</c> — a date truncated to <see cref="Grouping"/>, matching every date
/// that shares that prefix.
/// </summary>
/// <remarks>
/// The parts are nullable rather than defaulted because the schema makes only <c>year</c> and
/// <c>dateTimeGrouping</c> required, and a part that was absent must stay absent on save.
/// </remarks>
internal sealed class XLDateGroupCriteria
{
    internal required XLDateTimeGrouping Grouping { get; init; }

    internal ushort? Year { get; init; }

    internal ushort? Month { get; init; }

    internal ushort? Day { get; init; }

    internal ushort? Hour { get; init; }

    internal ushort? Minute { get; init; }

    internal ushort? Second { get; init; }
}

/// <summary>
/// <c>top10</c> (<c>CT_Top10</c>) — keep the highest or lowest N items, or N percent of them.
/// </summary>
internal sealed class XLTop10Criteria : XLFilterCriteria
{
    /// <summary>
    /// Whether the top of the range is kept rather than the bottom. Default <c>true</c>.
    /// </summary>
    internal bool Top { get; init; } = true;

    /// <summary>
    /// Whether <see cref="Value"/> is a percentage rather than an item count. Default <c>false</c>.
    /// </summary>
    internal bool Percent { get; init; }

    /// <summary>
    /// The item count or percentage (<c>val</c>). Required by the schema.
    /// </summary>
    internal required double Value { get; init; }

    /// <summary>
    /// The cached cut-off the count resolved to when the filter was last applied
    /// (<c>filterVal</c>). Optional, and stale as soon as the data changes, so it is preserved
    /// rather than relied on.
    /// </summary>
    internal double? FilterValue { get; init; }
}

/// <summary>
/// <c>customFilters</c> (<c>CT_CustomFilters</c>) — up to two comparisons joined by AND or OR.
/// </summary>
internal sealed class XLCustomFiltersCriteria : XLFilterCriteria
{
    /// <summary>
    /// Whether the criteria are joined by AND rather than OR. Default <c>false</c>.
    /// </summary>
    internal bool And { get; init; }

    internal IReadOnlyList<XLCustomFilterCriterion> Filters { get; init; } = [];
}

/// <summary>
/// One <c>customFilter</c> — an operator and the value to compare against.
/// </summary>
internal sealed class XLCustomFilterCriterion
{
    /// <summary>
    /// The comparison. Default <see cref="XLFilterOperator.Equal"/>, which Excel treats as a
    /// wildcard match rather than an equality test.
    /// </summary>
    internal XLFilterOperator Operator { get; init; } = XLFilterOperator.Equal;

    /// <summary>
    /// The value compared against, as written. Optional in the schema, and always text: OOXML has
    /// no type for it, so a number is stored in its invariant form.
    /// </summary>
    internal string? Value { get; init; }
}

/// <summary>
/// <c>dynamicFilter</c> (<c>CT_DynamicFilter</c>) — criteria evaluated against the data rather
/// than fixed, such as "above average" or "this quarter".
/// </summary>
internal sealed class XLDynamicFilterCriteria : XLFilterCriteria
{
    /// <summary>
    /// The <c>ST_DynamicFilterType</c> token, e.g. <c>aboveAverage</c> or <c>lastQuarter</c>.
    /// Held as the raw token rather than <see cref="XLFilterDynamicType"/>, which covers only the
    /// two average variants: mapping the other three dozen onto it would silently turn
    /// "this year" into "above average" on save.
    /// </summary>
    internal required string Type { get; init; }

    /// <summary>
    /// The cached value the criteria resolved to — the average, for the average filters.
    /// </summary>
    internal double? Value { get; init; }

    /// <summary>
    /// Upper bound of the cached range, for the date filters that span one.
    /// </summary>
    internal double? MaxValue { get; init; }

    /// <summary>
    /// <see cref="Value"/> as an ISO 8601 date, written by Excel alongside the serial number.
    /// Preserved as text so it cannot drift from the serial through a parse and reformat.
    /// </summary>
    internal string? ValueIso { get; init; }

    /// <summary>
    /// <see cref="MaxValue"/> as an ISO 8601 date.
    /// </summary>
    internal string? MaxValueIso { get; init; }
}

/// <summary>
/// <c>colorFilter</c> (<c>CT_ColorFilter</c>) — keep the cells drawn in one colour.
/// </summary>
internal sealed class XLColorFilterCriteria : XLFilterCriteria
{
    /// <summary>
    /// Index into the differential formats, which is where the colour itself lives.
    /// </summary>
    internal uint? DifferentialFormatId { get; init; }

    /// <summary>
    /// Whether the fill colour is matched rather than the font colour. Default <c>true</c>.
    /// </summary>
    internal bool CellColor { get; init; } = true;
}

/// <summary>
/// <c>iconFilter</c> (<c>CT_IconFilter</c>) — keep the cells showing one icon of a conditional
/// formatting icon set.
/// </summary>
/// <remarks>
/// XLibur cannot evaluate this one: it would have to resolve the conditional format that assigns
/// the icons first. It is modelled anyway so that a file carrying one survives a round trip.
/// </remarks>
internal sealed class XLIconFilterCriteria : XLFilterCriteria
{
    /// <summary>
    /// The <c>ST_IconSetType</c> token, e.g. <c>3TrafficLights1</c>. Required by the schema.
    /// </summary>
    internal required string IconSet { get; init; }

    /// <summary>
    /// Zero-based index of the icon within the set. Absent means "no icon".
    /// </summary>
    internal uint? IconId { get; init; }
}
