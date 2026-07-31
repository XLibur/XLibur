namespace XLibur.Excel;

/// <summary>
/// A filter applied to a pivot table field — the label, value, date and top-N filters Excel
/// offers from a field's dropdown. One entry of the <c>filters</c> collection.
/// </summary>
/// <remarks>
/// <para>
/// Distinct from <see cref="XLPivotTable.Filters"/>, which models the report-filter axis (the
/// <c>pageFields</c> element, shown above the table). These are the filters that narrow which
/// items of a field appear at all.
/// </para>
/// <para>
/// XLibur has no API for creating or interpreting these yet, so the type exists to carry them
/// through a load/save unchanged. The attributes are modelled because they are the part anyone
/// exposing an API would need; <see cref="AutoFilterXml"/> is kept verbatim because
/// <c>CT_AutoFilter</c> is a whole sub-tree of its own (<c>filterColumn</c> with any of
/// <c>filters</c>, <c>top10</c>, <c>customFilters</c>, <c>dynamicFilter</c>, <c>colorFilter</c>
/// or <c>iconFilter</c>) and modelling it would be a large surface with nothing reading it.
/// </para>
/// </remarks>
internal sealed class XLPivotFilter
{
    internal XLPivotFilter(uint field, uint id, string type, string autoFilterXml)
    {
        Field = field;
        Id = id;
        Type = type;
        AutoFilterXml = autoFilterXml;
    }

    /// <summary>
    /// Index of the pivot field the filter applies to.
    /// </summary>
    internal uint Field { get; }

    /// <summary>
    /// Identifier of the filter, unique within the pivot table.
    /// </summary>
    internal uint Id { get; }

    /// <summary>
    /// The <c>ST_PivotFilterType</c> token, e.g. <c>captionEqual</c> or <c>dateNewerThan</c>.
    /// Held as the raw token rather than an enum: there are around forty values, XLibur does not
    /// act on any of them, and a token round-trips a file written by a newer Excel that uses a
    /// value this build has never heard of.
    /// </summary>
    internal string Type { get; }

    /// <summary>
    /// The <c>autoFilter</c> child, verbatim, including its element tags. Required by the
    /// schema, so every filter has one.
    /// </summary>
    internal string AutoFilterXml { get; }

    /// <summary>
    /// Order the filter is evaluated in relative to the other filters. Default <c>0</c>.
    /// </summary>
    internal int EvaluationOrder { get; init; }

    /// <summary>
    /// Index of the OLAP member property field the filter is based on (<c>mpFld</c>).
    /// </summary>
    internal uint? MemberPropertyField { get; init; }

    /// <summary>
    /// OLAP measure hierarchy the filter is based on (<c>iMeasureHier</c>).
    /// </summary>
    internal uint? MeasureHierarchy { get; init; }

    /// <summary>
    /// OLAP measure field the filter is based on (<c>iMeasureFld</c>).
    /// </summary>
    internal uint? MeasureField { get; init; }

    /// <summary>
    /// Name shown for the filter.
    /// </summary>
    internal string? Name { get; init; }

    /// <summary>
    /// Description shown for the filter.
    /// </summary>
    internal string? Description { get; init; }

    /// <summary>
    /// First string operand of the filter, for the types that take one.
    /// </summary>
    internal string? StringValue1 { get; init; }

    /// <summary>
    /// Second string operand, for the types that take a range (e.g. <c>captionBetween</c>).
    /// </summary>
    internal string? StringValue2 { get; init; }
}
