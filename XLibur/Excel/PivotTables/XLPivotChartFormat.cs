namespace XLibur.Excel;

/// <summary>
/// A link between a pivot area and a formatting record of a pivot chart. Excel writes one per
/// chart element it has formatted by hand, so dropping them loses the manual formatting of a
/// PivotChart even though the chart part itself survives.
/// </summary>
/// <remarks>
/// Distinct from <see cref="XLPivotTable.ChartFormat"/>, which is the pivot table's own
/// <c>chartFormat</c> attribute — an id used by <c>/chartSpace/pivotSource/fmtId/@val</c> to tie
/// a chart to this table. This type is an entry of the <c>chartFormats</c> collection.
/// </remarks>
internal sealed class XLPivotChartFormat
{
    internal XLPivotChartFormat(XLPivotArea pivotArea)
    {
        PivotArea = pivotArea;
    }

    /// <summary>
    /// Pivot area the chart formatting applies to.
    /// </summary>
    internal XLPivotArea PivotArea { get; }

    /// <summary>
    /// Zero-based index of the chart element this record formats.
    /// </summary>
    internal uint Chart { get; init; }

    /// <summary>
    /// Index of the formatting record held by the chart part itself. XLibur does not interpret
    /// it — the value is round-tripped so the chart keeps pointing at the right record.
    /// </summary>
    internal uint Format { get; init; }

    /// <summary>
    /// Whether the record formats a whole data series rather than a single data point.
    /// </summary>
    internal bool Series { get; init; }
}
