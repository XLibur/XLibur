using System.Collections.Generic;
using XLibur.Excel.AutoFilters;

namespace XLibur.Excel;

/// <summary>
/// The <c>autoFilter</c> child of a pivot table filter (<c>CT_AutoFilter</c>) — the criteria that
/// say what the filter actually does.
/// </summary>
/// <remarks>
/// <para>
/// Despite the shared element type, this is not a worksheet autofilter: it is a criteria record
/// Excel applies to pivot items, so it hides no rows and its <see cref="Reference"/> does not
/// describe any. That is why it holds <see cref="XLFilterColumnCriteria"/> rather than an
/// <see cref="XLAutoFilter"/>, which owns a live range and evaluates against cells.
/// </para>
/// <para>
/// <c>sortState</c> and <c>extLst</c> are carried verbatim. Sorting is modelled elsewhere in
/// XLibur far more narrowly than <c>CT_SortState</c> allows, and pivot filters do not appear to
/// use it in practice, so preserving the sub-tree beats modelling it badly.
/// </para>
/// </remarks>
internal sealed class XLPivotAutoFilter
{
    /// <summary>
    /// The <c>ref</c> attribute. Optional, and not a range XLibur reads anything from.
    /// </summary>
    internal string? Reference { get; init; }

    /// <summary>
    /// The <c>filterColumn</c> children. Empty is legal, if pointless.
    /// </summary>
    internal IReadOnlyList<XLFilterColumnCriteria> Columns { get; init; } = [];

    /// <summary>
    /// The <c>sortState</c> child, verbatim, including its element tags.
    /// </summary>
    internal string? SortStateXml { get; init; }

    /// <summary>
    /// The <c>extLst</c> child, verbatim, including its element tags.
    /// </summary>
    internal string? ExtensionListXml { get; init; }
}
