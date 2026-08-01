using System.Collections.Generic;

namespace XLibur.Excel.AutoFilters;

/// <summary>
/// The criteria of one <c>filterColumn</c> (<c>CT_FilterColumn</c>), free of any binding to a
/// worksheet.
/// </summary>
/// <remarks>
/// <para>
/// A declarative record of what a filter says, not of what it does. <see cref="XLFilterColumn"/>
/// is the worksheet-bound counterpart: it evaluates criteria against cells and owns an
/// <see cref="XLAutoFilter.Range"/>. A pivot table's <c>autoFilter</c> has neither — Excel
/// applies its criteria to pivot items — so the two share this type rather than the other.
/// </para>
/// <para>
/// Every attribute is carried, including the ones XLibur cannot act on, because the criteria are
/// written back from here. Anything omitted here is lost on save.
/// </para>
/// </remarks>
internal sealed class XLFilterColumnCriteria
{
    /// <summary>
    /// Zero-based index of the filtered column, relative to the filtered range (<c>colId</c>).
    /// </summary>
    internal required uint ColumnId { get; init; }

    /// <summary>
    /// Whether the filter dropdown button is hidden. Default <c>false</c>.
    /// </summary>
    internal bool HiddenButton { get; init; }

    /// <summary>
    /// Whether the filter dropdown button is shown. Default <c>true</c>.
    /// </summary>
    internal bool ShowButton { get; init; } = true;

    /// <summary>
    /// The criteria themselves — one of the six children the schema allows, or <c>null</c> when
    /// the column carries none (a valid, if unusual, state: the element then only sets the
    /// button attributes).
    /// </summary>
    internal XLFilterCriteria? Criteria { get; init; }

    /// <summary>
    /// The <c>extLst</c> child, verbatim, including its element tags. Nothing in XLibur reads
    /// the extensions, but dropping them would lose whatever a newer Excel put there.
    /// </summary>
    internal string? ExtensionListXml { get; init; }
}
