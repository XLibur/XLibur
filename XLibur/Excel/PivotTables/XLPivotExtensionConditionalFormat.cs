using System.Collections.Generic;

namespace XLibur.Excel;

/// <summary>
/// A pivot table's link, by rule id, to a conditional format rule that its sheet keeps only in the
/// sheet's <c>x14</c> extension, marked <c>pivot="1"</c>. One entry of the <c>x14:conditionalFormats</c>
/// list in the pivot table's own <c>x14</c> extension.
/// </summary>
/// <remarks>
/// Excel writes a rule only in the extension when the 2007 schema cannot hold it, as for a formula
/// that refers to another sheet. XLibur does not model such a rule. The sheet writes it back as it
/// was loaded, id included, and only its formula text, its range and its priority follow an edit. So
/// this keeps the id as it was loaded, and a save names the rule by the priority the sheet wrote it
/// with (<see cref="XLibur.Excel.ConditionalFormats.XLConditionalFormats.TryGetExtensionRulePriority"/>,
/// #552), and the two sides go
/// on naming each other. A rule
/// the sheet holds in the 2007 schema is linked by priority instead, through
/// <see cref="XLPivotConditionalFormat"/>.
/// </remarks>
internal sealed class XLPivotExtensionConditionalFormat
{
    private readonly List<XLPivotArea> _areas = new();

    internal XLPivotExtensionConditionalFormat(string ruleId, uint priority)
    {
        RuleId = ruleId;
        Priority = priority;
    }

    /// <summary>
    /// The <c>id</c> of the rule on the sheet, braces included, as it was loaded.
    /// </summary>
    internal string RuleId { get; }

    /// <summary>
    /// The <c>priority</c> of the rule on the sheet, as it was loaded. A save writes the priority the
    /// sheet gave the rule instead, and this only where the sheet holds no rule with this id.
    /// </summary>
    internal uint Priority { get; }

    /// <inheritdoc cref="XLPivotConditionalFormat.Scope"/>
    internal XLPivotCfScope Scope { get; init; } = XLPivotCfScope.SelectedCells;

    /// <inheritdoc cref="XLPivotConditionalFormat.Type"/>
    internal XLPivotCfRuleType Type { get; init; } = XLPivotCfRuleType.None;

    /// <summary>
    /// Areas of the pivot table the rule applies to.
    /// </summary>
    internal IReadOnlyList<XLPivotArea> Areas => _areas;

    internal void AddArea(XLPivotArea pivotArea)
    {
        _areas.Add(pivotArea);
    }

    /// <inheritdoc cref="XLPivotConditionalFormat.RemoveEmptiedAreas"/>
    /// <remarks>
    /// The 2007 list and this one are treated as one population for this pruning (#552, #585): a fix
    /// here always runs alongside <see cref="XLPivotConditionalFormat.RemoveEmptiedAreas"/>, never in
    /// its place.
    /// </remarks>
    internal bool RemoveEmptiedAreas(uint removedPosition, uint remainingValueCount)
    {
        var hadAreas = _areas.Count > 0;
        _areas.RemoveAll(area => XLPivotTable.RenumberDataFieldPositions(area, removedPosition, remainingValueCount));
        return hadAreas && _areas.Count == 0;
    }
}
