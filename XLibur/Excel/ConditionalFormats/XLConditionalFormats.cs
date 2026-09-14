using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using XLibur.Excel.CalcEngine.Visitors;
using XLibur.Excel.Coordinates;
using XLibur.Extensions;

namespace XLibur.Excel.ConditionalFormats;

/// <summary>
/// A container for conditional formatting of a <see cref="XLWorksheet"/>. It contains
/// a collection of <see cref="XLConditionalFormat"/>. Doesn't contain pivot table formats,
/// they are in pivot table <see cref="XLPivotTable.ConditionalFormats"/>,
/// </summary>
internal sealed class XLConditionalFormats : IXLConditionalFormats, ISheetListener, IWorkbookListener
{
    private readonly List<IXLConditionalFormat> _conditionalFormats = [];

    /// <summary>
    /// The formula text of each <c>x14</c> rule this library keeps in the sheet's extension list but
    /// does not model, by the rule's id, in the order the rule holds it.
    /// </summary>
    /// <remarks>
    /// Excel writes a rule that refers to another sheet only in the <c>x14</c> extension (the
    /// <c>rename-*</c> and <c>delete-*</c> fixtures). XLibur loads a data bar rule from there into the
    /// model, and writes every other rule back as it was loaded, so this text is all of such a rule
    /// that it holds. A sheet rename or delete rewrites it here, and the writer puts it back
    /// (<c>ConditionalFormattingWriter</c>).
    /// </remarks>
    private readonly Dictionary<string, string[]> _extensionRuleFormulas = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// The range (<c>xm:sqref</c>) of each <c>x14</c> rule this library keeps but does not model, by
    /// the rule's id, held as the areas a modelled rule's coverage is held as.
    /// </summary>
    /// <remarks>
    /// A row or column insert or delete moves it by the transform that moves a modelled rule's
    /// <see cref="XLConditionalFormat.Areas"/> (issue #499), and the writer puts it back. An empty list
    /// is a rule the edit removed, as it removes a modelled rule whose coverage transforms to nothing.
    /// A rule whose range does not read as areas has no entry, and keeps its range as it was loaded.
    /// </remarks>
    private readonly Dictionary<string, XLAreaList> _extensionRuleAreas = new(StringComparer.OrdinalIgnoreCase);

    private readonly XLWorksheet _worksheet;

    internal XLConditionalFormats(XLWorksheet worksheet)
    {
        _worksheet = worksheet;
    }

    private static readonly List<XLConditionalFormatType> CFTypesExcludedFromConsolidation =
    [
        XLConditionalFormatType.DataBar,
        XLConditionalFormatType.ColorScale,
        XLConditionalFormatType.IconSet,
        XLConditionalFormatType.Top10,
        XLConditionalFormatType.AboveAverage,
        XLConditionalFormatType.IsDuplicate,
        XLConditionalFormatType.IsUnique
    ];

    public void Add(IXLConditionalFormat conditionalFormat)
    {
        _conditionalFormats.Add(conditionalFormat);
    }

    public IEnumerator<IXLConditionalFormat> GetEnumerator()
    {
        return _conditionalFormats.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public void Remove(Predicate<IXLConditionalFormat> predicate)
    {
        _conditionalFormats.RemoveAll(predicate);
    }

    /// <summary>
    /// Keeps the formula text of an <c>x14</c> rule this library does not model (see
    /// <see cref="_extensionRuleFormulas"/>).
    /// </summary>
    internal void SeedExtensionRuleFormulas(string ruleId, string[] formulas)
    {
        _extensionRuleFormulas[ruleId] = formulas;
    }

    /// <summary>
    /// The formula text of the unmodelled <c>x14</c> rule <paramref name="ruleId"/>, as a sheet rename
    /// or delete has left it.
    /// </summary>
    internal bool TryGetExtensionRuleFormulas(string ruleId, [NotNullWhen(true)] out string[]? formulas)
        => _extensionRuleFormulas.TryGetValue(ruleId, out formulas);

    /// <summary>
    /// Keeps the range of an <c>x14</c> rule this library does not model (see
    /// <see cref="_extensionRuleAreas"/>).
    /// </summary>
    internal void SeedExtensionRuleAreas(string ruleId, XLAreaList areas)
    {
        _extensionRuleAreas[ruleId] = areas;
    }

    /// <summary>
    /// The range of the unmodelled <c>x14</c> rule <paramref name="ruleId"/>, as row and column
    /// inserts and deletes have left it. Empty when an edit removed every cell it applied to.
    /// </summary>
    internal bool TryGetExtensionRuleAreas(string ruleId, [NotNullWhen(true)] out XLAreaList? areas)
        => _extensionRuleAreas.TryGetValue(ruleId, out areas);

    #region ISheetListener

    /// <summary>
    /// Moves every rule's coverage over the edit, removes a rule whose coverage transforms to
    /// nothing, and then re-points the references in every rule's formulas.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Coverage is the value-typed <see cref="XLAreaList"/>, not live repository ranges, so the
    /// transform is pure and can never alias or double-shift across overlapping coverage
    /// (ClosedXML issue #2850). Mirrors <see cref="XLDataValidations"/>, which shifts its sqref
    /// coverage the same way.
    /// </para>
    /// <para>
    /// The two passes are one listener, in this order, for the reason they are in
    /// <see cref="XLDataValidations"/>: the coverage pass removes a rule, and the formula pass must
    /// not then shift a rule that is gone. Their scopes differ. Coverage moves only for an edit on
    /// this sheet; formulas for an edit on any sheet, because a rule here may refer to another
    /// (<c>Data!$A$2&gt;0</c>). So every sheet's collection is a listener
    /// (<see cref="XLWorksheet.GetSheetListeners"/>), and the formula shifter does the filtering
    /// (issue #499, D77).
    /// </para>
    /// </remarks>
    void ISheetListener.OnInsertAreaAndShiftDown(in SheetEdit edit) => Shift<RowAxis>(in edit);

    void ISheetListener.OnInsertAreaAndShiftRight(in SheetEdit edit) => Shift<ColumnAxis>(in edit);

    void ISheetListener.OnDeleteAreaAndShiftUp(in SheetEdit edit) => Shift<RowAxis>(in edit);

    void ISheetListener.OnDeleteAreaAndShiftLeft(in SheetEdit edit) => Shift<ColumnAxis>(in edit);

    private void Shift<TAxis>(in SheetEdit edit)
        where TAxis : struct, IGridAxis
    {
        ShiftCoverage<TAxis>(in edit);
        ShiftFormulas<TAxis>(in edit);
    }

    private void ShiftCoverage<TAxis>(in SheetEdit edit)
        where TAxis : struct, IGridAxis
    {
        if (edit.Sheet != _worksheet || (_conditionalFormats.Count == 0 && _extensionRuleAreas.Count == 0))
            return;

        // CoverageArea, not edit.Area: coverage derives the |Shift| lines the edit moves from the
        // range and the shift rather than trusting the shifter's area. See SheetEdit.
        var axis = default(TAxis);
        var affected = edit.CoverageArea<TAxis>();
        foreach (var cf in _conditionalFormats.OfType<XLConditionalFormat>().ToList())
        {
            var newAreas = edit.Shift > 0
                ? axis.InsertAndShift(cf.Areas, affected)
                : axis.DeleteAndShift(cf.Areas, affected);

            if (newAreas.Count == 0)
            {
                Remove(f => f == cf);
                continue;
            }

            // Before the formula pass shifts the formulas; see TryGetSurvivingAnchor.
            if (TryGetSurvivingAnchor<TAxis>(in edit, affected, cf.Areas, newAreas, out var anchor, out var surviving))
                cf.RebaseFormulas(anchor, surviving);

            cf.SetAreas(newAreas);
        }

        // A kept x14 rule's range goes through the same transform, so it cannot part company with a
        // modelled rule over the same cells (issue #499). An emptied range stays, empty, so that the
        // writer knows to remove the rule.
        foreach (var (ruleId, areas) in _extensionRuleAreas.ToList())
        {
            if (areas.Count == 0)
                continue;

            var newAreas = edit.Shift > 0
                ? axis.InsertAndShift(areas, affected)
                : axis.DeleteAndShift(areas, affected);

            if (TryGetSurvivingAnchor<TAxis>(in edit, affected, areas, newAreas, out var anchor, out var surviving)
                && _extensionRuleFormulas.TryGetValue(ruleId, out var formulas))
            {
                for (var i = 0; i < formulas.Length; i++)
                {
                    if (XLConditionalFormat.TryRebaseFormula(formulas[i], anchor, surviving, out var rebased))
                        formulas[i] = rebased;
                }
            }

            _extensionRuleAreas[ruleId] = newAreas;
        }
    }

    /// <summary>
    /// Whether a delete removed a rule's anchor, the first cell of its range, which its formulas are
    /// written relative to, while the rule survives; and if so, the anchor and the cell to rebase the
    /// formulas onto: the rule's new first cell, where it stood before the edit.
    /// </summary>
    /// <remarks>
    /// Excel rebases such a rule's formulas onto the first cell that survives before it shifts them
    /// (<c>cf-anchor-*.xlsx</c>), so deleting row 2 leaves <c>$A2&gt;5</c> on <c>A2:C10</c> as
    /// <c>$A2&gt;5</c> on <c>A2:C9</c>, where shifting alone gives <c>#REF!&gt;5</c>. Any other edit
    /// only shifts, even one that cannot move the range: a row inserted at 1 turns <c>$A1&gt;5</c> on
    /// the whole column <c>E:E</c> into <c>$A2&gt;5</c>. A reference to a deleted cell other than the
    /// anchor still becomes <c>#REF!</c>, as it does in a cell formula.
    /// </remarks>
    /// <param name="edit">The edit, on this sheet.</param>
    /// <param name="deleted">The region the edit deletes, as coverage sees it.</param>
    /// <param name="before">The rule's range before the edit.</param>
    /// <param name="after">The rule's range after the edit; not empty.</param>
    /// <param name="anchor">The first cell of <paramref name="before"/>, the cell the formulas are written relative to.</param>
    /// <param name="surviving">The first cell of <paramref name="after"/>, at its position before the edit.</param>
    private static bool TryGetSurvivingAnchor<TAxis>(in SheetEdit edit, Area deleted, XLAreaList before,
        XLAreaList after, out Point anchor, out Point surviving)
        where TAxis : struct, IGridAxis
    {
        anchor = default;
        surviving = default;
        if (edit.Shift >= 0 || before.Count == 0 || after.Count == 0)
            return false;

        anchor = before[0].FirstPoint;
        if (!deleted.Contains(anchor))
            return false;

        // The new first cell moves back to where it stood: a cell past the deletion, within the
        // deleted lines' cross extent, had moved |Shift| lines towards it.
        var axis = default(TAxis);
        var first = after[0].FirstPoint;
        var moved = axis.IndexOf(first) >= axis.IndexOf(deleted.FirstPoint)
                    && axis.CrossOf(first) >= axis.CrossOf(deleted.FirstPoint)
                    && axis.CrossOf(first) <= axis.CrossOf(deleted.LastPoint);
        surviving = moved ? axis.PointAt(axis.IndexOf(first) - edit.Shift, axis.CrossOf(first)) : first;
        return true;
    }

    /// <summary>
    /// Re-points the references in the formulas of every rule on this sheet, modelled or kept as it
    /// was loaded, for an edit on any sheet (issue #499, D77).
    /// </summary>
    private void ShiftFormulas<TAxis>(in SheetEdit edit)
        where TAxis : struct, IGridAxis
    {
        foreach (var cf in _conditionalFormats.OfType<XLConditionalFormat>())
            cf.ShiftFormulas<TAxis>(in edit);

        foreach (var (ruleId, formulas) in _extensionRuleFormulas)
        {
            // A kept rule whose range an edit removed is not written back, so its text is left alone.
            if (_extensionRuleAreas.TryGetValue(ruleId, out var areas) && areas.Count == 0)
                continue;

            for (var i = 0; i < formulas.Length; i++)
            {
                if (XLConditionalFormat.TryShiftFormula<TAxis>(formulas[i], _worksheet, in edit, out var shifted))
                    formulas[i] = shifted;
            }
        }
    }

    #endregion ISheetListener

    #region IWorkbookListener

    /// <summary>
    /// The renamed sheet is renamed in every formula of every rule, modelled or kept as it was loaded:
    /// an expression and a scale's formula value alike. The <c>rename-*</c> fixture shows Excel doing
    /// so for both (D64).
    /// </summary>
    void IWorkbookListener.OnSheetRenamed(string oldSheetName, string newSheetName)
        => RewriteSheet(SheetRewrite.Rename(oldSheetName, newSheetName));

    /// <summary>
    /// A reference to the deleted sheet becomes <c>#REF!</c> in every formula of every rule. The
    /// <c>delete-*</c> fixture shows Excel writing <c>#REF!&gt;0</c> for the expression and
    /// <c>#REF!</c> for the scale's value (D64).
    /// </summary>
    void IWorkbookListener.OnSheetDeleting(string sheetName)
        => RewriteSheet(SheetRewrite.Delete(_worksheet.Workbook, sheetName));

    private void RewriteSheet(SheetRewrite rewrite)
    {
        foreach (var cf in _conditionalFormats.OfType<XLConditionalFormat>())
            cf.RewriteSheet(_worksheet.Name, rewrite);

        // A pivot table's formats are on this sheet too, though its pivot table holds them (#498).
        foreach (var pivotTable in _worksheet.PivotTables)
        {
            foreach (var pivotFormat in pivotTable.ConditionalFormats)
                pivotFormat.Format.RewriteSheet(_worksheet.Name, rewrite);
        }

        // The rewrite does not move a reference, so any origin reads the formula the same way.
        foreach (var formulas in _extensionRuleFormulas.Values)
        {
            for (var i = 0; i < formulas.Length; i++)
            {
                if (rewrite.TryRewrite(formulas[i], _worksheet.Name, new Point(1, 1), out var rewritten))
                    formulas[i] = rewritten;
            }
        }
    }

    #endregion IWorkbookListener

    /// <summary>
    /// The method consolidates the same conditional formats, which are located in adjacent ranges.
    /// </summary>
    internal void Consolidate()
    {
        var formats = _conditionalFormats
            .Where(cf => cf.Ranges.Count > 0)
            .ToList();
        _conditionalFormats.Clear();

        while (formats.Count > 0)
        {
            var item = formats[0];

            if (!CFTypesExcludedFromConsolidation.Contains(item.ConditionalFormatType))
            {
                var similarFormats = ConsolidateItem(item, formats);
                formats.RemoveAll(similarFormats.Contains);
            }

            _conditionalFormats.Add(item);
            formats.Remove(item);
        }
    }

    private static List<IXLConditionalFormat> ConsolidateItem(IXLConditionalFormat item,
        List<IXLConditionalFormat> formats)
    {
        var rangesToJoin = new XLRanges();
        item.Ranges.ForEach(rangesToJoin.Add);
        var firstRange = item.Ranges.First();
        var skippedRanges = new XLRanges();

        var baseAddress = new XLAddress(
            item.Ranges.Select(r => r.RangeAddress.FirstAddress.RowNumber).Min(),
            item.Ranges.Select(r => r.RangeAddress.FirstAddress.ColumnNumber).Min(),
            false, false);
        var baseCell = (XLCell)firstRange.Worksheet.Cell(baseAddress);

        var similarFormats = FindSimilarFormats(formats, rangesToJoin, skippedRanges, IsSameFormat);

        var consAreas = XLAreaList.FromRanges(rangesToJoin).GetConsolidated();
        ((XLConditionalFormat)item).SetAreas(consAreas);

        var targetCell = (XLCell)item.Ranges.First().FirstCell();
        ((XLConditionalFormat)item).AdjustFormulas(baseCell, targetCell, leaveRefusedUnchanged: true);

        return similarFormats;

        bool IsSameFormat(IXLConditionalFormat f) => f != item &&
                                                     f.Ranges.First().Worksheet.Position ==
                                                     firstRange.Worksheet.Position &&
                                                     XLConditionalFormat.NoRangeComparer.Equals(f, item);
    }

    private static List<IXLConditionalFormat> FindSimilarFormats(
        List<IXLConditionalFormat> formats,
        XLRanges rangesToJoin,
        XLRanges skippedRanges,
        Func<IXLConditionalFormat, bool> isSameFormat)
    {
        List<IXLConditionalFormat> similarFormats = [];
        var i = 1;
        bool stop;
        do
        {
            stop = i >= formats.Count;

            if (!stop)
            {
                var nextFormat = formats[i];

                var intersectsSkipped =
                    skippedRanges.Any(left => nextFormat.Ranges.GetIntersectedRanges(left.RangeAddress).Any());

                var isSame = isSameFormat(nextFormat);

                if (isSame && !intersectsSkipped)
                {
                    similarFormats.Add(nextFormat);
                    nextFormat.Ranges.ForEach(rangesToJoin.Add);
                }
                else if (rangesToJoin.Any(left => nextFormat.Ranges.GetIntersectedRanges(left.RangeAddress).Any()) ||
                         intersectsSkipped)
                {
                    stop = true;
                }

                if (!isSame)
                    nextFormat.Ranges.ForEach(skippedRanges.Add);
            }

            i++;
        } while (!stop);

        return similarFormats;
    }

    public void RemoveAll()
    {
        _conditionalFormats.Clear();
    }

    /// <summary>
    /// Reorders the conditional formats according to the original priority. Done during the load process.
    /// </summary>
    public void ReorderAccordingToOriginalPriority()
    {
        var reorderedFormats = _conditionalFormats.OrderBy(cf => ((XLConditionalFormat)cf).Priority).ToList();
        _conditionalFormats.Clear();
        _conditionalFormats.AddRange(reorderedFormats);
    }
}
