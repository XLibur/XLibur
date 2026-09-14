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
        var affected = edit.CoverageArea<TAxis>();
        foreach (var cf in _conditionalFormats.OfType<XLConditionalFormat>().ToList())
        {
            var pieces = CutIntoPieces<TAxis>(in edit, affected, cf.Areas);
            if (pieces.Count == 0)
            {
                Remove(f => f == cf);
                continue;
            }

            // Each piece after the first becomes a rule of its own, next to the rule it came from, with
            // its priority, type, values and style. Each piece's formulas are rebased onto its own
            // origin before the formula pass shifts them; see CutIntoPieces.
            var anchor = XLConditionalFormat.AnchorOf(cf.Areas);
            var index = _conditionalFormats.IndexOf(cf);
            for (var i = pieces.Count - 1; i > 0; i--)
            {
                var piece = cf.CopyOnto(pieces[i].Areas);
                piece.RebaseFormulas(anchor, pieces[i].Origin);
                piece.IsSplitByEdit = true;
                _conditionalFormats.Insert(index + 1, piece);
            }

            cf.RebaseFormulas(anchor, pieces[0].Origin);
            cf.SetAreas(pieces[0].Areas);
            if (pieces.Count > 1)
                cf.IsSplitByEdit = true;
        }

        // A kept x14 rule's range goes through the same transform, so it cannot part company with a
        // modelled rule over the same cells (issue #499). An emptied range stays, empty, so that the
        // writer knows to remove the rule.
        foreach (var (ruleId, areas) in _extensionRuleAreas.ToList())
        {
            if (areas.Count == 0)
                continue;

            var pieces = CutIntoPieces<TAxis>(in edit, affected, areas);
            var newAreas = new XLAreaList(pieces.SelectMany(p => p.Areas).ToList());

            // KNOWN GAP: a kept rule the edit cuts into two pieces is not split, because splitting it
            // means duplicating its XML with a new rule id per piece (issue #499). It stays one rule over
            // both pieces, with its formulas rebased as for a rule the edit does not cut: onto where its
            // new anchor stood. That is right for the piece that holds the anchor, and anchored wrongly
            // for the other. Pinned by Known_gap_a_kept_rule_an_edit_cuts_stays_one_rule.
            if (newAreas.Count > 0 && _extensionRuleFormulas.TryGetValue(ruleId, out var formulas))
            {
                var anchor = XLConditionalFormat.AnchorOf(areas);
                var origin = OriginOf<TAxis>(in edit, affected, XLConditionalFormat.AnchorOf(newAreas));
                for (var i = 0; i < formulas.Length; i++)
                {
                    if (XLConditionalFormat.TryRebaseFormula(formulas[i], anchor, origin, out var rebased))
                        formulas[i] = rebased;
                }
            }

            _extensionRuleAreas[ruleId] = newAreas;
        }
    }

    /// <summary>
    /// The pieces a row or column edit on this sheet cuts a rule's range into, each with its origin:
    /// the piece's anchor (<see cref="XLConditionalFormat.AnchorOf"/>) as it stood before the edit.
    /// Empty when the edit removes every cell of the range.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A rule's formulas are written relative to its range's anchor. For each piece
    /// Excel moves them onto the piece's origin, keeping each relative reference's offset from the
    /// anchor, and then shifts them as a cell formula there (<c>cf-anchor-*.xlsx</c>,
    /// <c>cf-partial-*.xlsx</c>). Deleting row 2 under <c>$A2&gt;5</c> on <c>A2:C10</c> leaves one
    /// piece, <c>A2:C9</c>, whose origin is the old <c>A3</c>: the formula is rebased to
    /// <c>$A3&gt;5</c> and shifted back to <c>$A2&gt;5</c>, where shifting alone gives <c>#REF!&gt;5</c>.
    /// </para>
    /// <para>
    /// An insert or delete of cells across part of the sheet cuts a rule when it leaves some of the
    /// rule's cells moved and others where they were, within one area or across areas. The rule is then
    /// two pieces: the cells that did not move, and the cells that did. Deleting <c>A2</c> with a shift
    /// left cuts <c>A2:C10</c> into <c>A3:C10</c>, whose origin is <c>A3</c>, and <c>A2:B2</c>, whose
    /// origin is the old <c>B2</c>. Deleting <c>A1</c> with a shift up cuts no single area of
    /// <c>A2:A10 C1:C10</c>, but cuts the rule into <c>A1:A9</c> and <c>C1:C10</c>, so that <c>C1</c>
    /// keeps reading what it read.
    /// </para>
    /// <para>
    /// A whole-row or whole-column edit leaves one piece, over every area of the rule, as Excel does for
    /// the <c>cf-anchor-*.xlsx</c> ranges and as <c>ConditionalFormatRangeShiftTests</c> pins for two
    /// areas. If any of the rule's cells stays, its anchor lies before the edit's lines and stays too,
    /// so the formula shifts as it is; no cell's formula is moved onto another's. So does a range that
    /// cannot move: <c>$A1&gt;5</c> on the whole column <c>E:E</c> only shifts, to <c>$A2&gt;5</c>
    /// after a row is inserted at 1.
    /// </para>
    /// </remarks>
    /// <param name="edit">The edit, on this sheet.</param>
    /// <param name="affected">The region the edit inserts or deletes, as coverage sees it.</param>
    /// <param name="areas">The rule's range before the edit.</param>
    private static List<(XLAreaList Areas, Point Origin)> CutIntoPieces<TAxis>(in SheetEdit edit, Area affected,
        XLAreaList areas)
        where TAxis : struct, IGridAxis
    {
        var axis = default(TAxis);
        var result = edit.Shift > 0 ? axis.InsertAndShift(areas, affected) : axis.DeleteAndShift(areas, affected);
        var parts = new List<(Area Part, bool Moved)>(result.Count);
        foreach (var part in result)
            parts.Add((part, HasMoved<TAxis>(in edit, affected, part.FirstPoint)));

        var pieces = new List<(XLAreaList Areas, Point Origin)>(2);
        if (parts.Count == 0)
            return pieces;

        // Judged over the whole rule, not area by area: one area moving whole while another stays cuts it
        // as surely as an area cut in two. A whole-line edit never cuts it; see the remarks.
        var cut = !axis.IsEntireLine(edit.Range)
                  && parts.Exists(p => p.Moved)
                  && parts.Exists(p => !p.Moved);

        if (!cut)
        {
            var whole = new XLAreaList(parts.Select(p => p.Part).ToList());
            pieces.Add((whole, OriginOf<TAxis>(in edit, affected, XLConditionalFormat.AnchorOf(whole))));
            return pieces;
        }

        // The piece whose first part came first keeps the rule's place.
        var firstMoved = parts[0].Moved;
        foreach (var moved in new[] { firstMoved, !firstMoved })
        {
            var group = new XLAreaList(parts.Where(p => p.Moved == moved).Select(p => p.Part).ToList());
            pieces.Add((group, OriginOf<TAxis>(in edit, affected, XLConditionalFormat.AnchorOf(group))));
        }

        return pieces;
    }

    /// <summary>
    /// Whether the cell at <paramref name="point"/>, after the edit, got there by moving: it is in the
    /// edited lines' cross extent, and past where a delete began or past the lines an insert added.
    /// </summary>
    private static bool HasMoved<TAxis>(in SheetEdit edit, Area affected, Point point)
        where TAxis : struct, IGridAxis
    {
        var axis = default(TAxis);
        var cross = axis.CrossOf(point);
        if (cross < axis.CrossOf(affected.FirstPoint) || cross > axis.CrossOf(affected.LastPoint))
            return false;

        return axis.IndexOf(point) >= axis.IndexOf(affected.FirstPoint) + Math.Max(edit.Shift, 0);
    }

    /// <summary>Where the cell at <paramref name="point"/>, after the edit, stood before it.</summary>
    private static Point OriginOf<TAxis>(in SheetEdit edit, Area affected, Point point)
        where TAxis : struct, IGridAxis
    {
        var axis = default(TAxis);
        return HasMoved<TAxis>(in edit, affected, point)
            ? axis.PointAt(axis.IndexOf(point) - edit.Shift, axis.CrossOf(point))
            : point;
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

            // A piece of a rule an edit cut apart is left as it is (see IsSplitByEdit).
            if (!CFTypesExcludedFromConsolidation.Contains(item.ConditionalFormatType)
                && item is XLConditionalFormat { IsSplitByEdit: false })
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

        // Read from the rule's anchor and re-expressed for the consolidated range's anchor
        // (XLConditionalFormat.AnchorOf). The target used to be the consolidated range's first area's
        // first cell, and consolidation returns areas row by row, so that cell could be right of the
        // anchor: B1:B5 A3:A5 with A1>0 was saved as B1>0, and each later save moved it a column more.
        var baseAnchor = XLConditionalFormat.AnchorOf(((XLConditionalFormat)item).Areas);
        var baseCell = (XLCell)firstRange.Worksheet.Cell(baseAnchor.Row, baseAnchor.Column);

        var similarFormats = FindSimilarFormats(formats, rangesToJoin, skippedRanges, IsSameFormat);

        var consAreas = XLAreaList.FromRanges(rangesToJoin).GetConsolidated();
        ((XLConditionalFormat)item).SetAreas(consAreas);

        var targetAnchor = XLConditionalFormat.AnchorOf(consAreas);
        var targetCell = (XLCell)firstRange.Worksheet.Cell(targetAnchor.Row, targetAnchor.Column);
        ((XLConditionalFormat)item).AdjustFormulas(baseCell, targetCell, leaveRefusedUnchanged: true);

        return similarFormats;

        bool IsSameFormat(IXLConditionalFormat f) => f != item &&
                                                     f is XLConditionalFormat { IsSplitByEdit: false } &&
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
