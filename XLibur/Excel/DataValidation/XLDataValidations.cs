using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel.Coordinates;
using XLibur.Excel.Ranges.Index;
using XLibur.Extensions;

namespace XLibur.Excel;

internal sealed class XLDataValidations : IXLDataValidations, ISheetListener
{
    private readonly XLRangeIndex<XLDataValidationIndexEntry> _dataValidationIndex;

    private readonly List<IXLDataValidation> _dataValidations = new List<IXLDataValidation>();
    private readonly XLWorksheet _worksheet;

    /// <summary>
    /// The flag used to avoid unnecessary check for splitting intersected ranges when we already
    /// are performing the splitting.
    /// </summary>
    private bool _skipSplittingExistingRanges;

    public XLDataValidations(XLWorksheet worksheet)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        _dataValidationIndex = new XLRangeIndex<XLDataValidationIndexEntry>(_worksheet);
    }

    internal XLWorksheet Worksheet => _worksheet;

    #region IXLDataValidations Members

    IXLWorksheet IXLDataValidations.Worksheet => _worksheet;

    public IXLDataValidation Add(IXLDataValidation dataValidation)
    {
        return Add(dataValidation, skipIntersectionsCheck: false);
    }

    public bool ContainsSingle(IXLRange range)
    {
        int count = 0;
        foreach (var xlDataValidation in _dataValidations.Where(dv => dv.Ranges.Contains(range)))
        {
            count++;
            if (count > 1) return false;
        }

        return count == 1;
    }

    public void Delete(Predicate<IXLDataValidation> predicate)
    {
        var dataValidationsToRemove = _dataValidations.Where(dv => predicate(dv))
            .ToList();

        dataValidationsToRemove.ForEach(Delete);
    }

    public void Delete(IXLDataValidation dataValidation)
    {
        if (!_dataValidations.Remove(dataValidation))
            return;
        var xlDataValidation = (XLDataValidation)dataValidation;
        xlDataValidation.RangeAdded -= OnRangeAdded;
        xlDataValidation.RangeRemoved -= OnRangeRemoved;
        xlDataValidation.CoverageReplaced -= OnCoverageReplaced;

        _dataValidationIndex.RemoveAll(e => ReferenceEquals(e.DataValidation, xlDataValidation));
    }

    public void Delete(IXLRange range)
    {
        ArgumentNullException.ThrowIfNull(range);

        var dataValidationsToRemove = _dataValidationIndex.GetIntersectedRanges((XLRangeAddress)range.RangeAddress)
            .Select(e => e.DataValidation)
            .Distinct()
            .ToList();

        dataValidationsToRemove.ForEach(Delete);
    }

    /// <summary>
    /// Get all data validation rules applied to ranges that intersect the specified range.
    /// </summary>
    public IEnumerable<IXLDataValidation> GetAllInRange(IXLRangeAddress rangeAddress)
    {
        if (rangeAddress == null || !rangeAddress.IsValid)
            return Enumerable.Empty<IXLDataValidation>();

        return _dataValidationIndex.GetIntersectedRanges((XLRangeAddress)rangeAddress)
            .Select(indexEntry => indexEntry.DataValidation)
            .Distinct();
    }

    /// <summary>
    /// Whether any data validation rule covers a cell of <paramref name="rangeAddress"/>. Cheaper than
    /// <see cref="GetAllInRange"/> when only the yes/no answer is wanted: it stops at the first hit and
    /// skips the de-duplication.
    /// </summary>
    internal bool AnyInRange(in XLRangeAddress rangeAddress)
    {
        return rangeAddress.IsValid && _dataValidationIndex.Intersects(in rangeAddress);
    }

    #region ISheetListener

    /// <summary>
    /// Moves this sheet's validation coverage over the edit, then re-points the cell references
    /// inside every rule's criteria formulas.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The two passes are one listener because their order is a requirement, not an accident: the
    /// coverage pass deletes a rule whose coverage transforms to nothing, and the formula pass must
    /// not then rewrite a rule that is gone. Expressing that as "this method does A then B" is
    /// stronger than two positions in <see cref="XLWorksheet.GetSheetListeners"/>, which is why the
    /// enumeration yields one <see cref="XLDataValidations"/> per sheet rather than this sheet's
    /// twice.
    /// </para>
    /// <para>
    /// Their scopes differ. Coverage (<c>sqref</c>) is sheet-scoped and guards on
    /// <c>edit.Sheet</c>. Criteria formulas — <c>formula1</c>/<c>formula2</c>, behind
    /// <see cref="IXLDataValidation.MinValue"/> and <see cref="IXLDataValidation.MaxValue"/> — are
    /// not: a dependent dropdown on another sheet may point through OFFSET or MATCH at the sheet
    /// that was edited, so every sheet's collection is visited and
    /// <see cref="XLCellFormulaShifter"/> does the filtering, exactly as for
    /// <see cref="XLDefinedNames"/>.
    /// </para>
    /// </remarks>
    void ISheetListener.OnInsertAreaAndShiftDown(in SheetEdit edit) => Shift<RowAxis>(in edit);

    void ISheetListener.OnInsertAreaAndShiftRight(in SheetEdit edit) => Shift<ColumnAxis>(in edit);

    void ISheetListener.OnDeleteAreaAndShiftUp(in SheetEdit edit) => Shift<RowAxis>(in edit);

    void ISheetListener.OnDeleteAreaAndShiftLeft(in SheetEdit edit) => Shift<ColumnAxis>(in edit);

    private void Shift<TAxis>(in SheetEdit edit)
        where TAxis : struct, IGridAxis
    {
        var axis = default(TAxis);

        if (edit.Sheet == _worksheet && _dataValidations.Count > 0)
        {
            // CoverageArea, not edit.Area: coverage derives the |Shift| lines the edit moves from
            // the range and the shift rather than trusting the shifter's area. See SheetEdit.
            var affected = edit.CoverageArea<TAxis>();
            foreach (var dv in _dataValidations.OfType<XLDataValidation>().ToList())
            {
                // Coverage is the area model, not live repository ranges, so the transform is pure
                // and can never alias or double-shift (ClosedXML issue #2850), and the write-back
                // reindexes without split-on-add — an extended range can no longer split a
                // not-yet-shifted rule it transiently overlaps (the drop-on-insert bug).
                var newAreas = edit.Shift > 0
                    ? axis.InsertAndShift(dv.Areas, affected)
                    : axis.DeleteAndShift(dv.Areas, affected);

                if (newAreas.Count == 0)
                    Delete(v => v == dv);
                else
                    dv.SetAreas(newAreas);
            }
        }

        foreach (var dv in _dataValidations.ToList())
        {
            if (!string.IsNullOrEmpty(dv.MinValue))
                dv.MinValue = axis.ShiftFormula(dv.MinValue, _worksheet, edit.Range, edit.Shift);
            if (!string.IsNullOrEmpty(dv.MaxValue))
                dv.MaxValue = axis.ShiftFormula(dv.MaxValue, _worksheet, edit.Range, edit.Shift);
        }
    }

    #endregion ISheetListener

    public IEnumerator<IXLDataValidation> GetEnumerator()
    {
        return _dataValidations.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    /// <summary>
    /// Get the data validation rule for the range with the specified address if it exists.
    /// </summary>
    /// <param name="rangeAddress">A range address.</param>
    /// <param name="dataValidation">Data validation rule which ranges collection includes the specified
    /// address. The specified range should be fully covered with the data validation rule.
    /// For example, if the rule is applied to ranges A1:A3,C1:C3 then this method will
    /// return True for ranges A1:A3, C1:C2, A2:A3, and False for ranges A1:C3, A1:C1, etc.</param>
    /// <returns>True is the data validation rule was found, false otherwise.</returns>
    public bool TryGet(IXLRangeAddress rangeAddress, out IXLDataValidation? dataValidation)
    {
        dataValidation = null;
        if (rangeAddress == null || !rangeAddress.IsValid)
            return false;

        var candidates = _dataValidationIndex.GetIntersectedRanges((XLRangeAddress)rangeAddress)
            .Where(c => c.RangeAddress.Contains(rangeAddress.FirstAddress) &&
                        c.RangeAddress.Contains(rangeAddress.LastAddress));

        var candidate = candidates.FirstOrDefault();
        if (candidate is null)
            return false;

        dataValidation = candidate.DataValidation;

        return true;
    }

    internal IXLDataValidation Add(IXLDataValidation dataValidation, bool skipIntersectionsCheck)
    {
        ArgumentNullException.ThrowIfNull(dataValidation);

        XLDataValidation xlDataValidation;
        if (dataValidation is not XLDataValidation validation ||
            dataValidation.Ranges.Any(r => r.Worksheet != Worksheet))
        {
            xlDataValidation = new XLDataValidation(dataValidation, Worksheet);
        }
        else
        {
            xlDataValidation = validation;
        }

        xlDataValidation.RangeAdded += OnRangeAdded;
        xlDataValidation.RangeRemoved += OnRangeRemoved;
        xlDataValidation.CoverageReplaced += OnCoverageReplaced;

        foreach (var range in xlDataValidation.Ranges)
        {
            ProcessRangeAdded(range, xlDataValidation, skipIntersectionsCheck);
        }

        _dataValidations.Add(xlDataValidation);

        return xlDataValidation;
    }

    #endregion IXLDataValidations Members

    public void Consolidate()
    {
        Func<IXLDataValidation, IXLDataValidation, bool> areEqual = (dv1, dv2) =>
        {
            return
                dv1.IgnoreBlanks == dv2.IgnoreBlanks &&
                dv1.InCellDropdown == dv2.InCellDropdown &&
                dv1.ShowErrorMessage == dv2.ShowErrorMessage &&
                dv1.ShowInputMessage == dv2.ShowInputMessage &&
                dv1.InputTitle == dv2.InputTitle &&
                dv1.InputMessage == dv2.InputMessage &&
                dv1.ErrorTitle == dv2.ErrorTitle &&
                dv1.ErrorMessage == dv2.ErrorMessage &&
                dv1.ErrorStyle == dv2.ErrorStyle &&
                dv1.AllowedValues == dv2.AllowedValues &&
                dv1.Operator == dv2.Operator &&
                dv1.MinValue == dv2.MinValue &&
                dv1.MaxValue == dv2.MaxValue &&
                dv1.Value == dv2.Value;
        };

        var rules = _dataValidations.ToList();
        rules.ForEach(Delete);

        while (rules.Count > 0)
        {
            var similarRules = rules.Where(r => areEqual(rules[0], r)).ToList();
            rules.RemoveAll(similarRules.Contains);

            var consRule = (XLDataValidation)similarRules[0];

            // Merge every similar rule's coverage and collapse adjacent/overlapping blocks with the
            // value-typed area model (mirrors XLConditionalFormats.ConsolidateItem). Add() reindexes.
            var mergedAreas = new XLAreaList(
                similarRules.Cast<XLDataValidation>().SelectMany(dv => dv.Areas).ToList());
            consRule.SetAreas(mergedAreas.GetConsolidated());
            Add(consRule);
        }
    }

    private void OnRangeAdded(object? sender, RangeEventArgs e)
    {
        ProcessRangeAdded(e.Range, (XLDataValidation)sender!, skipIntersectionCheck: false);
    }

    private void OnRangeRemoved(object? sender, RangeEventArgs e)
    {
        ProcessRangeRemoved(e.Range);
    }

    private void OnCoverageReplaced(object? sender, EventArgs e)
    {
        ReindexRule((XLDataValidation)sender!);
    }

    private void ProcessRangeAdded(IXLRange range, XLDataValidation dataValidation, bool skipIntersectionCheck)
    {
        if (!skipIntersectionCheck)
        {
            SplitExistingRanges(range.RangeAddress);
        }

        var indexEntry = new XLDataValidationIndexEntry(range.RangeAddress, dataValidation);
        _dataValidationIndex.Add(indexEntry);
    }

    private void ProcessRangeRemoved(IXLRange range)
    {
        var entries = _dataValidationIndex.GetIntersectedRanges((XLRangeAddress)range.RangeAddress)
            .Where(e => Equals(e.RangeAddress, range.RangeAddress));
        entries.ToArray().ForEach(entry => _dataValidationIndex.Remove(entry.RangeAddress));
    }

    /// <summary>
    /// Rebuild a single rule's spatial-index entries from its <see cref="XLDataValidation.Areas"/>.
    /// Called when a rule replaces its whole coverage in one step (a structural shift or
    /// consolidation via <see cref="XLDataValidation.SetAreas"/>, signalled by
    /// <see cref="XLDataValidation.CoverageReplaced"/>). Entries are keyed off immutable area values
    /// (never a live repository range), so a subsequent blanket range shift cannot desync them —
    /// which is what let the old defensive full-index rebuild (ReconcileIndex) go.
    /// </summary>
    private void ReindexRule(XLDataValidation dataValidation)
    {
        _dataValidationIndex.RemoveAll(e => ReferenceEquals(e.DataValidation, dataValidation));
        foreach (var area in dataValidation.Areas)
        {
            var address = XLRangeAddress.FromSheetRange(_worksheet, area);
            _dataValidationIndex.Add(new XLDataValidationIndexEntry(address, dataValidation));
        }
    }

    private void SplitExistingRanges(IXLRangeAddress rangeAddress)
    {
        if (_skipSplittingExistingRanges) return;

        // Distinct because the index holds one entry per area, so a rule covering several
        // intersecting areas would otherwise be split once per area. SplitBy is idempotent —
        // later calls find nothing left intersecting — but there is no reason to repeat it,
        // and the emptied-rule sweep below needs each rule considered once.
        var split = _dataValidationIndex.GetIntersectedRanges((XLRangeAddress)rangeAddress)
            .Select(entry => entry.DataValidation)
            .Distinct()
            .ToList();

        try
        {
            _skipSplittingExistingRanges = true;

            foreach (var dataValidation in split)
            {
                dataValidation.SplitBy(rangeAddress);
            }
        }
        finally
        {
            _skipSplittingExistingRanges = false;
        }

        // A rule whose coverage lay entirely inside rangeAddress has no areas left: the new
        // rule took every cell it applied to. Excel cannot express a rule that applies to
        // nothing — sqref must be non-empty — and the rule is now unreachable, so drop it
        // rather than leave it in the collection to be written as sqref="".
        foreach (var dataValidation in split)
        {
            if (dataValidation.Areas.Count == 0)
                Delete(dataValidation);
        }
    }

    /// <summary>
    /// Class used for indexing data validation rules.
    /// </summary>
    private sealed class XLDataValidationIndexEntry : IXLAddressable
    {
        public XLDataValidationIndexEntry(IXLRangeAddress rangeAddress, XLDataValidation dataValidation)
        {
            RangeAddress = rangeAddress;
            DataValidation = dataValidation;
        }

        public XLDataValidation DataValidation { get; }

        /// <summary>
        ///   Gets an object with the boundaries of this range.
        /// </summary>
        public IXLRangeAddress RangeAddress { get; }
    }
}
