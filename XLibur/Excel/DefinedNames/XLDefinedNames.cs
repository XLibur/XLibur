using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using XLibur.Excel.Coordinates;
using XLibur.Extensions;

namespace XLibur.Excel;

/// <summary>
/// A collection of a named ranges, either for workbook or for worksheet.
/// </summary>
internal sealed class XLDefinedNames : IXLDefinedNames, IEnumerable<XLDefinedName>, ISheetListener
{
    private readonly Dictionary<string, XLDefinedName> _namedRanges = new(XLHelper.NameComparer);

    internal XLWorkbook Workbook { get; set; }

    internal XLWorksheet? Worksheet { get; set; }

    internal XLNamedRangeScope Scope { get; }

    public XLDefinedNames(XLWorksheet worksheet)
        : this(worksheet.Workbook)
    {
        Worksheet = worksheet;
        Scope = XLNamedRangeScope.Worksheet;
    }

    public XLDefinedNames(XLWorkbook workbook)
    {
        Workbook = workbook;
        Scope = XLNamedRangeScope.Workbook;
    }

    #region ISheetListener

    /// <summary>
    /// Re-points every name in this collection at the edited range's new addresses.
    /// </summary>
    /// <remarks>
    /// <b>Deliberately unguarded.</b> Unlike <see cref="XLHyperlinks"/> this does not compare
    /// <c>edit.Sheet</c> with its own sheet, because a name defined on one sheet may refer to
    /// another, and every collection in the workbook — each sheet's and the workbook's own — is
    /// yielded for every edit. The filtering happens one level down:
    /// <see cref="XLCellFormulaShifter"/> re-points only references to the sheet that was edited and
    /// leaves references to any other sheet alone, which is the same reasoning that let the
    /// hardcoded pass this replaced visit every worksheet.
    /// <para>
    /// A reference that shifts to nothing is dropped from the list rather than left as an error, so
    /// a name whose every reference is deleted ends up referring to an empty string.
    /// </para>
    /// </remarks>
    void ISheetListener.OnInsertAreaAndShiftDown(in SheetEdit edit) => MoveDefinedNames<RowAxis>(in edit);

    void ISheetListener.OnInsertAreaAndShiftRight(in SheetEdit edit) => MoveDefinedNames<ColumnAxis>(in edit);

    void ISheetListener.OnDeleteAreaAndShiftUp(in SheetEdit edit) => MoveDefinedNames<RowAxis>(in edit);

    void ISheetListener.OnDeleteAreaAndShiftLeft(in SheetEdit edit) => MoveDefinedNames<ColumnAxis>(in edit);

    private void MoveDefinedNames<TAxis>(in SheetEdit edit)
        where TAxis : struct, IGridAxis
    {
        var axis = default(TAxis);

        // The edit arrives as an `in` parameter, which the Select lambda below cannot capture.
        var sheet = edit.Sheet;
        var range = edit.Range;
        var shift = edit.Shift;

        foreach (var definedName in this)
        {
            var sheetRefs = definedName.GetSheetReferencesList();
            if (sheetRefs.Count == 0)
                continue;
            var newRangeList = sheetRefs
                .Select(r => axis.ShiftFormula(r, sheet, range, shift))
                .Where(newReference => newReference.Length > 0)
                .ToList();
            definedName.SetRefersTo(string.Join(",", newRangeList));
        }
    }

    #endregion ISheetListener

    #region IXLNamedRanges Members

    [Obsolete("Use DefinedName instead.")]
    IXLDefinedName IXLDefinedNames.NamedRange(string name) => DefinedName(name);

    IXLDefinedName IXLDefinedNames.DefinedName(string name) => DefinedName(name);

    internal XLDefinedName DefinedName(string name)
    {
        if (_namedRanges.TryGetValue(name, out XLDefinedName? range))
            return range;

        throw new KeyNotFoundException($"Name {name} not found.");
    }

    public IXLDefinedName Add(string name, string rangeAddress)
    {
        return Add(name, rangeAddress, null);
    }

    public IXLDefinedName Add(string name, IXLRange range)
    {
        return Add(name, range, null);
    }

    public IXLDefinedName Add(string name, IXLRanges ranges)
    {
        return Add(name, ranges, null);
    }

    public IXLDefinedName Add(string name, string rangeAddress, string? comment)
    {
        return Add(name, rangeAddress, comment, validateName: true, validateRangeAddress: true);
    }

    /// <summary>
    /// Adds the specified range name.
    /// </summary>
    /// <param name="name">Name of the range.</param>
    /// <param name="rangeAddress">The range address.</param>
    /// <param name="comment">The comment.</param>
    /// <param name="validateName">if set to <c>true</c> validates the name.</param>
    /// <param name="validateRangeAddress">if set to <c>true</c> range address will be checked for validity.</param>
    /// <exception cref="NotSupportedException"></exception>
    /// <exception cref="ArgumentException">
    /// For named ranges in the workbook scope, specify the sheet name in the reference.
    /// </exception>
    internal IXLDefinedName Add(string name, string rangeAddress, string? comment, bool validateName, bool validateRangeAddress)
    {
        if (validateRangeAddress)
            rangeAddress = ValidateAndResolveAddress(name, rangeAddress);

        var namedRange = new XLDefinedName(this, name, validateName, rangeAddress, comment);
        _namedRanges.Add(name, namedRange);
        return namedRange;
    }

    public IXLDefinedName Add(string name, IXLRange range, string? comment)
    {
        var ranges = new XLRanges { range };
        return Add(name, ranges, comment);
    }

    public IXLDefinedName Add(string name, IXLRanges ranges, string? comment)
    {
        var formula = string.Join(",", ranges.Select(r => r.RangeAddress.ToStringFixed(XLReferenceStyle.A1, true)));
        var namedRange = new XLDefinedName(this, name, true, formula, comment);
        _namedRanges.Add(name, namedRange);
        return namedRange;
    }

    internal XLDefinedName Add(string name, XLDefinedName namedRange)
    {
        _namedRanges.Add(name, namedRange);
        return namedRange;
    }

    private string ValidateAndResolveAddress(string name, string rangeAddress)
    {
        var match = XLHelper.NamedRangeReferenceRegex.Match(rangeAddress);
        if (match.Success)
            return rangeAddress;

        if (!XLHelper.IsValidRangeAddress(rangeAddress))
            return rangeAddress;

        var range = ResolveRange(rangeAddress);

        if (range == null)
            throw new ArgumentException(string.Format(
                "The range address '{0}' for the named range '{1}' is not a valid range.", rangeAddress,
                name));

        if (Scope == XLNamedRangeScope.Workbook && !XLHelper.NamedRangeReferenceRegex.IsMatch(range.ToString()!))
            throw new ArgumentException(
                "For named ranges in the workbook scope, specify the sheet name in the reference.");

        return range.ToString()!;
    }

    private IXLRange? ResolveRange(string rangeAddress)
    {
        if (Scope == XLNamedRangeScope.Worksheet)
            return Worksheet!.Range(rangeAddress);

        if (Scope == XLNamedRangeScope.Workbook)
            return Workbook.Range(rangeAddress);

        throw new NotSupportedException($"Scope {Scope} is not supported");
    }

    public void Delete(string name)
    {
        _namedRanges.Remove(name);
    }

    public void Delete(int index)
    {
        _namedRanges.Remove(_namedRanges.ElementAt(index).Key);
    }

    public void DeleteAll()
    {
        _namedRanges.Clear();
    }

    /// <summary>
    /// Returns a subset of named ranges that do not have invalid references.
    /// </summary>
    public IEnumerable<IXLDefinedName> ValidNamedRanges()
    {
        return _namedRanges.Values.Where(nr => nr.IsValid);
    }

    /// <summary>
    /// Returns a subset of named ranges that do have invalid references.
    /// </summary>
    public IEnumerable<IXLDefinedName> InvalidNamedRanges()
    {
        return _namedRanges.Values.Where(nr => !nr.IsValid);
    }

    #endregion IXLNamedRanges Members

    IEnumerator<XLDefinedName> IEnumerable<XLDefinedName>.GetEnumerator() => GetEnumerator();

    IEnumerator<IXLDefinedName> IEnumerable<IXLDefinedName>.GetEnumerator() => GetEnumerator();

    public Dictionary<string, XLDefinedName>.ValueCollection.Enumerator GetEnumerator()
    {
        return _namedRanges.Values.GetEnumerator();
    }

    #region IEnumerable Members

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion IEnumerable Members

    public bool TryGetValue(string name, [NotNullWhen(true)] out IXLDefinedName? range)
    {
        if (TryGetScopedValue(name, out var sheetDefinedName))
        {
            range = sheetDefinedName;
            return true;
        }

        range = Scope == XLNamedRangeScope.Workbook
            ? Workbook.DefinedName(name)
            : null;

        return range is not null;
    }

    internal bool TryGetScopedValue(string name, [NotNullWhen(true)] out XLDefinedName? definedName)
    {
        if (_namedRanges.TryGetValue(name, out definedName))
        {
            return true;
        }

        return false;
    }

    public bool Contains(string name)
    {
        if (_namedRanges.ContainsKey(name)) return true;

        if (Scope == XLNamedRangeScope.Workbook)
            return Workbook.DefinedName(name) is not null;
        return false;
    }

    internal void OnWorksheetDeleted(string worksheetName)
    {
        _namedRanges.Values
            .ForEach(nr => nr.OnWorksheetDeleted(worksheetName));
    }
}
