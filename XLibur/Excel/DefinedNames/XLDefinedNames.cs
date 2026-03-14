using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace XLibur.Excel;

/// <summary>
/// A collection of a named ranges, either for workbook or for worksheet.
/// </summary>
internal class XLDefinedNames : IXLDefinedNames, IEnumerable<XLDefinedName>
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

    #region IXLNamedRanges Members

    [Obsolete]
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
        // When loading named ranges from an existing file, we do not validate the range address or name.
        if (validateRangeAddress)
        {
            var match = XLHelper.NamedRangeReferenceRegex.Match(rangeAddress);

            if (!match.Success)
            {
                if (XLHelper.IsValidRangeAddress(rangeAddress))
                {
                    IXLRange? range;
                    if (Scope == XLNamedRangeScope.Worksheet)
                        range = Worksheet!.Range(rangeAddress);
                    else if (Scope == XLNamedRangeScope.Workbook)
                        range = Workbook.Range(rangeAddress);
                    else
                        throw new NotSupportedException($"Scope {Scope} is not supported");

                    if (range == null)
                        throw new ArgumentException(string.Format(
                            "The range address '{0}' for the named range '{1}' is not a valid range.", rangeAddress,
                            name));

                    if (Scope == XLNamedRangeScope.Workbook || !XLHelper.NamedRangeReferenceRegex.Match(range.ToString()!).Success)
                        throw new ArgumentException(
                            "For named ranges in the workbook scope, specify the sheet name in the reference.");

                    rangeAddress = range.ToString()!;
                }
            }
        }

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

    public void Delete(string rangeName)
    {
        _namedRanges.Remove(rangeName);
    }

    public void Delete(int rangeIndex)
    {
        _namedRanges.Remove(_namedRanges.ElementAt(rangeIndex).Key);
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

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion IEnumerable Members

    public bool TryGetValue(string name, [NotNullWhen(true)] out IXLDefinedName? definedName)
    {
        if (TryGetScopedValue(name, out var sheetDefinedName))
        {
            definedName = sheetDefinedName;
            return true;
        }

        definedName = Scope == XLNamedRangeScope.Workbook
            ? Workbook.DefinedName(name)
            : null;

        return definedName is not null;
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
