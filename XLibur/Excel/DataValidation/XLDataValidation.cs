using System;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel.Coordinates;
using XLibur.Extensions;

namespace XLibur.Excel;

internal sealed class RangeEventArgs : EventArgs
{
    public RangeEventArgs(IXLRange range)
    {
        Range = range ?? throw new ArgumentNullException(nameof(range));
    }

    public IXLRange Range { get; }
}

internal sealed class XLDataValidation : IXLDataValidation
{
    private readonly XLWorksheet _worksheet;

    public XLDataValidation(IXLRange range)
        : this((XLWorksheet)(range ?? throw new ArgumentNullException(nameof(range))).Worksheet)
    {
        AddRange(range);
    }

    public XLDataValidation(IXLDataValidation dataValidation, XLWorksheet worksheet)
        : this(worksheet)
    {
        _worksheet = worksheet;
        CopyFrom(dataValidation);
    }

    private XLDataValidation(XLWorksheet worksheet)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        Areas = XLAreaList.Empty;
        Initialize();
    }

    /// <summary>
    /// Raised when a single range is appended to the coverage. The container uses it to split any
    /// pre-existing overlapping rule (one validation per cell) and add the matching index entry.
    /// </summary>
    internal event EventHandler<RangeEventArgs>? RangeAdded;

    /// <summary>
    /// Raised when a single range is removed from the coverage. The container uses it to drop the
    /// matching index entry.
    /// </summary>
    internal event EventHandler<RangeEventArgs>? RangeRemoved;

    /// <summary>
    /// Raised when the whole coverage is replaced in one step (a structural shift or consolidation
    /// writing back a value-typed area transform via <see cref="SetAreas"/>). The container reindexes
    /// this rule from its <see cref="Areas"/>. Unlike <see cref="RangeAdded"/> it performs no
    /// split-on-add: a shift preserves the disjointness the rules already had.
    /// </summary>
    internal event EventHandler? CoverageReplaced;

    /// <summary>
    /// The rule's coverage, as a value-typed <see cref="XLAreaList"/>. This is the source of truth:
    /// coverage lives here rather than as live repository ranges, so structural (row/column insert
    /// &amp; delete) shifts run as pure area transforms and can never alias or double-shift (ClosedXML
    /// issue #2850). <see cref="Ranges"/> is a projection of it. Mirrors
    /// <see cref="XLibur.Excel.ConditionalFormats.XLConditionalFormat.Areas"/>.
    /// </summary>
    internal XLAreaList Areas { get; private set; }

    internal XLWorksheet Worksheet => _worksheet;

    public void Clear()
    {
        Initialize();
    }

    public void CopyFrom(IXLDataValidation dataValidation)
    {
        if (dataValidation == this) return;

        if (Areas.Count == 0)
            AddRanges(dataValidation.Ranges);

        IgnoreBlanks = dataValidation.IgnoreBlanks;
        InCellDropdown = dataValidation.InCellDropdown;
        ShowErrorMessage = dataValidation.ShowErrorMessage;
        ShowInputMessage = dataValidation.ShowInputMessage;
        InputTitle = dataValidation.InputTitle;
        InputMessage = dataValidation.InputMessage;
        ErrorTitle = dataValidation.ErrorTitle;
        ErrorMessage = dataValidation.ErrorMessage;
        ErrorStyle = dataValidation.ErrorStyle;
        AllowedValues = dataValidation.AllowedValues;
        Operator = dataValidation.Operator;
        MinValue = dataValidation.MinValue;
        MaxValue = dataValidation.MaxValue;
    }

    public bool IsDirty()
    {
        return
            AllowedValues != XLAllowedValues.AnyValue
            || (ShowInputMessage &&
                (!string.IsNullOrWhiteSpace(InputTitle) || !string.IsNullOrWhiteSpace(InputMessage)))
            || (ShowErrorMessage &&
                (!string.IsNullOrWhiteSpace(ErrorTitle) || !string.IsNullOrWhiteSpace(ErrorMessage)));
    }

    /// <summary>
    /// Carve <paramref name="rangeAddress"/> out of the coverage, keeping the non-overlapping
    /// remainder. Each intersecting area is replaced by its remainder pieces one at a time (via
    /// <see cref="RemoveRange"/> / <see cref="AddRange"/>) so the container maintains its index
    /// incrementally — the same granular flow the range-based implementation used, which keeps the
    /// index enumeration order (and therefore copy order) stable. Remainder pieces are emitted in
    /// reading order (top-to-bottom, then left-to-right).
    /// </summary>
    internal void SplitBy(IXLRangeAddress rangeAddress)
    {
        var excludedArea = XLSheetRange.FromRangeAddress(rangeAddress);

        foreach (var area in Areas.IntersectingWith(excludedArea).ToList())
        {
            var pieces = new List<XLSheetRange>();
            area.Exclude(excludedArea, pieces);
            pieces.Sort((a, b) => a.TopRow != b.TopRow ? a.TopRow - b.TopRow : a.LeftColumn - b.LeftColumn);

            RemoveRange(MaterializeRange(area));
            foreach (var piece in pieces)
                AddRange(MaterializeRange(piece));
        }
    }

    private void Initialize()
    {
        AllowedValues = XLAllowedValues.AnyValue;
        IgnoreBlanks = true;
        ShowErrorMessage = true;
        ShowInputMessage = true;
        InCellDropdown = true;
        InputTitle = string.Empty;
        InputMessage = string.Empty;
        ErrorTitle = string.Empty;
        ErrorMessage = string.Empty;
        ErrorStyle = XLErrorStyle.Stop;
        Operator = XLOperator.Between;
        Value = string.Empty;
        MinValue = string.Empty;
        MaxValue = string.Empty;
    }

    #region IXLDataValidation Members

    private string maxValue = string.Empty;
    private string minValue = string.Empty;
    public XLAllowedValues AllowedValues { get; set; }

    public XLDateCriteria Date
    {
        get
        {
            AllowedValues = XLAllowedValues.Date;
            return new XLDateCriteria(this);
        }
    }

    public XLDecimalCriteria Decimal
    {
        get
        {
            AllowedValues = XLAllowedValues.Decimal;
            return new XLDecimalCriteria(this);
        }
    }

    public string ErrorMessage { get; set; } = string.Empty;
    public XLErrorStyle ErrorStyle { get; set; }
    public string ErrorTitle { get; set; } = string.Empty;
    public bool IgnoreBlanks { get; set; }
    public bool InCellDropdown { get; set; }
    public string InputMessage { get; set; } = string.Empty;
    public string InputTitle { get; set; } = string.Empty;
    public string MaxValue { get => maxValue; set => maxValue = value; }
    public string MinValue { get => minValue; set => minValue = value; }
    public XLOperator Operator { get; set; }

    /// <summary>
    /// Coverage materialized as ranges on the owning worksheet, in reading order (top-to-bottom,
    /// then left-to-right) — the same ordering the range-backed <see cref="XLRanges"/> collection
    /// enumerated in, which callers (the sqref writer, copy) depend on. A fresh snapshot each call —
    /// mutating a returned range has no effect on the rule; change coverage via <see cref="AddRange"/>,
    /// <see cref="RemoveRange"/>, <see cref="ClearRanges"/>, or <see cref="SetAreas"/>. Projection of
    /// <see cref="Areas"/>, the source of truth.
    /// </summary>
    public IEnumerable<IXLRange> Ranges =>
        Areas.OrderBy(a => a.TopRow).ThenBy(a => a.LeftColumn).Select(MaterializeRange);

    private XLRange MaterializeRange(XLSheetRange area)
        => _worksheet.Range(area.TopRow, area.LeftColumn, area.BottomRow, area.RightColumn);

    /// <summary>
    /// Replace the coverage in one step. Used by the range shifter and consolidation to write back a
    /// value-typed area transform. Signals <see cref="CoverageReplaced"/> so the container reindexes.
    /// </summary>
    internal void SetAreas(XLAreaList areas)
    {
        Areas = areas;
        CoverageReplaced?.Invoke(this, EventArgs.Empty);
    }

    public bool ShowErrorMessage { get; set; }

    public bool ShowInputMessage { get; set; }

    public XLTextLengthCriteria TextLength
    {
        get
        {
            AllowedValues = XLAllowedValues.TextLength;
            return new XLTextLengthCriteria(this);
        }
    }

    public XLTimeCriteria Time
    {
        get
        {
            AllowedValues = XLAllowedValues.Time;
            return new XLTimeCriteria(this);
        }
    }

    public string Value
    {
        get { return MinValue; }
        set { MinValue = value; }
    }

    public XLWholeNumberCriteria WholeNumber
    {
        get
        {
            AllowedValues = XLAllowedValues.WholeNumber;
            return new XLWholeNumberCriteria(this);
        }
    }

    /// <summary>
    /// Add a range to the collection of ranges this rule applies to.
    /// If the specified range does not belong to the worksheet of the data validation
    /// rule it is transferred to the target worksheet.
    /// </summary>
    /// <param name="range">A range to add.</param>
    public void AddRange(IXLRange range)
    {
        ArgumentNullException.ThrowIfNull(range);

        if (range.Worksheet != Worksheet)
            range = Worksheet.Range(((XLRangeAddress)range.RangeAddress).WithoutWorksheet());

        Areas = Areas.With(XLSheetRange.FromRangeAddress(range.RangeAddress));

        RangeAdded?.Invoke(this, new RangeEventArgs(range));
    }

    /// <summary>
    /// Add a collection of ranges to the collection of ranges this rule applies to.
    /// Ranges that do not belong to the worksheet of the data validation
    /// rule are transferred to the target worksheet.
    /// </summary>
    /// <param name="ranges">Ranges to add.</param>
    public void AddRanges(IEnumerable<IXLRange> ranges)
    {
        ranges ??= [];

        foreach (var range in ranges)
        {
            AddRange(range);
        }
    }

    /// <summary>
    /// Detach data validation rule of all ranges it applies to.
    /// </summary>
    public void ClearRanges()
    {
        var removedRanges = Ranges.ToList();
        Areas = XLAreaList.Empty;

        foreach (var range in removedRanges)
        {
            RangeRemoved?.Invoke(this, new RangeEventArgs(range));
        }
    }

    public void Custom(string customValidation)
    {
        AllowedValues = XLAllowedValues.Custom;
        Value = customValidation;
    }

    public void List(string list)
    {
        List(list, true);
    }

    public void List(string list, bool inCellDropdown)
    {
        AllowedValues = XLAllowedValues.List;
        InCellDropdown = inCellDropdown;
        Value = QuoteListValueIfNeeded(list);
    }

    private string QuoteListValueIfNeeded(string list)
    {
        if (list.Length == 0 || list[0] == '=' || list[0] == '"')
            return list;

        if (XLHelper.IsValidRangeAddress(list))
            return list;

        if (_worksheet.DefinedNames.Contains(list) ||
            _worksheet.Workbook.DefinedNames.Contains(list))
            return list;

        return "\"" + list + "\"";
    }

    public void List(IXLRange range)
    {
        List(range, true);
    }

    public void List(IXLRange range, bool inCellDropdown)
    {
        List(range.RangeAddress.ToStringFixed(XLReferenceStyle.A1, true));
    }

    /// <summary>
    /// Remove the specified range from the collection of range this rule applies to.
    /// </summary>
    /// <param name="range">A range to remove.</param>
    public bool RemoveRange(IXLRange range)
    {
        if (range == null)
            return false;

        var area = XLSheetRange.FromRangeAddress(range.RangeAddress);
        var newAreas = Areas.Without(area);
        if (newAreas.Count == Areas.Count)
            return false;

        Areas = newAreas;
        RangeRemoved?.Invoke(this, new RangeEventArgs(range));
        return true;
    }

    #endregion IXLDataValidation Members

}
