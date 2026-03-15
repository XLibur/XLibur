using System.Collections.Generic;

namespace XLibur.Excel;

/// <summary>
/// Shared implementation of <see cref="IXLPivotField"/> fluent setters and property
/// delegates. Subclasses provide the concrete <see cref="GetField"/> resolution.
/// </summary>
internal abstract class XLPivotFieldBase : IXLPivotField
{
    private protected readonly XLPivotTable PivotTable;

    private protected XLPivotFieldBase(XLPivotTable pivotTable)
    {
        PivotTable = pivotTable;
    }

    // --- Abstract / virtual members that subclasses must provide ---

    public abstract string SourceName { get; }
    public abstract string CustomName { get; set; }
    public abstract IReadOnlyList<XLCellValue> SelectedValues { get; }
    public abstract IXLPivotField AddSelectedValue(XLCellValue value);
    public abstract IXLPivotField AddSelectedValues(IEnumerable<XLCellValue> values);
    public abstract IXLPivotFieldStyleFormats StyleFormats { get; }
    public abstract bool IsOnRowAxis { get; }
    public abstract bool IsOnColumnAxis { get; }
    public abstract bool IsInFilterList { get; }
    public abstract int Offset { get; }

    /// <summary>
    /// Get the underlying <see cref="XLPivotTableField"/> for this pivot field.
    /// </summary>
    private protected abstract XLPivotTableField GetField();

    /// <summary>
    /// Get a value from the underlying field, or return a default when the field
    /// cannot provide one (e.g. data fields on an axis).
    /// </summary>
    private protected virtual T GetFieldValue<T>(System.Func<XLPivotTableField, T> getter, T defaultValue)
    {
        return getter(GetField());
    }

    // --- Properties delegated to GetField() ---

    public string SubtotalCaption
    {
        get => GetFieldValue(f => f.SubtotalCaption, string.Empty);
        set => GetField().SubtotalCaption = value;
    }

    public IReadOnlyCollection<XLSubtotalFunction> Subtotals => GetField().Subtotals;

    public bool IncludeNewItemsInFilter
    {
        get => GetFieldValue(f => f.IncludeNewItemsInFilter, false);
        set => GetField().IncludeNewItemsInFilter = value;
    }

    public bool Outline
    {
        get => GetFieldValue(f => f.Outline, true);
        set => GetField().Outline = value;
    }

    public bool Compact
    {
        get => GetFieldValue(f => f.Compact, true);
        set => GetField().Compact = value;
    }

    public bool? SubtotalsAtTop
    {
        get => GetFieldValue(f => f.SubtotalTop, true);
        set => GetField().SubtotalTop = value ?? true;
    }

    public bool RepeatItemLabels
    {
        get => GetFieldValue(f => f.RepeatItemLabels, false);
        set => GetField().RepeatItemLabels = value;
    }

    public bool InsertBlankLines
    {
        get => GetFieldValue(f => f.InsertBlankRow, false);
        set => GetField().InsertBlankRow = value;
    }

    public bool ShowBlankItems
    {
        get => GetFieldValue(f => f.ShowAll, true);
        set => GetField().ShowAll = value;
    }

    public bool InsertPageBreaks
    {
        get => GetFieldValue(f => f.InsertPageBreak, false);
        set => GetField().InsertPageBreak = value;
    }

    public virtual bool Collapsed
    {
        get => GetFieldValue(f => f.Collapsed, false);
        set => GetField().Collapsed = value;
    }

    public XLPivotSortType SortType
    {
        get => GetFieldValue(f => f.SortType, XLPivotSortType.Default);
        set => GetField().SortType = value;
    }

    // --- Fluent setters ---

    public IXLPivotField SetCustomName(string value)
    {
        CustomName = value;
        return this;
    }

    public IXLPivotField SetSubtotalCaption(string value)
    {
        SubtotalCaption = value;
        return this;
    }

    public IXLPivotField SetSubtotal(XLSubtotalFunction function, bool enabled)
    {
        if (enabled)
            GetField().AddSubtotal(function);
        else
            GetField().RemoveSubtotal(function);

        return this;
    }

    public IXLPivotField AddSubtotal(XLSubtotalFunction value)
    {
        GetField().AddSubtotal(value);
        return this;
    }

    public IXLPivotField SetIncludeNewItemsInFilter(bool value = true)
    {
        IncludeNewItemsInFilter = value;
        return this;
    }

    public IXLPivotField SetLayout(XLPivotLayout value)
    {
        GetField().SetLayout(value);
        return this;
    }

    public IXLPivotField SetSubtotalsAtTop(bool value = true)
    {
        SubtotalsAtTop = value;
        return this;
    }

    public IXLPivotField SetRepeatItemLabels(bool value = true)
    {
        RepeatItemLabels = value;
        return this;
    }

    public IXLPivotField SetInsertBlankLines(bool value = true)
    {
        InsertBlankLines = value;
        return this;
    }

    public IXLPivotField SetShowBlankItems(bool value = true)
    {
        ShowBlankItems = value;
        return this;
    }

    public IXLPivotField SetInsertPageBreaks(bool value = true)
    {
        InsertPageBreaks = value;
        return this;
    }

    public IXLPivotField SetCollapsed(bool value = true)
    {
        Collapsed = value;
        return this;
    }

    public IXLPivotField SetSort(XLPivotSortType value)
    {
        SortType = value;
        return this;
    }
}
