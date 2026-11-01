using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace XLibur.Excel;

/// <summary>
/// Page/filter fields of a <see cref="XLPivotTable"/>. It determines filter values and layout.
/// It is accessible through fluent API <see cref="XLPivotTable.ReportFilters"/>.
/// </summary>
internal sealed class XLPivotTableFilters : IXLPivotFields
{
    private readonly XLPivotTable _pivotTable;

    /// <summary>
    /// Filter fields in correct order. The layout is determined by
    /// <see cref="XLPivotTable.FilterFieldsPageWrap"/> and
    /// <see cref="XLPivotTable.FilterAreaOrder"/>.
    /// </summary>
    private readonly List<XLPivotPageField> _fields = new();

    internal XLPivotTableFilters(XLPivotTable pivotTable)
    {
        _pivotTable = pivotTable;
    }

    IXLPivotField IXLPivotFields.Add(string sourceName) => Add(sourceName, sourceName);

    IXLPivotField IXLPivotFields.Add(string sourceName, string customName) => Add(sourceName, customName);

    public void Clear()
    {
        var filterHeight = GetSizeWithGap().Height;

        foreach (var field in _fields)
            _pivotTable.RemoveFieldFromAxis((FieldIndex)field.Field);

        _fields.Clear();

        _pivotTable.MoveAreaForFilterHeightChange(filterHeight);
    }

    public bool Contains(string sourceName)
    {
        return IndexOf(sourceName) >= 0;
    }

    public bool Contains(IXLPivotField pivotField)
    {
        return Contains(pivotField.SourceName);
    }

    public IXLPivotField Get(string sourceName)
    {
        if (!_pivotTable.TryGetSourceNameFieldIndex(sourceName, out var fieldIndex))
            throw new KeyNotFoundException($"Field with source name '{sourceName}' not found in {XLPivotAxis.AxisPage}.");

        var filterField = _fields.SingleOrDefault(f => f.Field == fieldIndex);
        if (filterField is null)
            throw new KeyNotFoundException($"Field with source name '{sourceName}' not found in {XLPivotAxis.AxisPage}.");

        return new XLPivotTablePageField(_pivotTable, filterField);
    }

    public IXLPivotField Get(int index)
    {
        if (index < 0 || index >= _fields.Count)
            throw new ArgumentOutOfRangeException(nameof(index));

        return new XLPivotTablePageField(_pivotTable, _fields[index]);
    }

    IEnumerator<IXLPivotField> IEnumerable<IXLPivotField>.GetEnumerator() => GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public IEnumerator<XLPivotTablePageField> GetEnumerator()
    {
        foreach (var field in _fields)
            yield return new XLPivotTablePageField(_pivotTable, field);
    }

    public int IndexOf(string sourceName)
    {
        if (!_pivotTable.TryGetSourceNameFieldIndex(sourceName, out var fieldIndex))
            return -1;

        return _fields.FindIndex(f => f.Field == fieldIndex);
    }

    public int IndexOf(IXLPivotField pf)
    {
        return IndexOf(pf.SourceName);
    }

    public void Remove(string sourceName)
    {
        var index = IndexOf(sourceName);
        if (index == -1)
            return;

        var filterHeight = GetSizeWithGap().Height;

        // index is the filter's position among the report filters, not a pivot field index.
        var fieldIndex = (FieldIndex)_fields[index].Field;
        _fields.RemoveAt(index);
        _pivotTable.RemoveFieldFromAxis(fieldIndex);

        _pivotTable.MoveAreaForFilterHeightChange(filterHeight);
    }

    internal IReadOnlyList<XLPivotPageField> Fields => _fields;

    internal XLPivotTablePageField Add(string sourceName, string customName)
    {
        if (sourceName == XLConstants.PivotTable.ValuesSentinalLabel)
            throw new ArgumentException($"The column '{sourceName}' does not appear in the source range.", nameof(sourceName));

        var filterHeight = GetSizeWithGap().Height;

        var fieldIndex = _pivotTable.AddFieldToAxis(sourceName, customName, XLPivotAxis.AxisPage);
        var filterField = new XLPivotPageField(fieldIndex);
        _fields.Add(filterField);

        _pivotTable.MoveAreaForFilterHeightChange(filterHeight);
        return new XLPivotTablePageField(_pivotTable, filterField);
    }

    internal bool Contains(FieldIndex fieldIndex)
    {
        return _fields.FindIndex(f => f.Field == fieldIndex) >= 0;
    }

    internal void AddField(XLPivotPageField pageField)
    {
        _fields.Add(pageField);
    }

    /// <summary>
    /// Number of rows/cols occupied by the filter area. Filter area is above the pivot table and it
    /// optional (i.e. size <c>0</c> indicates no filter).
    /// </summary>
    internal (int Width, int Height) GetSize()
    {
        return GetSize(_fields.Count, _pivotTable.FilterAreaOrder, _pivotTable.FilterFieldsPageWrap);
    }

    /// <summary>
    /// Number of rows/cols occupied by the filter area, including the gap below, if there is at least one filter.
    /// </summary>
    internal (int Width, int Height) GetSizeWithGap()
    {
        return GetSizeWithGap(_fields.Count, _pivotTable.FilterAreaOrder, _pivotTable.FilterFieldsPageWrap);
    }

    private static (int Width, int Height) GetSize(int fieldCount, XLFilterAreaOrder order, int filterWrap)
    {
        // A wrap of 0 is no wrap at all, so the fields never start a second line.
        if (filterWrap == 0)
            filterWrap = int.MaxValue;

        // The fields fill one line up to the wrap, then start the next. The line runs down the
        // sheet for DownThenOver and across it for OverThenDown, so the wrap caps how long a
        // line gets, and the number of lines is the area's other dimension.
        var lineLength = Math.Min(fieldCount, filterWrap);

        // Written this way rather than (fieldCount + filterWrap - 1) / filterWrap, which
        // overflows once filterWrap is int.MaxValue.
        var lineCount = fieldCount == 0 ? 0 : (fieldCount - 1) / filterWrap + 1;

        return order switch
        {
            XLFilterAreaOrder.DownThenOver => new ValueTuple<int, int>(lineCount, lineLength),
            XLFilterAreaOrder.OverThenDown => new ValueTuple<int, int>(lineLength, lineCount),
            _ => throw new UnreachableException(),
        };
    }

    private static (int Width, int Height) GetSizeWithGap(int fieldCount, XLFilterAreaOrder order, int filterWrap)
    {
        var filtersSize = GetSize(fieldCount, order, filterWrap);
        return filtersSize.Height > 0
            ? (filtersSize.Width, filtersSize.Height + 1)
            : filtersSize;
    }
}
