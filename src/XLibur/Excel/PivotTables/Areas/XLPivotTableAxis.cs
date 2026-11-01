using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace XLibur.Excel;

/// <summary>
/// A description of one axis (<see cref="XLPivotTable.RowAxis"/>/<see cref="XLPivotTable.ColumnAxis"/>)
/// of a <see cref="XLPivotTable"/>. It consists of fields in a specific order and values that make up
/// individual rows/columns of the axis.
/// </summary>
/// <remarks>
/// [ISO-29500] 18.10.1.17 colItems (Column Items), 18.10.1.84 rowItems (Row Items).
/// </remarks>
internal sealed class XLPivotTableAxis : IXLPivotFields
{
    private readonly XLPivotTable _pivotTable;

    private readonly XLPivotAxis _axis;

    /// <summary>
    /// Fields displayed on the axis, in the order of the fields on the axis.
    /// </summary>
    private readonly List<FieldIndex> _fields = new();

    /// <summary>
    /// Values of one row/column in an axis. An item names a value of each field on the axis by
    /// position, so items only ever come from a loaded file (<see cref="AddItem"/>): nothing here
    /// computes new ones for a field added in code. Adding or removing a field clears the list
    /// instead of trying to keep it in sync with <see cref="_fields"/>, and Excel lays the axis out
    /// again from what is left.
    /// </summary>
    private readonly List<XLPivotFieldAxisItem> _axisItems = new();

    internal XLPivotTableAxis(XLPivotTable pivotTable, XLPivotAxis axis)
    {
        _pivotTable = pivotTable;
        _axis = axis;
    }

    /// <summary>
    /// A list of fields to displayed on the axis. It determines which fields and in what order
    /// should the fields be displayed.
    /// </summary>
    internal IReadOnlyList<FieldIndex> Fields => _fields;

    /// <summary>
    /// Individual row/column parts of the axis.
    /// </summary>
    internal IReadOnlyList<XLPivotFieldAxisItem> Items => _axisItems;

    internal bool ContainsDataField => _fields.Any(x => x.IsDataField);

    IXLPivotField IXLPivotFields.Add(string sourceName) => Add(sourceName, sourceName);

    IXLPivotField IXLPivotFields.Add(string sourceName, string customName) => Add(sourceName, customName);

    void IXLPivotFields.Clear() => Clear();

    bool IXLPivotFields.Contains(string sourceName) => Contains(sourceName);

    bool IXLPivotFields.Contains(IXLPivotField pivotField) => Contains(pivotField.SourceName);

    IXLPivotField IXLPivotFields.Get(string sourceName)
    {
        if (!_pivotTable.TryGetSourceNameFieldIndex(sourceName, out var index) ||
            !_fields.Contains(index))
            throw new KeyNotFoundException($"Field with source name '{sourceName}' not found in {_axis}.");

        return new XLPivotTableAxisField(_pivotTable, index);
    }

    IXLPivotField IXLPivotFields.Get(int index)
    {
        if (index < 0 || index >= _fields.Count)
            throw new ArgumentOutOfRangeException(nameof(index));

        return new XLPivotTableAxisField(_pivotTable, _fields[index]);
    }

    int IXLPivotFields.IndexOf(string sourceName)
    {
        return IndexOf(sourceName);
    }

    int IXLPivotFields.IndexOf(IXLPivotField pf)
    {
        return IndexOf(pf.SourceName);
    }

    /// <summary>
    /// Take a field off the axis. The axis' <see cref="Items"/> go with it, for the same reason as
    /// in <see cref="RemoveDataField"/>: an item names a value of each field on the axis by position
    /// through its <c>x</c> elements, so once a field is gone from the axis, every item is stale. An
    /// axis with no items is what a table built in code writes anyway, and Excel lays the axis out
    /// again from the fields that are left.
    /// </summary>
    void IXLPivotFields.Remove(string sourceName)
    {
        var index = IndexOf(sourceName);
        if (index == -1)
            return;

        _pivotTable.RemoveFieldFromAxis(_fields[index]);
        _fields.RemoveAt(index);
        _axisItems.Clear();
    }

    IEnumerator<IXLPivotField> IEnumerable<IXLPivotField>.GetEnumerator() => GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public IEnumerator<XLPivotTableAxisField> GetEnumerator()
    {
        foreach (var fieldIndex in _fields)
            yield return new XLPivotTableAxisField(_pivotTable, fieldIndex);
    }

    internal int IndexOf(FieldIndex index)
    {
        return _fields.IndexOf(index);
    }

    /// <summary>
    /// Take the 'data' field (the <see cref="XLConstants.PivotTable.ValuesSentinalLabel"/>
    /// sentinel) off the axis, if the axis holds it. It names the data fields rather than a field
    /// of the cache, so it must go once there are no data fields left for it to name (#572).
    /// </summary>
    /// <remarks>
    /// The axis' <see cref="Items"/> go with it, as they do in <see cref="Clear"/>. A loaded file
    /// keeps the rendered rows/columns, and an item names a data field by index through its
    /// <c>i</c> attribute, so a file whose column axis showed three values saves
    /// <c>&lt;i i="1"&gt;</c> and <c>&lt;i i="2"&gt;</c>. Left behind, those would name data fields
    /// the saved file no longer has, which is the same dangling reference as the <c>-2</c> field
    /// itself. An axis with no items is what a table built in code writes anyway, and Excel lays
    /// the axis out again from the fields.
    /// </remarks>
    internal void RemoveDataField()
    {
        var index = IndexOf(FieldIndex.DataField);
        if (index < 0)
            return;

        _pivotTable.RemoveFieldFromAxis(FieldIndex.DataField);
        _fields.RemoveAt(index);
        _axisItems.Clear();
    }

    internal bool Contains(string sourceName)
    {
        if (!_pivotTable.TryGetSourceNameFieldIndex(sourceName, out var index))
            return false;

        return _fields.Contains(index);
    }

    /// <summary>
    /// Add field to the axis, as an index.
    /// </summary>
    internal void AddField(FieldIndex fieldIndex)
    {
        if (_pivotTable.IsFieldUsedOnAxis(fieldIndex))
            throw new ArgumentException("Field is already used on an axis.");

        _fields.Add(fieldIndex);
    }

    internal XLPivotTableAxisField AddField(string sourceName, string customName)
    {
        var index = _pivotTable.AddFieldToAxis(sourceName, customName, _axis);
        _fields.Add(index);

        // An item's x elements name a value of each field on the axis by position, one per field, so
        // an axis that already had items (a loaded file kept them, or a field was added and removed
        // before) has too few once another field joins it. Clearing them here is the same call as
        // IXLPivotFields.Remove and RemoveDataField make: Excel lays the axis out again from its
        // fields.
        _axisItems.Clear();

        return new XLPivotTableAxisField(_pivotTable, index);
    }

    private XLPivotTableAxisField Add(string sourceName, string customName)
    {
        var field = AddField(sourceName, customName);

        if (field.Offset == FieldIndex.DataField.Value)
            return field;

        // A field built in code has no subtotal setting of its own, so it gets the automatic subtotal
        // here. A field that a loaded file gave its subtotals keeps them, including none: Excel saves a
        // field on an axis with defaultSubtotal="0" and no subtotal item, and taking such a field off an
        // axis and putting it back must not invent a subtotal the file never had (#562).
        var pivotField = _pivotTable.PivotFields[field.Offset];
        if (!pivotField.SubtotalsFromFile)
            pivotField.AddSubtotal(XLSubtotalFunction.Automatic);

        return field;
    }

    /// <summary>
    /// Add a row/column axis values (i.e. values visible on the axis).
    /// </summary>
    internal void AddItem(XLPivotFieldAxisItem axisItem)
    {
        _axisItems.Add(axisItem);
    }

    internal void Clear()
    {
        foreach (var fieldIndex in _fields)
            _pivotTable.RemoveFieldFromAxis(fieldIndex);

        _axisItems.Clear();
        _fields.Clear();
    }

    private int IndexOf(string sourceName)
    {
        if (!_pivotTable.TryGetSourceNameFieldIndex(sourceName, out var fieldIndex))
            return -1;

        return _fields.IndexOf(fieldIndex);
    }
}
