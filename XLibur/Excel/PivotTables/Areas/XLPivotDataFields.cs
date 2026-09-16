using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace XLibur.Excel.PivotTables.Areas;

/// <summary>
/// A collection of <see cref="XLPivotDataField"/>.
/// </summary>
internal sealed class XLPivotDataFields : IXLPivotValues, IReadOnlyCollection<XLPivotDataField>
{
    private readonly XLPivotTable _pivotTable;

    /// <summary>
    /// Fields displayed in the data area of the pivot table, in the order fields are displayed.
    /// </summary>
    private readonly List<XLPivotDataField> _fields = [];

    internal XLPivotDataFields(XLPivotTable pivotTable)
    {
        _pivotTable = pivotTable;
    }

    public int Count => _fields.Count;

    #region IXLPivotValues

    public IXLPivotValue Add(string sourceName)
    {
        return AddField(sourceName, sourceName);
    }

    public IXLPivotValue Add(string sourceName, string customName)
    {
        return AddField(sourceName, customName);
    }

    public void Clear()
    {
        // Every position goes, so every 'data' field reference in the style formats is emptied and
        // its format dropped. Taking the positions off the top one at a time, rather than by a
        // rule of its own, keeps this and Remove on the one path, so the two cannot drift apart.
        // After taking position 'index' off, exactly 'index' values are left.
        for (var index = _fields.Count - 1; index >= 0; index--)
            _pivotTable.RemoveValueFromFormats(index, index);

        ClearWithoutPruningFormats();
    }

    public bool Contains(string customName)
    {
        return IndexOf(customName) != -1;
    }

    public bool Contains(IXLPivotValue pivotValue)
    {
        return Contains(pivotValue.CustomName);
    }

    public IXLPivotValue Get(string customName)
    {
        var dataField = _fields.SingleOrDefault(x => XLHelper.NameComparer.Equals(x.CustomName, customName));
        return dataField ?? throw new KeyNotFoundException($"Unable to find data field for '{customName}'.");
    }

    public IXLPivotValue Get(int index)
    {
        return _fields[index];
    }

    public int IndexOf(string customName)
    {
        return _fields.FindIndex(x => XLHelper.NameComparer.Equals(x.CustomName, customName));
    }

    public int IndexOf(IXLPivotValue pivotValue)
    {
        return IndexOf(pivotValue.CustomName);
    }

    public void Remove(string customName)
    {
        var index = IndexOf(customName);
        if (index == -1)
            return;

        var dataField = _fields[index];
        _fields.Remove(dataField);

        // The same field can back several values, such as a sum and a count of one column. The
        // flag says the field is in the data fields, so it stays while another value uses it.
        if (_fields.All(f => f.Field != dataField.Field))
            _pivotTable.RemoveFieldFromValues((FieldIndex)dataField.Field);

        // A style format names a value by its position here, so the positions after the removed
        // one have shifted and the references naming them must follow (#577). The field is already
        // out of _fields, so its count is how many values are left.
        _pivotTable.RemoveValueFromFormats(index, _fields.Count);

        SyncValuesSentinel();
    }

    IEnumerator<IXLPivotValue> IEnumerable<IXLPivotValue>.GetEnumerator()
    {
        return GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion

    internal XLPivotDataField AddField(string sourceName, string? customName)
    {
        if (!_pivotTable.TryGetSourceNameFieldIndex(sourceName, out var fieldIndex))
        {
            var validNames = string.Join("','", _pivotTable.PivotCache.FieldNames);
            throw new ArgumentOutOfRangeException(nameof(sourceName),
                $"Field '{sourceName}' is not in the fields of a pivot cache. Should be one of '{validNames}'.");
        }

        if (fieldIndex.IsDataField)
            throw new ArgumentException("'Values' field can be used only on row or column axis.");

        var dataField = new XLPivotDataField(_pivotTable, fieldIndex.Value)
        {
            DataFieldName = customName,
        };
        AddField(dataField);
        SyncValuesSentinel();

        return dataField;
    }

    /// <summary>
    /// Keep the 'data' field (the <see cref="XLConstants.PivotTable.ValuesSentinalLabel"/>
    /// sentinel) on the row/column axes in step with the values, so that a save never writes a
    /// <c>rowFields</c>/<c>colFields</c> entry of <c>-2</c> that names data fields the file does
    /// not have (#572). <see cref="AddField(string, string?)"/>, <see cref="Remove"/> and
    /// <see cref="Clear"/> all end here, so none of them can drift from the others.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The two directions are deliberately not mirror images of each other. The sentinel is
    /// <em>added</em> only from the second value on, because that is when Excel needs it to tell
    /// the values apart, and a file with several values and no sentinel on either axis makes Excel
    /// ask to repair. It is <em>removed</em> only when the last value goes, because that is when it
    /// names nothing at all.
    /// </para>
    /// <para>
    /// In between, with exactly one value, the sentinel is left exactly as it is. Excel itself
    /// never writes one there (of the 100 pivot tables Excel wrote in the test fixtures, all 27
    /// with two or more values carry the sentinel and none of the 73 with one or none does), but a
    /// caller may still place one by hand through <c>RowLabels</c>/<c>ColumnLabels</c>, which this
    /// library supports and round-trips. Nothing records which of the two put it there, so taking
    /// it off at one value would silently undo the caller's own placement; and unlike the empty
    /// case it is not a dangling reference, because there is still a data field for it to name.
    /// </para>
    /// </remarks>
    private void SyncValuesSentinel()
    {
        if (_fields.Count == 0)
        {
            // Only one axis can hold the sentinel, but clear both: it costs nothing and does not
            // rely on that invariant holding in a file we merely loaded.
            _pivotTable.RowAxis.RemoveDataField();
            _pivotTable.ColumnAxis.RemoveDataField();
            return;
        }

        if (_fields.Count > 1 &&
            !_pivotTable.RowAxis.ContainsDataField &&
            !_pivotTable.ColumnAxis.ContainsDataField)
        {
            _pivotTable.ColumnLabels.Add(XLConstants.PivotTable.ValuesSentinalLabel);
        }
    }

    /// <summary>
    /// Empty the values without touching the pivot table's style formats — the counterpart, on the
    /// removing side, of <see cref="AddField(XLPivotDataField)"/>.
    /// </summary>
    /// <remarks>
    /// <see cref="XLPivotTable.UpdateCacheFields"/> is the only caller: a cache refresh empties the
    /// values only to put the surviving ones straight back, so pruning here would throw away the
    /// formatting of a value the table still has on nothing more than a refresh. A refresh can
    /// also drop a value for good, when its source column is gone from the cache, and those are
    /// pruned by the caller before it gets here — this skips the pruning, it does not decide that
    /// none is needed. The public <see cref="Clear"/> prunes for itself, because there a caller
    /// really means every value to go.
    /// </remarks>
    internal void ClearWithoutPruningFormats()
    {
        foreach (var field in _fields)
            _pivotTable.RemoveFieldFromValues((FieldIndex)field.Field);

        _fields.Clear();
        SyncValuesSentinel();
    }

    /// <remarks>
    /// The loader's way in, and the one entry point that deliberately does not call
    /// <see cref="SyncValuesSentinel"/>: a loaded file states where its own 'data' field is, and
    /// the axis takes it straight from <c>rowFields</c>/<c>colFields</c>. Adding one here would
    /// put a sentinel on a file that never had one.
    /// </remarks>
    internal void AddField(XLPivotDataField dataField)
    {
        // Excel invariant - data field must have the flag if and only if it is in the data fields collection.
        _fields.Add(dataField);
        _pivotTable.PivotFields[dataField.Field].DataField = true;
    }

    public IEnumerator<XLPivotDataField> GetEnumerator()
    {
        return _fields.GetEnumerator();
    }
}
