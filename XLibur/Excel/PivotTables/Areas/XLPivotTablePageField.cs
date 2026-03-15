using System;
using System.Collections.Generic;
using System.Linq;

namespace XLibur.Excel;

/// <summary>
/// Fluent API for filter fields of a <see cref="XLPivotTable"/>. This class shouldn't contain any
/// state, only logic to change state per API.
/// </summary>
internal sealed class XLPivotTablePageField : XLPivotFieldBase
{
    private readonly XLPivotPageField _filterField;

    internal XLPivotTablePageField(XLPivotTable pivotTable, XLPivotPageField filterField)
        : base(pivotTable)
    {
        _filterField = filterField;
    }

    public override string SourceName => PivotTable.PivotCache.FieldNames[_filterField.Field];

    public override string CustomName
    {
        get => GetField().Name!;
        set => GetField().Name = value;
    }

    public override IReadOnlyList<XLCellValue> SelectedValues
    {
        get
        {
            var shownItems = GetField().Items.Where(i => !i.Hidden);
            var selectedValues = new List<XLCellValue>();
            foreach (var selectedItem in shownItems)
            {
                var selectedValue = selectedItem.GetValue();
                if (selectedValue is not null)
                    selectedValues.Add(selectedValue.Value);
            }

            return selectedValues;
        }
    }

    public override IXLPivotField AddSelectedValue(XLCellValue value)
    {
        // Try to keep the original behavior of XLibur - it always allows multiple selected items for added values.
        // But it's complete kludge with no sane semantic that will be nuked ASAP.
        var pivotField = GetField();

        var nothingSelected = _filterField.ItemIndex is null && !pivotField.MultipleItemSelectionAllowed;
        if (nothingSelected)
        {
            var fieldItem = pivotField.GetOrAddItem(value);
            _filterField.ItemIndex = fieldItem.ItemIndex;
            return this;
        }

        var oneItemSelected = _filterField.ItemIndex is not null && !pivotField.MultipleItemSelectionAllowed;
        if (oneItemSelected)
        {
            // Switch to multiple
            pivotField.MultipleItemSelectionAllowed = true;
            foreach (var item in pivotField.Items.Where(x => x.ItemType == XLPivotItemType.Data))
                item.Hidden = true;

            var selectedItem = pivotField.Items.Single(i => i.ItemIndex == _filterField.ItemIndex);
            selectedItem.Hidden = false;
            _filterField.ItemIndex = null;
            var fieldItem = pivotField.GetOrAddItem(value);
            fieldItem.Hidden = false;
            return this;
        }
        else
        {
            // Add another item to selected item filters.
            var fieldItem = pivotField.GetOrAddItem(value);
            fieldItem.Hidden = false;
            return this;
        }
    }

    public override IXLPivotField AddSelectedValues(IEnumerable<XLCellValue> values)
    {
        foreach (var value in values)
            AddSelectedValue(value);

        return this;
    }

    public override IXLPivotFieldStyleFormats StyleFormats => throw new NotImplementedException("Styles for filter fields are not yet implemented.");
    public override bool IsOnRowAxis => false;
    public override bool IsOnColumnAxis => false;
    public override bool IsInFilterList => true;
    public override int Offset => _filterField.Field;

    private protected override XLPivotTableField GetField()
    {
        return PivotTable.PivotFields[_filterField.Field];
    }
}
