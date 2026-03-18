using System;
using System.Collections.Generic;
using System.Linq;

namespace XLibur.Excel;

/// <summary>
/// A fluent API for one field in <see cref="XLPivotTableAxis"/>, either
/// <see cref="XLPivotTable.RowLabels"/> or <see cref="XLPivotTable.ColumnLabels"/>.
/// </summary>
internal sealed class XLPivotTableAxisField : XLPivotFieldBase
{
    private readonly FieldIndex _index;

    internal XLPivotTableAxisField(XLPivotTable pivotTable, FieldIndex index)
        : base(pivotTable)
    {
        _index = index;
    }

    public override string SourceName
    {
        get
        {
            if (_index.IsDataField)
                return XLConstants.PivotTable.ValuesSentinalLabel;

            return PivotTable.PivotCache.FieldNames[_index];
        }
    }

    public override string CustomName
    {
        get => GetFieldValue(f => f.Name!, PivotTable.DataCaption);
        set
        {
            if (_index.IsDataField)
            {
                PivotTable.DataCaption = value;
                return;
            }

            if (PivotTable.TryGetCustomNameFieldIndex(value, out var idx) && idx != _index)
                throw new ArgumentException($"Custom name '{value}' is already used by another field.");

            PivotTable.PivotFields[_index].Name = value;
        }
    }

    public override bool Collapsed
    {
        get => GetFieldValue(f => !f.Items.Any(i => i.ShowDetails), false);
        set
        {
            foreach (var item in GetField().Items)
                item.ShowDetails = !value;
        }
    }

    public override IReadOnlyList<XLCellValue> SelectedValues => Array.Empty<XLCellValue>();

    public override IXLPivotFieldStyleFormats StyleFormats => new XLPivotTableAxisFieldStyleFormats(PivotTable, this);

    public override bool IsOnRowAxis => GetFieldValue(f => f.Axis == XLPivotAxis.AxisRow, PivotTable.DataOnRows);

    public override bool IsOnColumnAxis => GetFieldValue(f => f.Axis == XLPivotAxis.AxisCol, !PivotTable.DataOnRows);

    public override bool IsInFilterList => false;

    public override int Offset => _index;

    public override IXLPivotField AddSelectedValue(XLCellValue value) => this;

    public override IXLPivotField AddSelectedValues(IEnumerable<XLCellValue> values) => this;

    internal XLPivotAxis Axis => IsOnColumnAxis ? XLPivotAxis.AxisCol : XLPivotAxis.AxisRow;

    /// <summary>
    /// Get position of the field on the axis, starting at 0.
    /// </summary>
    internal int Position
    {
        get
        {
            var axis = IsOnColumnAxis ? PivotTable.ColumnAxis : PivotTable.RowAxis;
            var position = axis.IndexOf(_index);
            if (position == -1)
                throw new InvalidOperationException("Field is not on the axis.");

            return position;
        }
    }

    private protected override XLPivotTableField GetField()
    {
        if (_index.IsDataField)
            throw new InvalidOperationException("Can't set this property on a data field.");

        return PivotTable.PivotFields[_index];
    }

    private protected override T GetFieldValue<T>(Func<XLPivotTableField, T> getter, T defaultValue)
    {
        if (_index.IsDataField)
            return defaultValue;

        return getter(PivotTable.PivotFields[_index]);
    }
}
