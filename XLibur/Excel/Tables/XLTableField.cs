using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using XLibur.Extensions;

namespace XLibur.Excel;

[DebuggerDisplay("{Name}")]
internal sealed class XLTableField : IXLTableField
{
    internal XLTotalsRowFunction totalsRowFunction;
    internal string? totalsRowLabel;
    private readonly XLTable _table;

    private IXLRangeColumn? _column;
    private string _name;

    public XLTableField(XLTable table, string name)
    {
        this._table = table;
        this._name = name;
    }

    public IXLRangeColumn Column
    {
        get
        {
            _column ??= _table.AsRange().Column(Index + 1);
            return _column;
        }
        internal set => _column = value;
    }

    public IXLCells DataCells
    {
        get
        {
            return Column.Cells(c =>
            {
                if (_table.ShowHeaderRow && c.Equals(HeaderCell))
                    return false;
                if (_table.ShowTotalsRow && c.Equals(TotalsCell))
                    return false;
                return true;
            });
        }
    }

    public IXLCell? HeaderCell => !_table.ShowHeaderRow ? null : Column.FirstCell();

    public int Index
    {
        get;
        internal set
        {
            if (field == value) return;
            field = value;
            _column = null;
        }
    }

    public string Name
    {
        get => _name;
        set
        {
            if (_name == value) return;

            if (_table.ShowHeaderRow)
                ((XLCell)_table.HeadersRow(false)!.Cell(Index + 1)).SetValue(value, setTableHeader: false, checkMergedRanges: true);

            _table.RenameField(_name, value);
            _name = value;
        }
    }

    public IXLTable Table => _table;

    public IXLCell? TotalsCell
    {
        get
        {
            if (!_table.ShowTotalsRow)
                return null;

            return Column.LastCell();
        }
    }

    public string TotalsRowFormulaA1
    {
        get => _table.TotalsRow()!.Cell(Index + 1).FormulaA1;
        set
        {
            totalsRowFunction = XLTotalsRowFunction.Custom;
            _table.TotalsRow()!.Cell(Index + 1).FormulaA1 = value;
        }
    }

    public string TotalsRowFormulaR1C1
    {
        get => _table.TotalsRow()!.Cell(Index + 1).FormulaR1C1;
        set
        {
            totalsRowFunction = XLTotalsRowFunction.Custom;
            _table.TotalsRow()!.Cell(Index + 1).FormulaR1C1 = value;
        }
    }

    public XLTotalsRowFunction TotalsRowFunction
    {
        get => totalsRowFunction;
        set
        {
            totalsRowFunction = value;
            UpdateTableFieldTotalsRowFormula();
        }
    }

    public string? TotalsRowLabel
    {
        get => totalsRowLabel;
        set
        {
            totalsRowFunction = XLTotalsRowFunction.None;
            ((XLCell)_table.TotalsRow()!.Cell(Index + 1)).SetValue(value, setTableHeader: false, checkMergedRanges: true);
            totalsRowLabel = value;
        }
    }

    public void Delete()
    {
        Delete(true);
    }

    internal void Delete(bool deleteUnderlyingRangeColumn)
    {
        var fields = _table.Fields.Cast<XLTableField>().ToArray();

        if (deleteUnderlyingRangeColumn)
        {
            _table.AsRange().ColumnQuick(Index + 1).Delete();
        }

        fields.Where(f => f.Index > Index).ForEach(f => f.Index--);
        _table.FieldNames.Remove(Name);
    }

    public bool IsConsistentDataType()
    {
        var dataTypes = Column
            .Cells()
            .Skip(_table.ShowHeaderRow ? 1 : 0)
            .Select(c => c.DataType);

        if (_table.ShowTotalsRow)
            dataTypes = dataTypes.SkipLast();

        var distinctDataTypes = dataTypes
            .GroupBy(dt => dt)
            .Select(g => new { Key = g.Key, Count = g.Count() });

        return distinctDataTypes.Count() == 1;
    }

    public bool IsConsistentFormula()
    {
        var formulas = Column
            .Cells()
            .Skip(_table.ShowHeaderRow ? 1 : 0)
            .Select(c => c.FormulaR1C1);

        if (_table.ShowTotalsRow)
            formulas = formulas.SkipLast();

        var distinctFormulas = formulas
            .GroupBy(f => f)
            .Select(g => new { Key = g.Key, Count = g.Count() });

        return distinctFormulas.Count() == 1;
    }

    public bool IsConsistentStyle()
    {
        var styles = Column
            .Cells()
            .Skip(_table.ShowHeaderRow ? 1 : 0)
            .OfType<XLCell>()
            .Select(c => c.StyleValue);

        if (_table.ShowTotalsRow)
            styles = styles.SkipLast();

        var distinctStyles = styles
            .Distinct();

        return distinctStyles.Count() == 1;
    }

    private static readonly IEnumerable<string> QuotedTableFieldCharacters = ["'", "#"];

    internal void UpdateTableFieldTotalsRowFormula()
    {
        if (TotalsRowFunction != XLTotalsRowFunction.None && TotalsRowFunction != XLTotalsRowFunction.Custom)
        {
            var cell = _table.TotalsRow()!.Cell(Index + 1);
            var formulaCode = TotalsRowFunction switch
            {
                XLTotalsRowFunction.Sum => "109",
                XLTotalsRowFunction.Minimum => "105",
                XLTotalsRowFunction.Maximum => "104",
                XLTotalsRowFunction.Average => "101",
                XLTotalsRowFunction.Count => "103",
                XLTotalsRowFunction.CountNumbers => "102",
                XLTotalsRowFunction.StandardDeviation => "107",
                XLTotalsRowFunction.Variance => "110",
                _ => string.Empty,
            };

            var modifiedName = Name;
            QuotedTableFieldCharacters.ForEach(c => modifiedName = modifiedName.Replace(c, "'" + c));

            if (modifiedName.StartsWith(' ') || modifiedName.EndsWith(' '))
            {
                modifiedName = "[" + modifiedName + "]";
            }

            var prependTableName = modifiedName.Contains(' ');

            cell.FormulaA1 = $"SUBTOTAL({formulaCode},{(prependTableName ? _table.Name : string.Empty)}[{modifiedName}])";
            var lastCell = _table.LastRow()!.Cell(Index + 1);
            if (lastCell.DataType != XLDataType.Text)
            {
                cell.Style.NumberFormat = lastCell.Style.NumberFormat;
            }
        }
    }
}
