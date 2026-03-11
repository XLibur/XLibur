using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace XLibur.Excel;

[DebuggerDisplay("{Name}")]
internal sealed class XLTableField : IXLTableField
{
    internal XLTotalsRowFunction totalsRowFunction;
    internal string? totalsRowLabel;
    private readonly XLTable table;

    private IXLRangeColumn? _column;
    private int index;
    private string name;

    public XLTableField(XLTable table, string name)
    {
        this.table = table;
        this.name = name;
    }

    public IXLRangeColumn Column
    {
        get
        {
            _column ??= table.AsRange().Column(Index + 1);
            return _column;
        }
        internal set
        {
            _column = value;
        }
    }

    public IXLCells DataCells
    {
        get
        {
            return Column.Cells(c =>
            {
                if (table.ShowHeaderRow && c.Equals(HeaderCell))
                    return false;
                if (table.ShowTotalsRow && c.Equals(TotalsCell))
                    return false;
                return true;
            });
        }
    }

    public IXLCell? HeaderCell
    {
        get
        {
            if (!table.ShowHeaderRow)
                return null;

            return Column.FirstCell();
        }
    }

    public int Index
    {
        get { return index; }
        internal set
        {
            if (index == value) return;
            index = value;
            _column = null;
        }
    }

    public string Name
    {
        get
        {
            return name;
        }
        set
        {
            if (name == value) return;

            if (table.ShowHeaderRow)
                ((XLCell)table.HeadersRow(false)!.Cell(Index + 1)).SetValue(value, setTableHeader: false, checkMergedRanges: true);

            table.RenameField(name, value);
            name = value;
        }
    }

    public IXLTable Table { get { return table; } }

    public IXLCell? TotalsCell
    {
        get
        {
            if (!table.ShowTotalsRow)
                return null;

            return Column.LastCell();
        }
    }

    public string TotalsRowFormulaA1
    {
        get { return table.TotalsRow()!.Cell(Index + 1).FormulaA1; }
        set
        {
            totalsRowFunction = XLTotalsRowFunction.Custom;
            table.TotalsRow()!.Cell(Index + 1).FormulaA1 = value;
        }
    }

    public string TotalsRowFormulaR1C1
    {
        get { return table.TotalsRow()!.Cell(Index + 1).FormulaR1C1; }
        set
        {
            totalsRowFunction = XLTotalsRowFunction.Custom;
            table.TotalsRow()!.Cell(Index + 1).FormulaR1C1 = value;
        }
    }

    public XLTotalsRowFunction TotalsRowFunction
    {
        get { return totalsRowFunction; }
        set
        {
            totalsRowFunction = value;
            UpdateTableFieldTotalsRowFormula();
        }
    }

    public string? TotalsRowLabel
    {
        get { return totalsRowLabel; }
        set
        {
            totalsRowFunction = XLTotalsRowFunction.None;
            ((XLCell)table.TotalsRow()!.Cell(Index + 1)).SetValue(value, setTableHeader: false, checkMergedRanges: true);
            totalsRowLabel = value;
        }
    }

    public void Delete()
    {
        Delete(true);
    }

    internal void Delete(bool deleteUnderlyingRangeColumn)
    {
        var fields = table.Fields.Cast<XLTableField>().ToArray();

        if (deleteUnderlyingRangeColumn)
        {
            table.AsRange().ColumnQuick(Index + 1).Delete();
        }

        fields.Where(f => f.Index > Index).ForEach(f => f.Index--);
        table.FieldNames.Remove(Name);
    }

    public bool IsConsistentDataType()
    {
        var dataTypes = Column
            .Cells()
            .Skip(table.ShowHeaderRow ? 1 : 0)
            .Select(c => c.DataType);

        if (table.ShowTotalsRow)
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
            .Skip(table.ShowHeaderRow ? 1 : 0)
            .Select(c => c.FormulaR1C1);

        if (table.ShowTotalsRow)
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
            .Skip(table.ShowHeaderRow ? 1 : 0)
            .OfType<XLCell>()
            .Select(c => c.StyleValue);

        if (table.ShowTotalsRow)
            styles = styles.SkipLast();

        var distinctStyles = styles
            .Distinct();

        return distinctStyles.Count() == 1;
    }

    private static IEnumerable<string> QuotedTableFieldCharacters = ["'", "#"];

    internal void UpdateTableFieldTotalsRowFormula()
    {
        if (TotalsRowFunction != XLTotalsRowFunction.None && TotalsRowFunction != XLTotalsRowFunction.Custom)
        {
            var cell = table.TotalsRow()!.Cell(Index + 1);
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

            if (modifiedName.StartsWith(" ") || modifiedName.EndsWith(" "))
            {
                modifiedName = "[" + modifiedName + "]";
            }

            var prependTableName = modifiedName.Contains(" ");

            cell.FormulaA1 = $"SUBTOTAL({formulaCode},{(prependTableName ? table.Name : string.Empty)}[{modifiedName}])";
            var lastCell = table.LastRow()!.Cell(Index + 1);
            if (lastCell.DataType != XLDataType.Text)
            {
                cell.Style.NumberFormat = lastCell.Style.NumberFormat;
            }
        }
    }
}
