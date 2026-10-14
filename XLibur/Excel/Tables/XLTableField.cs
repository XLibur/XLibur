using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using XLibur.Excel.Tables;
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
        _table = table;
        _name = name;
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
            totalsRowLabel = null;
            _table.TotalsRow()!.Cell(Index + 1).FormulaA1 = value;
        }
    }

    public string TotalsRowFormulaR1C1
    {
        get => _table.TotalsRow()!.Cell(Index + 1).FormulaR1C1;
        set
        {
            totalsRowFunction = XLTotalsRowFunction.Custom;
            totalsRowLabel = null;
            _table.TotalsRow()!.Cell(Index + 1).FormulaR1C1 = value;
        }
    }

    /// <remarks>
    /// A function replaces the field's label, as in Excel: the totals cell then holds the function's
    /// formula, and the table part would otherwise give the field a label the cell does not hold.
    /// <see cref="XLTotalsRowFunction.Custom"/> leaves the cell, and so the label, until a formula
    /// is set (<see cref="TotalsRowFormulaA1"/>).
    /// </remarks>
    public XLTotalsRowFunction TotalsRowFunction
    {
        get => totalsRowFunction;
        set
        {
            totalsRowFunction = value;
            if (value is not (XLTotalsRowFunction.None or XLTotalsRowFunction.Custom))
                totalsRowLabel = null;

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

    /// <summary>
    /// Sets the label the load read from the table part, and leaves the totals cell as the file has
    /// it (#742). Loading the label is then not a change to the cells, and a save can keep them as
    /// the file had them (#702, #721).
    /// </summary>
    /// <remarks>
    /// Excel shows the cell, not the label (checked over COM): a totals cell left blank under a label
    /// opens blank. The setter of <see cref="TotalsRowLabel"/> would instead replace whatever the
    /// cell holds, such as a number, with the label. It would also strip a leading quote, as any value
    /// set through the API does, though Excel writes a label typed as <c>''Total</c> as the text
    /// <c>'Total</c> with <c>quotePrefix</c>, and the same text in the table part, so the quote is
    /// part of the label.
    /// <para>
    /// A totals function the load read for the same column is kept, as is its formula in the cell, so
    /// a table part that gives a column both is written back as it was.
    /// </para>
    /// </remarks>
    internal void LoadTotalsRowLabel(string label) => totalsRowLabel = label;

    /// <summary>
    /// Sets the totals function the load read from the table part, and leaves a formula the totals
    /// cell already holds as the file has it (#749).
    /// </summary>
    /// <remarks>
    /// The setter of <see cref="TotalsRowFunction"/> writes the function's formula into the cell. Over
    /// a loaded formula that is a change: the formula is then dirty, so a save writes it without its
    /// cached value, and its text is XLibur's spelling, not the file's (<c>Table1[[Purchase price]]</c>
    /// for <c>Table1[Purchase price]</c>, <c>[Photo?]</c> for <c>Table1[Photo?]</c>). A cell with no
    /// formula still gets the function's formula, as it did before.
    /// </remarks>
    internal void LoadTotalsRowFunction(XLTotalsRowFunction function)
    {
        if (!TotalsCellHasFormula())
        {
            TotalsRowFunction = function;
            return;
        }

        // Clears the label as the setter does. A label the table part gives is loaded after this.
        totalsRowFunction = function;
        if (function is not (XLTotalsRowFunction.None or XLTotalsRowFunction.Custom))
            totalsRowLabel = null;
    }

    /// <summary>
    /// Sets the custom totals formula the load read from the table part, and leaves a formula the
    /// totals cell already holds as the file has it (#749). See <see cref="LoadTotalsRowFunction"/>.
    /// </summary>
    internal void LoadTotalsRowFormula(string formulaA1)
    {
        if (TotalsCellHasFormula())
        {
            totalsRowFunction = XLTotalsRowFunction.Custom;
            totalsRowLabel = null;
        }
        else
        {
            TotalsRowFormulaA1 = formulaA1;
        }
    }

    private bool TotalsCellHasFormula() => _table.TotalsRow()!.Cell(Index + 1).HasFormula;

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
            .Select(g => new { g.Key, Count = g.Count() });

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
            .Select(g => new { g.Key, Count = g.Count() });

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

            // Excel requires a column name that contains a space (leading, trailing or
            // interior) to be wrapped in an extra pair of brackets in a structured
            // reference, e.g. Table1[[Feb 2023]]. Emitting the single-bracket form
            // Table1[Feb 2023] yields a totals-row formula that Excel flags with a
            // #NAME error and "repairs" on open (see issue #2864).
            var hasSpace = modifiedName.Contains(' ');
            if (hasSpace)
            {
                modifiedName = "[" + modifiedName + "]";
            }

            var prependTableName = hasSpace;

            cell.FormulaA1 = $"SUBTOTAL({formulaCode},{(prependTableName ? _table.Name : string.Empty)}[{modifiedName}])";
            var lastCell = _table.LastRow()!.Cell(Index + 1);
            if (lastCell.DataType != XLDataType.Text)
            {
                cell.Style.NumberFormat = lastCell.Style.NumberFormat;
            }
        }
    }
}
