using System;
using System.Collections.Generic;
using XLibur.Excel.Coordinates;
using XLibur.Excel.Rows;

namespace XLibur.Excel;

/// <summary>
/// What a worksheet's cells were when the workbook load ended, kept so a save can tell whether the
/// sheet's <c>&lt;sheetData&gt;</c> would come out as the file had it (#702).
/// </summary>
/// <remarks>
/// It errs towards reporting a change: setting a cell to the value it already holds counts, and so
/// does materialising a column, which gives the cells of each existing row a style of their own. The
/// aim is never to miss a change. It records what the sheet-data writer reads:
/// <list type="bullet">
/// <item>The <see cref="ISlice.Version"/> of each cell slice. <see cref="FormulaSlice.Version"/> also
/// counts formulas edited in place, such as by a sheet rename or an array formula shifted by an insert
/// on another sheet.</item>
/// <item>The attributes the writer gives each row. Rows have many setters and no change counter, so
/// they are compared as they are, not through a version.</item>
/// <item>The styles a cell with no style of its own inherits: the sheet's, and those of its row and
/// column.</item>
/// <item>The workbook's date system, which sets the serial number a date is written as.</item>
/// </list>
/// Formulas get one more check, at save. A formula that is dirty then has its cached value
/// recalculated or dropped by the save. Editing another sheet can make it dirty without writing to
/// this one.
/// <para>
/// Not tracked, because the tiers of #702 exclude these sheets: labels in a table's totals row, which
/// the writer takes from the table; and the <c>cm</c> index of a dynamic array, which is
/// workbook-wide. A sparkline added to a blank cell can add an empty <c>&lt;row&gt;</c> element and
/// is not tracked either; it changes no cell.
/// </para>
/// </remarks>
internal sealed class SheetDataBaseline
{
    private readonly int _valueVersion;
    private readonly int _formulaVersion;
    private readonly int _styleVersion;
    private readonly int _miscVersion;
    private readonly XLStyleValue _sheetStyle;
    private readonly bool _use1904DateSystem;

    /// <summary>
    /// Every row that writes an attribute or passes a style to its cells, ordered by row number.
    /// </summary>
    private readonly RowAttributes[] _rows;

    /// <summary>
    /// Every column whose style differs from the sheet's, ordered by column number.
    /// </summary>
    private readonly (int Column, XLStyleValue Style)[] _columns;

    private SheetDataBaseline(XLWorksheet sheet)
    {
        var cells = sheet.Internals.CellsCollection;
        _valueVersion = cells.ValueSlice.Version;
        _formulaVersion = cells.FormulaSlice.Version;
        _styleVersion = cells.StyleSlice.Version;
        _miscVersion = cells.MiscSlice.Version;
        _sheetStyle = sheet.StyleValue;
        _use1904DateSystem = sheet.Workbook.Use1904DateSystem;
        _rows = CaptureRows(sheet);
        _columns = CaptureColumns(sheet);
    }

    /// <summary>
    /// Records the cells of <paramref name="sheet"/> as they are now.
    /// </summary>
    internal static SheetDataBaseline Capture(XLWorksheet sheet) => new(sheet);

    /// <summary>
    /// Would <paramref name="sheet"/>'s cells be written as they were when this baseline was captured?
    /// </summary>
    internal bool Matches(XLWorksheet sheet)
    {
        var cells = sheet.Internals.CellsCollection;
        if (cells.ValueSlice.Version != _valueVersion
            || cells.FormulaSlice.Version != _formulaVersion
            || cells.StyleSlice.Version != _styleVersion
            || cells.MiscSlice.Version != _miscVersion)
        {
            return false;
        }

        if (sheet.StyleValue != _sheetStyle || sheet.Workbook.Use1904DateSystem != _use1904DateSystem)
            return false;

        return RowsMatch(sheet) && ColumnsMatch(sheet) && !HasDirtyFormula(cells.FormulaSlice);
    }

    private static RowAttributes[] CaptureRows(XLWorksheet sheet)
    {
        var rows = new RowAttributes[sheet.Internals.RowsCollection.Count];
        var count = 0;
        foreach (var (rowNumber, row) in sheet.Internals.RowsCollection)
        {
            var attributes = RowAttributes.Of(rowNumber, row);
            if (!attributes.IsPlain(sheet.StyleValue))
                rows[count++] = attributes;
        }

        Array.Resize(ref rows, count);
        Array.Sort(rows, RowNumberComparer.Instance);
        return rows;
    }

    private static (int Column, XLStyleValue Style)[] CaptureColumns(XLWorksheet sheet)
    {
        var columns = new (int Column, XLStyleValue Style)[sheet.Internals.ColumnsCollection.Count];
        var count = 0;
        foreach (var (columnNumber, column) in sheet.Internals.ColumnsCollection)
        {
            if (column.StyleValue != sheet.StyleValue)
                columns[count++] = (columnNumber, column.StyleValue);
        }

        Array.Resize(ref columns, count);
        Array.Sort(columns, ColumnNumberComparer.Instance);
        return columns;
    }

    /// <summary>
    /// Each row that is not plain now must be one the baseline has, with the same attributes. Counting
    /// them then shows that none of the baseline's rows has gone plain or been removed.
    /// </summary>
    private bool RowsMatch(XLWorksheet sheet)
    {
        var matched = 0;
        foreach (var (rowNumber, row) in sheet.Internals.RowsCollection)
        {
            var attributes = RowAttributes.Of(rowNumber, row);
            if (attributes.IsPlain(sheet.StyleValue))
                continue;

            var index = Array.BinarySearch(_rows, attributes, RowNumberComparer.Instance);
            if (index < 0 || _rows[index] != attributes)
                return false;

            matched++;
        }

        return matched == _rows.Length;
    }

    /// <inheritdoc cref="RowsMatch"/>
    private bool ColumnsMatch(XLWorksheet sheet)
    {
        var matched = 0;
        foreach (var (columnNumber, column) in sheet.Internals.ColumnsCollection)
        {
            if (column.StyleValue == sheet.StyleValue)
                continue;

            var index = Array.BinarySearch(_columns, (columnNumber, column.StyleValue), ColumnNumberComparer.Instance);
            if (index < 0 || _columns[index].Style != column.StyleValue)
                return false;

            matched++;
        }

        return matched == _columns.Length;
    }

    private static bool HasDirtyFormula(FormulaSlice formulas)
    {
        if (formulas.IsEmpty)
            return false;

        using var enumerator = formulas.GetForwardEnumerator(Area.Full);
        while (enumerator.MoveNext())
        {
            if (enumerator.Current.IsDirty())
                return true;
        }

        return false;
    }

    /// <summary>
    /// The attributes of a row that the sheet-data writer reads, and the style its cells inherit.
    /// <see cref="Height"/> is kept only when it is written, which is when
    /// <see cref="XLRow.HeightChanged"/> is set.
    /// </summary>
    private readonly record struct RowAttributes(
        int Row,
        bool HeightChanged,
        double Height,
        bool Hidden,
        bool Collapsed,
        int OutlineLevel,
        bool ShowPhonetic,
        double? DyDescent,
        XLStyleValue Style)
    {
        internal static RowAttributes Of(int rowNumber, XLRow row) => new(
            rowNumber,
            row.HeightChanged,
            row.HeightChanged ? row.Height : 0,
            row.IsHidden,
            row.Collapsed,
            row.OutlineLevel,
            row.ShowPhonetic,
            row.DyDescent,
            row.StyleValue);

        /// <summary>
        /// Does the row write no attribute and pass on only the sheet's style?
        /// </summary>
        internal bool IsPlain(XLStyleValue sheetStyle) =>
            !HeightChanged && !Hidden && !Collapsed && OutlineLevel == 0 && !ShowPhonetic
            && DyDescent is null && Style == sheetStyle;
    }

    private sealed class RowNumberComparer : IComparer<RowAttributes>
    {
        internal static readonly RowNumberComparer Instance = new();

        public int Compare(RowAttributes x, RowAttributes y) => x.Row.CompareTo(y.Row);
    }

    private sealed class ColumnNumberComparer : IComparer<(int Column, XLStyleValue Style)>
    {
        internal static readonly ColumnNumberComparer Instance = new();

        public int Compare((int Column, XLStyleValue Style) x, (int Column, XLStyleValue Style) y)
            => x.Column.CompareTo(y.Column);
    }
}
