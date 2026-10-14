using System;
using System.Collections.Generic;
using XLibur.Excel.Coordinates;
using XLibur.Excel.Rows;

namespace XLibur.Excel;

/// <summary>
/// What a worksheet's cells were as the file had them, kept so a save can tell whether the sheet's
/// <c>&lt;sheetData&gt;</c> would come out as the file had it (#702).
/// </summary>
/// <remarks>
/// It is captured as soon as the load has read the cells. Later steps of the load can write to a
/// cell: loading a table names its empty header cells after the fields. Those count as changes, since
/// the file's cells do not hold them. The misc slice is the exception. Comments, threaded comments
/// and in-cell images are loaded into it after the cells. The first two are not written into the
/// cells' markup, and an image is written as the <c>vm</c> index its cell was read with, so the
/// version is taken again when the load ends (<see cref="RecordAnnotations"/>).
/// <para>
/// It errs towards reporting a change: setting a cell to the value it already holds counts, and so
/// does materialising a column, which gives the cells of each existing row a style of their own. The
/// aim is never to miss a change. It records what the sheet-data writer reads:
/// <list type="bullet">
/// <item>The <see cref="ISlice.Version"/> of each cell slice. <see cref="FormulaSlice.Version"/> also
/// counts formulas edited in place, such as by a sheet rename or an array formula shifted by an insert
/// on another sheet. The value slice's version counts a recalculated cached value only when the value
/// differs (<see cref="ValueSlice.SetFormulaResult"/>), so a recalculation that gives the values the
/// file had is no change (#720).</item>
/// <item>The attributes the writer gives each row. Rows have many setters and no change counter, so
/// they are compared as they are, not through a version.</item>
/// <item>The styles a cell with no style of its own inherits: the sheet's, and those of its row and
/// column.</item>
/// <item>The workbook's date system, which sets the serial number a date is written as.</item>
/// </list>
/// Formulas get one more check, at save. A formula that is dirty then and holds a cached value has
/// that value recalculated or dropped by the save. Editing another sheet can make it dirty without
/// writing to this one, and so can a structural edit anywhere in the workbook, which marks every
/// formula dirty. A formula that is dirty with no cached value is written as the file had it, without
/// one.
/// </para>
/// <para>
/// Not tracked: labels in a table's totals row, which the writer takes from the table, so the save
/// checks the tables as it finds them instead (<c>SheetDataWriter.WritesTotalsRowsAsValues</c>); and,
/// because the tiers of #702 exclude these sheets, the <c>cm</c> index the save gives a dynamic-array formula
/// that has none, which is workbook-wide (a sheet with a formula is excluded). A <c>cm</c> index
/// the load read is in the misc slice and written as loaded. A sparkline added to a blank cell can add an empty <c>&lt;row&gt;</c> element and
/// is not tracked either; it changes no cell.
/// </para>
/// </remarks>
internal sealed class SheetDataBaseline
{
    private readonly int _valueVersion;
    private readonly int _formulaVersion;
    private readonly int _styleVersion;
    private int _miscVersion;
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
    /// Takes in the comments, threaded comments and in-cell images the load added to the misc slice
    /// after the cells were captured. A cell image is written into the cells as <c>vm</c> and a rich
    /// value index; the save sets both before it writes a sheet, so a change it makes to either shows
    /// up as a change here.
    /// </summary>
    internal void RecordAnnotations(XLWorksheet sheet) => _miscVersion = sheet.Internals.CellsCollection.MiscSlice.Version;

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

        return RowsMatch(sheet) && ColumnsMatch(sheet) && !HasStaleCachedValue(cells);
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

    /// <summary>
    /// Does a dirty formula hold a cached value? The save would recalculate that value or write the
    /// formula without it, where the file has it.
    /// </summary>
    /// <remarks>
    /// Called once the slice versions match, so each cell holds the value it held when the baseline
    /// was captured. A formula that is dirty with no value had none then either, as the loader
    /// leaves a formula whose cell has no <c>&lt;v&gt;</c>, and the save writes it without one too.
    /// A save that calculates it first writes a value, which changes the value slice's version.
    /// </remarks>
    private static bool HasStaleCachedValue(XLCellsCollection cells)
    {
        var formulas = cells.FormulaSlice;
        if (formulas.IsEmpty)
            return false;

        using var enumerator = formulas.GetForwardEnumerator(Area.Full);
        while (enumerator.MoveNext())
        {
            if (enumerator.Current.IsDirty() && cells.ValueSlice.GetCellValue(enumerator.Point).Type != XLDataType.Blank)
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
