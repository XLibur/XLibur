using System;
using System.Collections.Generic;
using System.Xml;
using XLibur.Excel.Coordinates;
using XLibur.Excel.IO;
using XLibur.Extensions;
using static XLibur.Excel.IO.OpenXmlConst;

namespace XLibur.Excel.Streaming;

/// <summary>
/// A worksheet being written by an <see cref="XLStreamingWorkbook"/>. Rows are appended in
/// ascending order and serialised immediately; nothing already written can be revisited.
/// </summary>
public sealed class XLStreamingWorksheet
{
    private readonly XLStreamingWorkbook _workbook;
    private readonly char[] _cellRef = new char[CellXmlWriter.CellRefBufferLength];
    private readonly Dictionary<(int First, int Last), XLStreamingColumn> _columns = [];

    private XmlWriter? _xml;
    private bool _started;
    private bool _completed;

    private int _nextRowNumber = 1;
    private int _openRowNumber;
    private uint _openRowStyleId;
    private XLStyleValue _openRowStyleValue = XLStyleValue.Default;
    private int _nextColumnNumber = 1;

    private int _freezeRows;
    private int _freezeColumns;

    internal XLStreamingWorksheet(XLStreamingWorkbook workbook, string name, int index)
    {
        _workbook = workbook;
        Name = name;
        Index = index;
    }

    /// <summary>
    /// The zip entry a worksheet's XML lives in. Sheets are numbered from 1 in the order they
    /// were added, which is also their <c>sheetId</c> and the ordinal in their relationship id.
    /// </summary>
    internal static string EntryName(int index) => $"xl/worksheets/sheet{index}.xml";

    /// <summary>The sheet name, as it appears on the tab.</summary>
    public string Name { get; }

    /// <summary>
    /// Range the autofilter dropdowns cover, e.g. <c>"A1:D1"</c>, or <c>null</c> for none.
    /// Written after the rows, so it may be set at any point before <see cref="Complete"/>.
    /// </summary>
    public string? AutoFilterRange { get; set; }

    /// <summary>
    /// Number of the row that will be written by the next <see cref="AddRow()"/> or
    /// <see cref="AppendRow(XLCellValue[])"/>, 1-based.
    /// </summary>
    public int NextRowNumber => _nextRowNumber;

    /// <summary>1-based position of the sheet, used for its part name, id and relationship.</summary>
    internal int Index { get; }

    #region Layout - must be set before the first row

    /// <summary>
    /// Presentation settings for a single column, 1-based. Must be called before the first row
    /// is appended, because columns are written ahead of the rows.
    /// </summary>
    public XLStreamingColumn Column(int columnNumber) => Columns(columnNumber, columnNumber);

    /// <summary>
    /// Presentation settings for an inclusive range of columns, 1-based. Ranges must not
    /// overlap. Must be called before the first row is appended.
    /// </summary>
    public XLStreamingColumn Columns(int firstColumn, int lastColumn)
    {
        ThrowIfStarted(nameof(Columns));
        if (firstColumn < 1 || firstColumn > XLHelper.MaxColumnNumber)
            throw new ArgumentOutOfRangeException(nameof(firstColumn));
        if (lastColumn < firstColumn || lastColumn > XLHelper.MaxColumnNumber)
            throw new ArgumentOutOfRangeException(nameof(lastColumn));

        var key = (firstColumn, lastColumn);
        if (!_columns.TryGetValue(key, out var column))
        {
            column = new XLStreamingColumn(firstColumn, lastColumn);
            _columns.Add(key, column);
        }

        return column;
    }

    /// <summary>
    /// Freeze the top <paramref name="rowCount"/> rows so they stay visible while scrolling.
    /// Must be called before the first row is appended.
    /// </summary>
    public void FreezeRows(int rowCount) => FreezePanes(rowCount, _freezeColumns);

    /// <summary>
    /// Freeze the leftmost <paramref name="columnCount"/> columns. Must be called before the
    /// first row is appended.
    /// </summary>
    public void FreezeColumns(int columnCount) => FreezePanes(_freezeRows, columnCount);

    /// <summary>
    /// Freeze the top <paramref name="rowCount"/> rows and leftmost
    /// <paramref name="columnCount"/> columns. Must be called before the first row is appended.
    /// </summary>
    public void FreezePanes(int rowCount, int columnCount)
    {
        ThrowIfStarted(nameof(FreezePanes));
        ArgumentOutOfRangeException.ThrowIfNegative(rowCount);
        ArgumentOutOfRangeException.ThrowIfNegative(columnCount);

        _freezeRows = rowCount;
        _freezeColumns = columnCount;
    }

    #endregion Layout

    #region Rows

    /// <summary>
    /// Open a row for cell-by-cell writing. The row is closed by disposing the returned value,
    /// or implicitly when the next row starts or the sheet completes.
    /// </summary>
    public XLStreamingRow AddRow() => AddRow(null);

    /// <summary>
    /// Open a row for cell-by-cell writing, with row-level formatting. <paramref name="style"/>
    /// applies to cells in the row that carry no style of their own.
    /// </summary>
    public XLStreamingRow AddRow(IXLStyle? style, double? height = null, bool hidden = false)
    {
        ThrowIfCompleted();
        EnsureStarted();
        EndOpenRow();

        var xml = _xml!;
        var rowNumber = _nextRowNumber;

        if (style is null)
        {
            _openRowStyleValue = XLStyleValue.Default;
            _openRowStyleId = 0;
        }
        else
        {
            var key = XLStyle.GenerateKey(style);
            _openRowStyleValue = XLStyleValue.FromKey(ref key);
            _openRowStyleId = _workbook.Styles.GetOrAdd(_openRowStyleValue);
        }

        CellXmlWriter.WriteRowStart(xml, rowNumber, 0);

        if (height is not null)
        {
            xml.WriteAttribute("ht", height.Value.SaveRound());
            xml.WriteAttributeString("customHeight", TrueValue);
        }

        if (hidden)
            xml.WriteAttributeString("hidden", TrueValue);

        if (_openRowStyleId != 0)
        {
            xml.WriteAttribute("s", _openRowStyleId);
            xml.WriteAttributeString("customFormat", TrueValue);
        }

        _openRowNumber = rowNumber;
        _nextRowNumber = rowNumber + 1;
        _nextColumnNumber = 1;

        return new XLStreamingRow(this, rowNumber);
    }

    /// <summary>
    /// Append a row of values, starting at column A.
    /// </summary>
    public void AppendRow(params XLCellValue[] values)
    {
        ArgumentNullException.ThrowIfNull(values);
        AppendRow(values.AsSpan(), null);
    }

    /// <summary>
    /// Append a row of values, starting at column A, with an optional row-level style.
    /// </summary>
    public void AppendRow(ReadOnlySpan<XLCellValue> values, IXLStyle? style = null)
    {
        var row = AddRow(style);
        foreach (var value in values)
            row.Cell(value);

        EndOpenRow();
    }

    /// <summary>
    /// Leave <paramref name="count"/> rows empty. Skipped rows take no space in the file.
    /// </summary>
    public void SkipRows(int count)
    {
        ThrowIfCompleted();
        ArgumentOutOfRangeException.ThrowIfNegative(count);

        EndOpenRow();
        _nextRowNumber += count;
    }

    /// <summary>
    /// Finish the worksheet and close its part. Repeat calls do nothing. Called automatically
    /// when another worksheet is added or the workbook is finished.
    /// </summary>
    public void Complete()
    {
        if (_completed)
            return;

        EnsureStarted();
        EndOpenRow();

        _xml!.WriteEndElement(); // sheetData

        if (!string.IsNullOrEmpty(AutoFilterRange))
        {
            _xml.WriteStartElement("autoFilter", Main2006SsNs);
            _xml.WriteAttributeString("ref", AutoFilterRange);
            _xml.WriteEndElement();
        }

        _xml.WriteEndElement(); // worksheet
        _xml.WriteEndDocument();
        _xml.Dispose();
        _xml = null;
        _completed = true;
    }

    #endregion Rows

    #region Cell writing, driven by XLStreamingRow

    internal void WriteValueCell(int rowNumber, XLCellValue value, IXLStyle? style)
    {
        var xml = RequireOpenRow(rowNumber);
        var column = _nextColumnNumber++;

        // A date, a duration and a number are all stored as a serial number, so only the number
        // format distinguishes them on the way back in; likewise a leading apostrophe and a line
        // break are carried by the style, not the value. Same rules the cell setter applies.
        var styleValue = ResolveStyleValue(style);
        var adjusted = AdjustStyleForValue(styleValue, ref value);
        var styleId = ReferenceEquals(adjusted, styleValue) && style is null
            ? _openRowStyleId
            : _workbook.Styles.GetOrAdd(adjusted);

        if (value.Type == XLDataType.Blank)
        {
            // A blank cell is only worth writing when it carries formatting the row does not.
            if (styleId == _openRowStyleId)
                return;

            WriteCellStart(xml, rowNumber, column, null, styleId);
            xml.WriteEndElement(); // c
            return;
        }

        var shareStrings = _workbook.Options.StringStorage == XLStreamingStringStorage.SharedStrings;
        WriteCellStart(xml, rowNumber, column, CellXmlWriter.GetValueCellType(value.Type, shareStrings), styleId);

        if (value.Type == XLDataType.Text)
        {
            var text = value.GetText();
            if (shareStrings)
                CellXmlWriter.WriteSharedStringValue(xml, _workbook.SharedStrings.GetOrAdd(text));
            else
                CellXmlWriter.WriteInlineString(xml, text);
        }
        else
        {
            CellXmlWriter.WriteNonTextValue(xml, value, _workbook.Options.Use1904DateSystem);
        }

        xml.WriteEndElement(); // c
    }

    internal void WriteFormulaCell(int rowNumber, string formula, XLCellValue cachedValue, IXLStyle? style)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(formula);

        var xml = RequireOpenRow(rowNumber);
        var column = _nextColumnNumber++;
        var styleId = style is null ? _openRowStyleId : _workbook.Styles.GetOrAdd(ResolveStyleValue(style));

        // Formulas are stored without the leading '=' that a user typically writes.
        if (formula[0] == '=')
            formula = formula[1..];

        var dataType = cachedValue.Type != XLDataType.Blank
            ? CellXmlWriter.GetFormulaCellType(cachedValue.Type)
            : null;

        WriteCellStart(xml, rowNumber, column, dataType, styleId);

        xml.WriteStartElement("f", Main2006SsNs);
        xml.WriteString(formula);
        xml.WriteEndElement(); // f

        switch (cachedValue.Type)
        {
            case XLDataType.Blank:
                break;
            case XLDataType.Text:
                CellXmlWriter.WriteStringValue(xml, cachedValue.GetText());
                break;
            default:
                CellXmlWriter.WriteNonTextValue(xml, cachedValue, _workbook.Options.Use1904DateSystem);
                break;
        }

        xml.WriteEndElement(); // c
    }

    internal void SkipCells(int rowNumber, int count)
    {
        RequireOpenRow(rowNumber);
        ArgumentOutOfRangeException.ThrowIfNegative(count);
        _nextColumnNumber += count;
    }

    internal void MoveToColumn(int rowNumber, int columnNumber)
    {
        RequireOpenRow(rowNumber);
        if (columnNumber < _nextColumnNumber)
        {
            throw new ArgumentOutOfRangeException(nameof(columnNumber),
                $"Cells are written left to right; column {columnNumber} is before the next free column " +
                $"{_nextColumnNumber} of row {rowNumber}.");
        }

        _nextColumnNumber = columnNumber;
    }

    internal void EndRow(int rowNumber)
    {
        if (_openRowNumber == rowNumber)
            EndOpenRow();
    }

    /// <summary>
    /// The interned value of an explicit cell style, or the open row's style when there is none.
    /// </summary>
    /// <remarks>
    /// Not cached against the last <see cref="IXLStyle"/> instance seen: a style handed out by
    /// <see cref="XLStreamingWorkbook.CreateStyle"/> is documented as reusable, so the same
    /// instance legitimately carries different values from one cell to the next.
    /// </remarks>
    private XLStyleValue ResolveStyleValue(IXLStyle? style)
    {
        if (style is null)
            return _openRowStyleValue;

        var key = XLStyle.GenerateKey(style);
        return XLStyleValue.FromKey(ref key);
    }

    /// <summary>
    /// Apply the value-driven style adjustments, and strip the quote prefix from text once it
    /// has been captured in the style. Returns <paramref name="styleValue"/> unchanged when the
    /// value needs nothing.
    /// </summary>
    private static XLStyleValue AdjustStyleForValue(XLStyleValue styleValue, ref XLCellValue value)
    {
        switch (value.Type)
        {
            case XLDataType.DateTime:
                return XLValueStyleRules.HasGeneralNumberFormat(styleValue)
                    ? XLValueStyleRules.WithDateTimeFormat(styleValue, value.GetUnifiedNumber() % 1 == 0)
                    : styleValue;

            case XLDataType.TimeSpan:
                return XLValueStyleRules.HasGeneralNumberFormat(styleValue)
                    ? XLValueStyleRules.WithDurationFormat(styleValue)
                    : styleValue;

            case XLDataType.Text:
                {
                    var text = value.GetText();
                    var adjusted = XLValueStyleRules.AdjustForText(styleValue, text);
                    if (text.Length > 0 && text[0] == '\'')
                        value = text[1..];

                    return adjusted ?? styleValue;
                }

            default:
                return styleValue;
        }
    }

    private void WriteCellStart(XmlWriter xml, int rowNumber, int columnNumber, string? dataType, uint styleId)
    {
        if (columnNumber > XLHelper.MaxColumnNumber)
        {
            throw new InvalidOperationException(
                $"Row {rowNumber} of '{Name}' exceeds the {XLHelper.MaxColumnNumber} column limit.");
        }

        var length = new Point(rowNumber, columnNumber).Format(_cellRef);
        CellXmlWriter.WriteCellStart(xml, _cellRef, length, dataType, styleId);
    }

    private XmlWriter RequireOpenRow(int rowNumber)
    {
        if (_openRowNumber != rowNumber || _xml is null)
        {
            throw new InvalidOperationException(
                $"Row {rowNumber} of '{Name}' is no longer being written. A row can only be written to until " +
                "the next row starts or the worksheet completes.");
        }

        return _xml;
    }

    private void EndOpenRow()
    {
        if (_openRowNumber == 0)
            return;

        _xml!.WriteEndElement(); // row
        _openRowNumber = 0;
        _openRowStyleId = 0;
        _openRowStyleValue = XLStyleValue.Default;
    }

    #endregion Cell writing

    #region Part header

    /// <summary>
    /// Open the part and write everything that precedes <c>sheetData</c>. Deferred to the first
    /// row so the caller can still configure columns and panes after <c>AddWorksheet</c>.
    /// </summary>
    private void EnsureStarted()
    {
        if (_started)
            return;

        _started = true;
        _xml = _workbook.CreatePart(EntryName(Index));

        _xml.WriteStartDocument(true);
        _xml.WriteStartElement("worksheet", Main2006SsNs);

        WriteSheetViews(_xml);
        WriteColumns(_xml);

        _xml.WriteStartElement("sheetData", Main2006SsNs);
    }

    private void WriteSheetViews(XmlWriter xml)
    {
        xml.WriteStartElement("sheetViews", Main2006SsNs);
        xml.WriteStartElement("sheetView", Main2006SsNs);
        xml.WriteAttribute("workbookViewId", 0u);

        if (_freezeRows > 0 || _freezeColumns > 0)
        {
            xml.WriteStartElement("pane", Main2006SsNs);

            if (_freezeColumns > 0)
                xml.WriteAttribute("xSplit", _freezeColumns);

            if (_freezeRows > 0)
                xml.WriteAttribute("ySplit", _freezeRows);

            xml.WriteAttributeString("topLeftCell", new Point(_freezeRows + 1, _freezeColumns + 1).ToString());
            xml.WriteAttributeString("activePane", ResolveActivePane());
            xml.WriteAttributeString("state", "frozen");

            xml.WriteEndElement(); // pane
        }

        xml.WriteEndElement(); // sheetView
        xml.WriteEndElement(); // sheetViews
    }

    private string ResolveActivePane()
    {
        if (_freezeRows > 0 && _freezeColumns > 0)
            return "bottomRight";

        return _freezeRows > 0 ? "bottomLeft" : "topRight";
    }

    private void WriteColumns(XmlWriter xml)
    {
        if (_columns.Count == 0)
            return;

        var ordered = new List<XLStreamingColumn>(_columns.Values);
        ordered.Sort(static (a, b) => a.FirstColumn.CompareTo(b.FirstColumn));

        xml.WriteStartElement("cols", Main2006SsNs);
        foreach (var column in ordered)
        {
            xml.WriteStartElement("col", Main2006SsNs);
            xml.WriteAttribute("min", (uint)column.FirstColumn);
            xml.WriteAttribute("max", (uint)column.LastColumn);

            if (column.Width is not null)
            {
                xml.WriteAttribute("width", ColumnWriter.GetColumnWidth(column.Width.Value).SaveRound());
                xml.WriteAttributeString("customWidth", TrueValue);
            }

            if (column.Style is not null)
                xml.WriteAttribute("style", _workbook.Styles.GetOrAdd(column.Style));

            if (column.Hidden)
                xml.WriteAttributeString("hidden", TrueValue);

            if (column.OutlineLevel > 0)
                xml.WriteAttribute("outlineLevel", column.OutlineLevel);

            if (column.Collapsed)
                xml.WriteAttributeString("collapsed", TrueValue);

            xml.WriteEndElement(); // col
        }

        xml.WriteEndElement(); // cols
    }

    private void ThrowIfStarted(string member)
    {
        if (_started)
        {
            throw new InvalidOperationException(
                $"{member} affects XML written before the rows, so it must be set before the first row is " +
                $"appended to '{Name}'.");
        }
    }

    private void ThrowIfCompleted()
    {
        if (_completed)
            throw new InvalidOperationException($"Worksheet '{Name}' has been completed.");
    }

    #endregion Part header
}
