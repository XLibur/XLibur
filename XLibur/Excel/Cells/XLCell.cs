using XLibur.Excel.InsertData;
using XLibur.Extensions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using XLibur.Excel.Drawings;
using XLibur.Graphics;
using ClosedXML.Parser;
using XLibur.Excel.CalcEngine.Visitors;
using XLibur.Excel.Coordinates;
using XLibur.Excel.RichText;
using XLibur.Excel.Tables;

namespace XLibur.Excel;

[DebuggerDisplay("{Address}")]
internal sealed class XLCell : XLStylizedBase, IXLCell, IXLStylized
{
    private readonly XLCellsCollection _cellsCollection;

    private readonly XLSheetPoint _point;

    internal XLCell(XLWorksheet worksheet, int row, int column)
        : this(worksheet, new XLSheetPoint(row, column))
    {
    }

    internal XLCell(XLWorksheet worksheet, XLSheetPoint point)
    {
        _cellsCollection = worksheet.Internals.CellsCollection;
        _point = point;
    }

    public XLWorksheet Worksheet => _cellsCollection.Worksheet;

    public XLAddress Address => new(Worksheet, _point.Row, _point.Column, false, false);

    internal XLSheetPoint SheetPoint => _point;

    #region Slice fields

    /// <summary>
    /// A flag indicating if a string should be stored in the shared table or inline.
    /// </summary>
    public bool ShareString
    {
        get => _cellsCollection.ValueSlice.GetShareString(SheetPoint);
        set => _cellsCollection.ValueSlice.SetShareString(SheetPoint, value);
    }

    /// <summary>
    /// Gets the effective style by resolving inheritance (cell → row → column → worksheet).
    /// Sets an explicit style on the cell, overriding inheritance.
    /// </summary>
    internal override XLStyleValue StyleValue
    {
        get => GetInheritedStyle();
        private protected set => _cellsCollection.StyleSlice.Set(_point, value);
    }

    internal int MemorySstId => _cellsCollection.ValueSlice.GetShareStringId(SheetPoint);

    internal XLImmutableRichText? RichText => SliceRichText;

    internal XLCellValue SliceCellValue
    {
        get => _cellsCollection.ValueSlice.GetCellValue(SheetPoint);
        set
        {
            _cellsCollection.ValueSlice.SetCellValue(SheetPoint, value);
            Worksheet.Workbook.CalcEngine.MarkDirty(Worksheet, SheetPoint);
        }
    }

    internal XLImmutableRichText? SliceRichText
    {
        get => _cellsCollection.ValueSlice.GetRichText(SheetPoint);
        set => _cellsCollection.ValueSlice.SetRichText(SheetPoint, value!);
    }

    internal XLComment? SliceComment
    {
        get => _cellsCollection.MiscSlice[_point].Comment;
        set
        {
            ref readonly var original = ref _cellsCollection.MiscSlice[_point];
            if (original.Comment != value)
            {
                var modified = original;
                modified.Comment = value;
                _cellsCollection.MiscSlice.Set(_point, in modified);
            }
        }
    }

    internal uint? CellMetaIndex
    {
        get => _cellsCollection.MiscSlice[_point].CellMetaIndex;
        set
        {
            ref readonly var original = ref _cellsCollection.MiscSlice[_point];
            if (original.CellMetaIndex != value)
            {
                var modified = original;
                modified.CellMetaIndex = value;
                _cellsCollection.MiscSlice.Set(_point, in modified);
            }
        }
    }

    internal uint? ValueMetaIndex
    {
        get => _cellsCollection.MiscSlice[_point].ValueMetaIndex;
        set
        {
            ref readonly var original = ref _cellsCollection.MiscSlice[_point];
            if (original.ValueMetaIndex != value)
            {
                var modified = original;
                modified.ValueMetaIndex = value;
                _cellsCollection.MiscSlice.Set(_point, in modified);
            }
        }
    }

    internal XLCellImage? CellImage
    {
        get => _cellsCollection.MiscSlice[_point].CellImage;
        set
        {
            ref readonly var original = ref _cellsCollection.MiscSlice[_point];
            if (original.CellImage != value)
            {
                var modified = original;
                modified.CellImage = value;
                _cellsCollection.MiscSlice.Set(_point, in modified);
            }
        }
    }

    /// <summary>
    /// A formula in the cell. Null, if cell doesn't contain formula.
    /// </summary>
    internal XLCellFormula? Formula
    {
        get => _cellsCollection.FormulaSlice.Get(SheetPoint);
        set
        {
            _cellsCollection.FormulaSlice.Set(SheetPoint, value);

            // Because text values of evaluated formulas are stored in a worksheet part, mark it as inlined string and store in sst.
            // If we are clearing formula, we should enable shareString back on, because it is a default position.
            // If we are setting formula, we should disable shareString (=inline), because it must be written to the worksheet part
            var clearFormula = value is null;
            ShareString = clearFormula;
            Worksheet.Workbook.CalcEngine.MarkDirty(Worksheet, SheetPoint);
        }
    }

    #endregion Slice fields

    internal XLComment GetComment()
    {
        return SliceComment ?? CreateComment();
    }

    internal XLComment CreateComment(int? shapeId = null)
    {
        return SliceComment = new XLComment(this, shapeId: shapeId);
    }

    public XLRichText GetRichText()
    {
        var sliceRichText = SliceRichText;
        if (sliceRichText is not null)
            return new XLRichText(this, sliceRichText);

        return CreateRichText();
    }

    public XLRichText CreateRichText()
    {
        var font = new XLFont(GetStyleForRead().Font.Key);

        // Don't include rich text string with 0 length to a new rich text
        var richText = DataType == XLDataType.Blank
            ? new XLRichText(this, font)
            : new XLRichText(this, GetFormattedString(), font);
        SliceRichText = XLImmutableRichText.Create(richText);
        return richText;
    }

    #region IXLCell Members

    IXLWorksheet IXLCell.Worksheet => Worksheet;

    IXLAddress IXLCell.Address => Address;

    IXLRange IXLCell.AsRange()
    {
        return AsRange();
    }

    internal IXLCell SetValue(XLCellValue value, bool setTableHeader, bool checkMergedRanges)
    {
        if (checkMergedRanges && IsInferiorMergedCell())
            return this;

        var point = _point;

        SetValueAndStyle(value, point);

        // Only clear formula if cell actually has one, to avoid unnecessary
        // property setter overhead (TrimFormulaEqual, InvalidateFormula, etc.)
        if (_cellsCollection.FormulaSlice.Get(point) is not null)
            FormulaA1 = string.Empty;

        if (setTableHeader && Worksheet.Tables.Count > 0)
        {
            var cellRange = new XLSheetRange(point, point);
            foreach (var table in Worksheet.Tables)
                table.RefreshFieldsFromCells(cellRange);
        }

        return this;
    }

    /// <summary>
    /// Set value of a cell and its format (if necessary) from the passed value.
    /// It doesn't clear formulas or checks merged cells or tables.
    /// </summary>
    private void SetValueAndStyle(XLCellValue value, XLSheetPoint point)
    {
        var modifiedStyleValue = Worksheet.GetStyleForValue(value, point);
        if (modifiedStyleValue is not null)
            StyleValue = modifiedStyleValue;

        // Modify value after style, because we might strip the '
        if (value.Type == XLDataType.Text)
        {
            var text = value.GetText();
            if (text.Length > 0 && text[0] == '\'')
            {
                value = text.Substring(1);
            }
        }

        _cellsCollection.ValueSlice.SetCellValue(point, value);
        Worksheet.Workbook.CalcEngine.MarkDirty(Worksheet, point);
    }

    public bool GetBoolean() => Value.GetBoolean();

    public double GetDouble() => Value.GetNumber();

    public string GetText() => Value.GetText();

    public XLError GetError() => Value.GetError();

    public DateTime GetDateTime() => Value.GetDateTime();

    public TimeSpan GetTimeSpan() => Value.GetTimeSpan();

    public bool TryGetValue<T>(out T value)
    {
        XLCellValue currentValue;
        try
        {
            currentValue = Value;
        }
        catch
        {
            // May fail for formula evaluation
            value = default!;
            return false;
        }

        return XLCellValueConverter.TryConvert(currentValue, out value);
    }

    public T GetValue<T>()
    {
        if (TryGetValue(out T retVal))
            return retVal;

        throw new InvalidCastException($"Cannot convert {Address.ToStringRelative(true)}'s value to " + typeof(T));
    }

    public string GetString() => Value.ToString(CultureInfo.CurrentCulture);

    public string GetFormattedString(CultureInfo? culture = null)
    {
        XLCellValue value;
        try
        {
            // Need to get actual value because formula might be out of date or value wasn't set at all
            // Unimplemented functions and features throw exceptions
            value = Value;
        }
        catch
        {
            value = CachedValue;
        }

        return GetFormattedString(value, culture);
    }

    internal string GetFormattedString(XLCellValue value, CultureInfo? culture = null)
    {
        culture ??= CultureInfo.CurrentCulture;
        var format = GetFormat();
        return value.IsUnifiedNumber
            ? value.GetUnifiedNumber().ToExcelFormat(format, culture)
            : value.ToString(culture);
    }

    public void InvalidateFormula()
    {
        if (Formula is null)
        {
            return;
        }

        Formula.IsDirty = true;
    }

    /// <summary>
    /// Perform an evaluation of cell formula. If cell does not contain formula nothing happens, if cell does not need
    /// recalculation (<see cref="NeedsRecalculation"/> is False) nothing happens either, unless <paramref name="force"/> flag is specified.
    /// Otherwise recalculation is performed, result value is preserved in <see cref="CachedValue"/> and returned.
    /// </summary>
    /// <param name="force">Flag indicating whether a recalculation must be performed even is cell does not need it.</param>
    /// <returns>Null if cell does not contain a formula. Calculated value otherwise.</returns>
    public void Evaluate(bool force)
    {
        if (Formula is null)
        {
            return;
        }

        var shouldRecalculate = force || NeedsRecalculation;
        if (!shouldRecalculate)
        {
            return;
        }

        var wb = Worksheet.Workbook;
        if (force || !wb.CalcEngine.TryEvaluateSingleCell(Formula, SheetPoint, Worksheet))
        {
            wb.CalcEngine.Recalculate(wb, null);
        }
    }

    /// <summary>
    /// Set only value, don't clear formula, don't set format.
    /// Sets the value even for merged cells.
    /// </summary>
    internal void SetOnlyValue(XLCellValue value)
    {
        SliceCellValue = value;
    }

    public IXLCell SetValue(XLCellValue value)
    {
        return SetValue(value, true, true);
    }

    public override string ToString() => ToString("A");

    public string ToString(string format)
    {
        return (format.ToUpper()) switch
        {
            "A" => Address.ToString(),
            "F" => HasFormula ? FormulaA1 : string.Empty,
            "NF" => Style.NumberFormat.Format,
            "FG" => Style.Font.FontColor.ToString(),
            "BG" => Style.Fill.BackgroundColor.ToString(),
            "V" => GetFormattedString(),
            _ => throw new FormatException($"Format {format} was not recognised."),
        };
    }

    public XLCellValue Value
    {
        get
        {
            if (Formula is not null)
            {
                Evaluate(false);
            }

            return SliceCellValue;
        }
        set => SetValue(value);
    }

    public IXLTable InsertTable<T>(IEnumerable<T> data)
    {
        return InsertTable(data, null, true);
    }

    public IXLTable InsertTable<T>(IEnumerable<T> data, bool createTable)
    {
        return InsertTable(data, null, createTable);
    }

    public IXLTable InsertTable<T>(IEnumerable<T> data, string tableName)
    {
        return InsertTable(data, tableName, true);
    }

    public IXLTable InsertTable<T>(IEnumerable<T> data, string? tableName, bool createTable)
    {
        return InsertTable(data, tableName, createTable, addHeadings: true, transpose: false);
    }

    public IXLTable InsertTable<T>(IEnumerable<T> data, string? tableName, bool createTable, bool addHeadings,
        bool transpose)
    {
        var reader = InsertDataReaderFactory.CreateReader(data);
        return Worksheet.InsertTable(SheetPoint, reader, tableName, createTable, addHeadings, transpose);
    }

    public IXLTable? InsertTable(DataTable data)
    {
        return InsertTable(data, null, true);
    }

    public IXLTable? InsertTable(DataTable data, bool createTable)
    {
        return InsertTable(data, null, createTable);
    }

    public IXLTable? InsertTable(DataTable data, string tableName)
    {
        return InsertTable(data, tableName, true);
    }

    public IXLTable? InsertTable(DataTable data, string? tableName, bool createTable)
    {
        if (data == null || data.Columns.Count == 0)
            return null;

        if (createTable && Worksheet.Tables.Any<XLTable>(t => t.Contains(this)))
            throw new InvalidOperationException($"This cell '{Address}' is already part of a table.");

        var reader = InsertDataReaderFactory.CreateReader(data);
        return Worksheet.InsertTable(SheetPoint, reader, tableName, createTable, addHeadings: true, transpose: false);
    }

    public XLTableCellType TableCellType()
    {
        var table = Worksheet.Tables.FirstOrDefault<XLTable>(t => t.AsRange().Contains(this));
        if (table == null) return XLTableCellType.None;

        if (table.ShowHeaderRow && table.HeadersRow()!.RowNumber().Equals(_point.Row)) return XLTableCellType.Header;
        if (table.ShowTotalsRow && table.TotalsRow()!.RowNumber().Equals(_point.Row)) return XLTableCellType.Total;

        return XLTableCellType.Data;
    }

    public IXLRange? InsertData(IEnumerable data)
    {
        return data is null or string ? null : InsertData(data, transpose: false);
    }

    public IXLRange? InsertData(IEnumerable data, bool transpose)
    {
        if (data is null or string)
            return null;

        var reader = InsertDataReaderFactory.CreateReader(data);
        return Worksheet.InsertData(SheetPoint, reader, addHeadings: false, transpose: transpose);
    }

    public IXLRange? InsertData(DataTable dataTable)
    {
        if (dataTable == null)
            return null;

        var reader = InsertDataReaderFactory.CreateReader(dataTable);
        return Worksheet.InsertData(SheetPoint, reader, addHeadings: false, transpose: false);
    }

    public XLDataType DataType => SliceCellValue.Type;

    public IXLCell Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        return Clear(clearOptions, false);
    }

    internal IXLCell Clear(XLClearOptions clearOptions, bool calledFromRange)
    {
        if (!calledFromRange && IsMerged())
        {
            var firstOrDefault = Worksheet.Internals.MergedRanges.GetIntersectedRanges(Address).FirstOrDefault();
            firstOrDefault?.Clear(clearOptions);
        }
        else
        {
            ClearCellContent(clearOptions);
        }

        return this;
    }

    private void ClearCellContent(XLClearOptions clearOptions)
    {
        if (clearOptions.HasFlag(XLClearOptions.Contents))
        {
            SetHyperlink(null);
            SliceCellValue = Blank.Value;
            FormulaA1 = string.Empty;
            CellImage = null;
        }

        if (clearOptions.HasFlag(XLClearOptions.NormalFormats))
            SetStyle(Worksheet.Style);

        if (clearOptions.HasFlag(XLClearOptions.ConditionalFormats))
            AsRange().RemoveConditionalFormatting();

        if (clearOptions.HasFlag(XLClearOptions.Comments))
            SliceComment = null;

        if (clearOptions.HasFlag(XLClearOptions.Sparklines))
            AsRange().RemoveSparklines();

        if (clearOptions.HasFlag(XLClearOptions.DataValidation) && HasDataValidation)
        {
            var validation = CreateDataValidation();
            Worksheet.DataValidations.Delete(validation);
        }

        if (clearOptions.HasFlag(XLClearOptions.MergedRanges) && IsMerged())
            ClearMerged();
    }

    public void Delete(XLShiftDeletedCells shiftDeleteCells)
    {
        Worksheet.Range(Address, Address).Delete(shiftDeleteCells);
    }

    public string FormulaA1
    {
        get => Formula?.A1 ?? string.Empty;

        set
        {
            if (IsInferiorMergedCell())
                return;

            var formula = value.TrimFormulaEqual();
            if (!string.IsNullOrWhiteSpace(formula))
            {
                var fixedFunctionsFormula =
                    FormulaTransformation.FixFutureFunctions(formula, Worksheet.Name, SheetPoint);
                Formula = XLCellFormula.NormalA1(fixedFunctionsFormula);
            }
            else
            {
                Formula = null;
            }

            InvalidateFormula();
        }
    }

    public string FormulaR1C1
    {
        get => Formula?.GetFormulaR1C1(SheetPoint) ?? string.Empty;

        set
        {
            if (IsInferiorMergedCell())
                return;

            var formula = value.TrimFormulaEqual();
            if (!string.IsNullOrWhiteSpace(formula))
            {
                var formulaA1 = FormulaTransformation.SafeToA1(formula, _point.Row, _point.Column);
                var fixedFunctionsFormulaA1 =
                    FormulaTransformation.FixFutureFunctions(formulaA1, Worksheet.Name, SheetPoint);
                Formula = XLCellFormula.NormalA1(fixedFunctionsFormulaA1);
            }
            else
            {
                Formula = null;
            }

            InvalidateFormula();
        }
    }

    public XLHyperlink GetHyperlink()
    {
        if (Worksheet.Hyperlinks.TryGet(SheetPoint, out var hyperlink))
            return hyperlink;

        return CreateHyperlink();
    }

    /// <inheritdoc />
    public void SetHyperlink(XLHyperlink? hyperlink)
    {
        if (Worksheet.Hyperlinks.TryGet(SheetPoint, out var existingHyperlink))
            Worksheet.Hyperlinks.Delete(existingHyperlink);

        if (hyperlink is null)
            return;

        Worksheet.Hyperlinks.Add(SheetPoint, hyperlink);

        if (GetStyleForRead().Font.FontColor.Equals(Worksheet.StyleValue.Font.FontColor))
            Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Hyperlink);

        if (GetStyleForRead().Font.Underline == Worksheet.StyleValue.Font.Underline)
            Style.Font.Underline = XLFontUnderlineValues.Single;
    }

    internal void SetCellHyperlink(XLHyperlink hyperlink)
    {
        Worksheet.Hyperlinks.Clear(SheetPoint);
        Worksheet.Hyperlinks.Add(SheetPoint, hyperlink);
    }

    public XLHyperlink CreateHyperlink()
    {
        SetHyperlink(new XLHyperlink());
        return GetHyperlink();
    }

    public IXLCells InsertCellsAbove(int numberOfRows)
    {
        return AsRange().InsertRowsAbove(numberOfRows).Cells();
    }

    public IXLCells InsertCellsBelow(int numberOfRows)
    {
        return AsRange().InsertRowsBelow(numberOfRows).Cells();
    }

    public IXLCells InsertCellsAfter(int numberOfColumns)
    {
        return AsRange().InsertColumnsAfter(numberOfColumns).Cells();
    }

    public IXLCells InsertCellsBefore(int numberOfColumns)
    {
        return AsRange().InsertColumnsBefore(numberOfColumns).Cells();
    }

    public IXLCell AddToNamed(string rangeName)
    {
        AsRange().AddToNamed(rangeName);
        return this;
    }

    public IXLCell AddToNamed(string rangeName, XLScope scope)
    {
        AsRange().AddToNamed(rangeName, scope);
        return this;
    }

    public IXLCell AddToNamed(string rangeName, XLScope scope, string comment)
    {
        AsRange().AddToNamed(rangeName, scope, comment);
        return this;
    }

    /// <summary>
    /// Flag indicating that previously calculated cell value may be not valid anymore and has to be re-evaluated.
    /// </summary>
    public bool NeedsRecalculation => Formula is not null && Formula.IsDirty;

    public XLCellValue CachedValue => SliceCellValue;

    IXLRichText IXLCell.GetRichText() => GetRichText();

    public bool HasRichText => SliceRichText is not null;

    IXLRichText IXLCell.CreateRichText() => CreateRichText();

    IXLComment IXLCell.GetComment() => GetComment();

    public bool HasComment
    {
        get { return SliceComment != null; }
    }

    IXLComment IXLCell.CreateComment()
    {
        return CreateComment(shapeId: null);
    }

    public bool IsMerged()
    {
        return Worksheet.Internals.MergedRanges.Contains(this);
    }

    public IXLRange? MergedRange()
    {
        return Worksheet
            .Internals
            .MergedRanges
            .GetIntersectedRanges(this)
            .FirstOrDefault();
    }

    public bool IsEmpty()
    {
        return IsEmpty(XLCellsUsedOptions.AllContents);
    }

    public bool IsEmpty(XLCellsUsedOptions options)
    {
        if (!IsContentEmpty())
            return false;

        if (options.HasFlag(XLCellsUsedOptions.NormalFormats) && !IsFormatEmpty())
            return false;

        if (options.HasFlag(XLCellsUsedOptions.MergedRanges) && IsMerged())
            return false;

        if (options.HasFlag(XLCellsUsedOptions.Comments) && HasComment)
            return false;

        if (options.HasFlag(XLCellsUsedOptions.DataValidation) && HasDataValidation)
            return false;

        if (options.HasFlag(XLCellsUsedOptions.ConditionalFormats)
            && Worksheet.ConditionalFormats.SelectMany(cf => cf.Ranges).Any(range => range.Contains(this)))
            return false;

        if (options.HasFlag(XLCellsUsedOptions.Sparklines) && HasSparkline)
            return false;

        return true;
    }

    private bool IsContentEmpty()
    {
        if (CellImage is not null)
            return false;

        if (HasFormula)
            return false;

        return SliceCellValue.Type switch
        {
            XLDataType.Blank => true,
            XLDataType.Text => SliceCellValue.GetText().Length == 0,
            _ => false
        };
    }

    private bool IsFormatEmpty()
    {
        if (StyleValue.IncludeQuotePrefix)
            return false;

        if (!StyleValue.Equals(Worksheet.StyleValue))
            return false;

        if (Worksheet.Internals.RowsCollection.TryGetValue(_point.Row, out var row) &&
            !row.StyleValue.Equals(Worksheet.StyleValue))
            return false;

        if (Worksheet.Internals.ColumnsCollection.TryGetValue(_point.Column, out var column) &&
            !column.StyleValue.Equals(Worksheet.StyleValue))
            return false;

        return true;
    }

    public IXLColumn WorksheetColumn()
    {
        return Worksheet.Column(_point.Column);
    }

    public IXLRow WorksheetRow()
    {
        return Worksheet.Row(_point.Row);
    }

    public IXLCell CopyTo(IXLCell target)
    {
        ((XLCell)target).CopyFrom(this, XLCellCopyOptions.All);
        return target;
    }

    public IXLCell CopyTo(string target)
    {
        return CopyTo(XLCellCopyHelper.GetTargetCell(target, Worksheet));
    }

    public IXLCell CopyFrom(IXLCell otherCell)
    {
        return CopyFrom((XLCell)otherCell, XLCellCopyOptions.All);
    }

    public IXLCell CopyFrom(string otherCell)
    {
        return CopyFrom(XLCellCopyHelper.GetTargetCell(otherCell, Worksheet));
    }

    public IXLCell SetFormulaA1(string formula)
    {
        FormulaA1 = formula;
        return this;
    }

    public IXLCell SetDynamicFormulaA1(string formula)
    {
        if (IsInferiorMergedCell())
            return this;

        var trimmed = formula.TrimFormulaEqual();
        if (!string.IsNullOrWhiteSpace(trimmed))
        {
            var fixedFunctionsFormula =
                FormulaTransformation.FixFutureFunctions(trimmed, Worksheet.Name, SheetPoint);
            Formula = XLCellFormula.DynamicArrayA1(fixedFunctionsFormula);
        }
        else
        {
            Formula = null;
        }

        InvalidateFormula();
        return this;
    }

    public IXLCell SetFormulaR1C1(string formula)
    {
        FormulaR1C1 = formula;
        return this;
    }

    public bool HasSparkline => Sparkline != null;

    /// <summary> The sparkline assigned to the cell </summary>
    public IXLSparkline? Sparkline => Worksheet.SparklineGroups.GetSparkline(this);

    public bool HasCellImage => CellImage is not null;

    public IXLCell SetCellImage(Stream imageStream, XLPictureFormat format, string altText = "")
    {
        var store = Worksheet.Workbook.InCellImages;
        var imageIndex = store.Add(imageStream, format);
        CellImage = new XLCellImage(imageIndex, altText);
        SliceCellValue = Blank.Value;
        FormulaA1 = string.Empty;
        return this;
    }

    public void RemoveCellImage()
    {
        CellImage = null;
    }

    public IXLDataValidation GetDataValidation()
    {
        return FindDataValidation() ?? CreateDataValidation();
    }

    public bool HasDataValidation => FindDataValidation() != null;

    /// <summary>
    /// Get the data validation rule containing the current cell.
    /// </summary>
    /// <returns>The data validation rule applying to the current cell or null if there is no such rule.</returns>
    private IXLDataValidation? FindDataValidation()
    {
        Worksheet.DataValidations.TryGet(new XLRangeAddress(Address, Address), out var dataValidation);
        return dataValidation;
    }

    public IXLDataValidation CreateDataValidation()
    {
        var validation = new XLDataValidation(AsRange());
        Worksheet.DataValidations.Add(validation);
        return validation;
    }

    [Obsolete("Use GetDataValidation() to access the existing rule, or CreateDataValidation() to create a new one.")]
    public IXLDataValidation SetDataValidation()
    {
        return GetDataValidation();
    }

    public void Select()
    {
        AsRange().Select();
    }

    public IXLConditionalFormat AddConditionalFormat()
    {
        return AsRange().AddConditionalFormat();
    }

    public bool Active
    {
        get => Worksheet.ActiveCell == SheetPoint;
        set
        {
            if (value)
                Worksheet.ActiveCell = SheetPoint;
            else if (Active)
                Worksheet.ActiveCell = null;
        }
    }

    public IXLCell SetActive(bool value = true)
    {
        Active = value;
        return this;
    }

    public bool HasHyperlink => Worksheet.Hyperlinks.TryGet(SheetPoint, out _);

    /// <inheritdoc />
    public bool ShowPhonetic
    {
        get => _cellsCollection.MiscSlice[_point].HasPhonetic;
        set
        {
            ref readonly var original = ref _cellsCollection.MiscSlice[_point];
            if (original.HasPhonetic != value)
            {
                var modified = original;
                modified.HasPhonetic = value;
                _cellsCollection.MiscSlice.Set(_point, in modified);
            }
        }
    }

    #endregion IXLCell Members

    #region IXLStylized Members

    void IXLStylized.ModifyStyle(Func<XLStyleKey, XLStyleKey> modification)
    {
        // XLCell cannot have children, so the base method may be optimized
        var styleKey = modification(StyleValue.Key);
        StyleValue = XLStyleValue.FromKey(ref styleKey);
    }

    /// <summary>
    /// Direct style value setter used by XLStyle typed modify methods to bypass closure allocation.
    /// </summary>
    internal void SetStyleValue(XLStyleValue value)
    {
        StyleValue = value;
    }

    protected override IEnumerable<XLStylizedBase> Children
    {
        get { yield break; }
    }

    public override IXLRanges RangesUsed
    {
        get
        {
            var retVal = new XLRanges { AsRange() };
            return retVal;
        }
    }

    #endregion IXLStylized Members

    /// <summary>
    /// Materialize the cell's currently inherited style into explicit storage,
    /// so it is preserved when the parent row/column/worksheet style changes.
    /// </summary>
    internal void MaterializeStyle()
    {
        StyleValue = GetInheritedStyle();
    }

    private XLStyleValue GetInheritedStyle() => Worksheet.GetStyleValue(SheetPoint);

    public XLRange AsRange()
    {
        return Worksheet.Range(Address, Address);
    }

    #region Styles

    private XLStyleValue GetStyleForRead()
    {
        return StyleValue;
    }

    private void SetStyle(IXLStyle styleToUse)
    {
        Style = styleToUse;
    }

    public bool IsDefaultWorksheetStyle()
    {
        return StyleValue == Worksheet.StyleValue;
    }

    #endregion Styles

    public void DeleteComment()
    {
        Clear(XLClearOptions.Comments);
    }

    public void DeleteSparkline()
    {
        Clear(XLClearOptions.Sparklines);
    }

    private string GetFormat()
    {
        var style = GetStyleForRead();
        if (!string.IsNullOrWhiteSpace(style.NumberFormat.Format)) return style.NumberFormat.Format;
        var formatCodes = XLPredefinedFormat.FormatCodes;
        return formatCodes.TryGetValue(style.NumberFormat.NumberFormatId, out var format) ? format : string.Empty;
    }

    public IXLCell CopyFrom(IXLRangeBase rangeBase)
        => XLCellCopyHelper.CopyFromRange(this, rangeBase);

    private void ClearMerged()
        => XLCellCopyHelper.ClearMerged(this);

    internal string GetFormulaR1C1(string value)
    {
        return XLCellFormula.GetFormula(value, FormulaConversionType.A1ToR1C1,
            _point);
    }

    internal string GetFormulaA1(string value)
    {
        return XLCellFormula.GetFormula(value, FormulaConversionType.R1C1ToA1,
            _point);
    }

    internal void CopyValuesFrom(XLCell source)
        => XLCellCopyHelper.CopyValues(this, source);

    internal IXLCell CopyFromInternal(XLCell otherCell, XLCellCopyOptions options)
        => XLCellCopyHelper.CopyFromInternal(this, otherCell, options);

    public IXLCell CopyFrom(IXLCell otherCell, XLCellCopyOptions options)
    {
        CopyFromInternal((XLCell)otherCell, options);
        return this;
    }

    internal void CopyDataValidation(XLCell otherCell, IXLDataValidation otherDv)
        => XLCellCopyHelper.CopyDataValidation(this, otherCell, otherDv);

    internal void ShiftFormulaRows(XLRange shiftedRange, int rowsShifted)
        => FormulaA1 = XLCellFormulaShifter.ShiftFormulaRows(FormulaA1, Worksheet, shiftedRange, rowsShifted);

    internal void ShiftFormulaColumns(XLRange shiftedRange, int columnsShifted)
        => FormulaA1 = XLCellFormulaShifter.ShiftFormulaColumns(FormulaA1, Worksheet, shiftedRange, columnsShifted);

    private XLCell CellShift(int rowsToShift, int columnsToShift)
    {
        return Worksheet.Cell(_point.Row + rowsToShift, _point.Column + columnsToShift);
    }

    #region XLCell Above

    IXLCell IXLCell.CellAbove()
    {
        return CellAbove();
    }

    IXLCell IXLCell.CellAbove(int step)
    {
        return CellAbove(step);
    }

    public XLCell CellAbove()
    {
        return CellAbove(1);
    }

    public XLCell CellAbove(int step)
    {
        return CellShift(step * -1, 0);
    }

    #endregion XLCell Above

    #region XLCell Below

    IXLCell IXLCell.CellBelow()
    {
        return CellBelow();
    }

    IXLCell IXLCell.CellBelow(int step)
    {
        return CellBelow(step);
    }

    public XLCell CellBelow()
    {
        return CellBelow(1);
    }

    public XLCell CellBelow(int step)
    {
        return CellShift(step, 0);
    }

    #endregion XLCell Below

    #region XLCell Left

    IXLCell IXLCell.CellLeft()
    {
        return CellLeft();
    }

    IXLCell IXLCell.CellLeft(int step)
    {
        return CellLeft(step);
    }

    public XLCell CellLeft()
    {
        return CellLeft(1);
    }

    public XLCell CellLeft(int step)
    {
        return CellShift(0, step * -1);
    }

    #endregion XLCell Left

    #region XLCell Right

    IXLCell IXLCell.CellRight()
    {
        return CellRight();
    }

    IXLCell IXLCell.CellRight(int step)
    {
        return CellRight(step);
    }

    public XLCell CellRight()
    {
        return CellRight(1);
    }

    public XLCell CellRight(int step)
    {
        return CellShift(0, step);
    }

    #endregion XLCell Right

    public bool HasFormula => Formula is not null;

    public bool HasArrayFormula => Formula?.Type == FormulaType.Array;

    public IXLRangeAddress? FormulaReference
    {
        get
        {
            if (Formula is null)
                return null;

            var range = Formula.Range;
            if (range == default)
                return null;

            return XLRangeAddress.FromSheetRange(Worksheet, range);
        }
        set
        {
            if (Formula is null)
            {
                if (IsInferiorMergedCell())
                    return;

                throw new ArgumentException("Cell doesn't contain a formula.");
            }

            if (value is null)
            {
                Formula.Range = default;
                return;
            }

            if (value.Worksheet is not null && Worksheet != value.Worksheet)
                throw new ArgumentException("The reference worksheet must be same as worksheet of the cell or null.");

            Formula.Range = XLSheetRange.FromRangeAddress(value);
        }
    }

    public IXLRange CurrentRegion
        => Worksheet.Range(XLCellRegionHelper.FindCurrentRegion(Worksheet, _point.Row, _point.Column));

    internal bool IsInferiorMergedCell()
    {
        return IsMerged() && !Address.Equals(MergedRange()!.RangeAddress.FirstAddress);
    }

    internal bool IsSuperiorMergedCell()
    {
        return IsMerged() && Address.Equals(MergedRange()!.RangeAddress.FirstAddress);
    }

    internal void GetGlyphBoxes(IXLGraphicEngine engine, Dpi dpi, List<GlyphBox> output)
        => XLCellGlyphHelper.GetGlyphBoxes(this, engine, dpi, output);

    public override int GetHashCode()
    {
        unchecked
        {
            return (SheetPoint.GetHashCode() * 397) ^ Worksheet.GetHashCode();
        }
    }

    public override bool Equals(object? obj)
    {
        return obj is XLCell cell && cell.Worksheet == Worksheet && cell.SheetPoint == SheetPoint;
    }

}
