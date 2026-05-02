using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using XLibur.Excel.Coordinates;
using XLibur.Excel.Rows;
using XLibur.Excel.Tables;
using XLibur.Extensions;
using static XLibur.Excel.IO.OpenXmlConst;
using static XLibur.Excel.XLWorkbook;

namespace XLibur.Excel.IO;

internal static class SheetDataWriter
{
    /// <summary>
    /// Day offset between the 1900 and 1904 date systems used by Excel.
    /// </summary>
    private const int Date1904OffsetDays = 1462;

    /// <summary>
    /// An array to convert data type for a formula cell. Key is <see cref="XLDataType"/>.
    /// It saves some performance through direct indexation instead of switch.
    /// </summary>
    private static readonly string?[] FormulaDataType =
    [
        null, // blank
        "b", // boolean
        null, // number, default value, no need to save type
        "str", // text, formula can only save this type, no inline or shared string
        "e", // error
        null, // datetime, saved as serialized date-time
        null // timespan, saved as serialized date-time
    ];

    /// <summary>
    /// An array to convert a data type for a cell that only contains a value. Key is <see cref="XLDataType"/>.
    /// It saves some performance through direct indexation instead of switch.
    /// </summary>
    private static readonly string?[] ValueDataType =
    [
        null, // blank
        "b", // boolean
        null, // number, default value, no need to save type
        "s", // text, the default is a shared string, but there also can be inline string depending on ShareString property
        "e", // error
        null, // datetime, saved as serialized date-time
        null // timespan, saved as serialized date-time
    ];

    internal static void StreamSheetData(XmlWriter xml, XLWorksheet xlWorksheet, SaveContext context,
        SaveOptions options)
    {
        var maxColumn = GetMaxColumn(xlWorksheet);

        xml.WriteStartElement("sheetData", Main2006SsNs);

        var tableTotalCells = CollectTableTotalCells(xlWorksheet);

        // A rather complicated state machine, so rows and cells can be written in a single loop
        var rowState = new RowWriterState();
        var rows = GetSortedRowNumbers(xlWorksheet);
        var cellCtx = new CellWriteContext
        {
            CellsCollection = xlWorksheet.Internals.CellsCollection,
            CellRef = new char[10], // Buffer must be enough to hold span and rowNumber as strings
            SaveContext = context,
            SaveOptions = options,
            TableTotalCells = tableTotalCells,
            Use1904DateSystem = xlWorksheet.Workbook.Use1904DateSystem,
        };
        uint rowStyleId = 0;
        XLStyleValue? lastCachedStyle = null;
        uint lastCachedStyleId = 0;
        var enumerator = new XLCellsCollection.SlicesEnumerator(XLSheetRange.Full, cellCtx.CellsCollection);
        while (enumerator.MoveNext())
        {
            var point = enumerator.Current;
            var currentRowNumber = point.Row;

            WriteIntermediateRows(xml, xlWorksheet, rows, currentRowNumber, maxColumn, context, ref rowState);

            if (IsBlankAndEmpty(cellCtx.CellsCollection, point))
                continue;

            if (rowState.OpenedRowNumber != currentRowNumber)
            {
                if (rowState.IsRowOpened)
                    xml.WriteEndElement(); // row

                rowStyleId = ResolveRowStyleId(xlWorksheet, currentRowNumber, ref rowState.RowPropIndex, context);

                xlWorksheet.Internals.RowsCollection.TryGetValue(currentRowNumber, out var row);
                WriteStartRow(xml, row, currentRowNumber, maxColumn, context);

                rowState.IsRowOpened = true;
                rowState.OpenedRowNumber = currentRowNumber;
            }

            var cellStyleId =
                ResolveCellStyleId(xlWorksheet, point, ref lastCachedStyle, ref lastCachedStyleId, context);

            WriteCellAtPoint(xml, ref cellCtx, point, rowStyleId, cellStyleId);
        }

        if (rowState.IsRowOpened)
            xml.WriteEndElement(); // row

        WriteTrailingRows(xml, xlWorksheet, rows, rowState.RowPropIndex, context);

        xml.WriteEndElement(); // SheetData
    }

    private static HashSet<XLSheetPoint>? CollectTableTotalCells(XLWorksheet xlWorksheet)
    {
        if (xlWorksheet.Tables.Count == 0)
            return null;

        HashSet<XLSheetPoint>? cells = null;
        foreach (var table in xlWorksheet.Tables)
        {
            if (!table.ShowTotalsRow)
                continue;

            cells ??= [];
            foreach (var cell in table.TotalsRow()!.CellsUsed())
                cells.Add(((XLCell)cell).SheetPoint);
        }

        return cells;
    }

    private static List<int> GetSortedRowNumbers(XLWorksheet xlWorksheet)
    {
        if (xlWorksheet.Internals.RowsCollection.Count <= 0)
            return [];

        var rows = xlWorksheet.Internals.RowsCollection.Keys.ToList();
        rows.Sort();
        return rows;
    }

    private static void WriteIntermediateRows(
        XmlWriter xml, XLWorksheet xlWorksheet, List<int> rows,
        int currentRowNumber, int maxColumn, SaveContext context,
        ref RowWriterState state)
    {
        while (state.RowPropIndex < rows.Count && rows[state.RowPropIndex] < currentRowNumber)
        {
            if (state.IsRowOpened)
            {
                xml.WriteEndElement(); // row
                state.IsRowOpened = false;
            }

            var rowNumber = rows[state.RowPropIndex];
            var xlRow = xlWorksheet.Internals.RowsCollection[rowNumber];
            if (RowHasCustomProps(xlRow))
            {
                WriteStartRow(xml, xlRow, rowNumber, maxColumn, context);
                state.IsRowOpened = true;
                state.OpenedRowNumber = rowNumber;
            }

            state.RowPropIndex++;
        }
    }

    private static bool RowHasCustomProps(XLRow xlRow)
    {
        return xlRow.HeightChanged ||
               xlRow.IsHidden ||
               xlRow.StyleValue != xlRow.Worksheet.StyleValue ||
               xlRow.Collapsed ||
               xlRow.OutlineLevel > 0;
    }

    private static bool IsBlankAndEmpty(XLCellsCollection cellsCollection, XLSheetPoint point)
    {
        var cellValue = cellsCollection.ValueSlice.GetCellValue(point);
        if (cellValue.Type != XLDataType.Blank)
            return false;

        var xlCell = cellsCollection.GetCell(point);
        return xlCell.IsEmpty(XLCellsUsedOptions.All
                              & ~XLCellsUsedOptions.ConditionalFormats
                              & ~XLCellsUsedOptions.DataValidation
                              & ~XLCellsUsedOptions.MergedRanges);
    }

    private static uint ResolveRowStyleId(XLWorksheet xlWorksheet, int currentRowNumber,
        ref int rowPropIndex, SaveContext context)
    {
        if (xlWorksheet.Internals.RowsCollection.TryGetValue(currentRowNumber, out var row))
        {
            rowPropIndex++;
            return context.SharedStyles[row.StyleValue].StyleId;
        }

        return 0;
    }

    private static uint ResolveCellStyleId(XLWorksheet xlWorksheet, XLSheetPoint point,
        ref XLStyleValue? lastCachedStyle, ref uint lastCachedStyleId, SaveContext context)
    {
        var cellStyleValue = xlWorksheet.GetStyleValue(point);
        if (ReferenceEquals(cellStyleValue, lastCachedStyle))
            return lastCachedStyleId;

        lastCachedStyle = cellStyleValue;
        lastCachedStyleId = context.SharedStyles[cellStyleValue].StyleId;
        return lastCachedStyleId;
    }

    private static void WriteCellAtPoint(XmlWriter xml, ref CellWriteContext ctx,
        XLSheetPoint point, uint rowStyleId, uint cellStyleId)
    {
        var formula = ctx.CellsCollection.FormulaSlice.Get(point);
        if (formula is not null)
        {
            WriteFormulaCellDirect(xml, ref ctx, point, formula, cellStyleId);
            return;
        }

        if (ctx.TableTotalCells is not null && ctx.TableTotalCells.Contains(point))
        {
            WriteTotalLabelCellDirect(xml, ref ctx, point, cellStyleId);
            return;
        }

        // Single slice lookup for both value and share-string flag, avoiding a
        // second Lut traversal that GetShareString would require separately.
        var cellValue = ctx.CellsCollection.ValueSlice.GetCellValueAndShareString(point, out var shareString);
        if (cellValue.Type != XLDataType.Blank)
        {
            WriteValueOnlyCell(xml, ref ctx, point, cellStyleId, cellValue, shareString);
        }
        else if (rowStyleId != cellStyleId)
        {
            WriteBlankStyledCell(xml, ctx.CellsCollection, point, ctx.CellRef, cellStyleId);
        }
    }

    /// <summary>
    /// Write a cell that has a formula directly from slice data, without allocating an
    /// <see cref="XLCell"/> wrapper. Mirrors the legacy <c>WriteCellWithFormula</c> +
    /// <c>WriteStartCell</c> path.
    /// </summary>
    private static void WriteFormulaCellDirect(XmlWriter xml, ref CellWriteContext ctx,
        XLSheetPoint point, XLCellFormula formula, uint cellStyleId)
    {
        var cellsCollection = ctx.CellsCollection;
        var xlWorksheet = cellsCollection.Worksheet;
        var saveContext = ctx.SaveContext;

        if (ctx.SaveOptions.EvaluateFormulasBeforeSaving && formula.IsDirty(xlWorksheet.Workbook))
        {
            try
            {
                var workbook = xlWorksheet.Workbook;
                if (!workbook.CalcEngine.TryEvaluateSingleCell(formula, point, xlWorksheet))
                    workbook.CalcEngine.Recalculate(workbook, null);
            }
            catch
            {
                // Match XLCell.Evaluate(false) tolerance: unimplemented features should not
                // abort the save. The cell is left with whatever cached value (if any) it
                // already has.
            }
        }

        // Determine cell type from cached value (preserves type round-trip for formulas
        // whose evaluation is unsupported).
        var cachedValue = cellsCollection.ValueSlice.GetCellValue(point);
        var cachedValueType = cachedValue.Type;
        var dataType = cachedValueType != XLDataType.Blank ? FormulaDataType[(int)cachedValueType] : null;

        Span<char> cellRefSpan = ctx.CellRef;
        var cellRefLen = point.Format(cellRefSpan);
        ref readonly var misc = ref cellsCollection.MiscSlice[point];

        // Compute "cm" attribute: explicit MiscSlice override, or workbook-wide dynamic-array
        // metadata index for dynamic-array formulas without an explicit override.
        var cmIndex = misc.CellMetaIndex;
        if (cmIndex is null && formula.IsDynamicArray && saveContext.DynamicArrayMetaIndex is not null)
            cmIndex = saveContext.DynamicArrayMetaIndex.Value;

        WriteStartFormulaCellDirect(xml, ctx.CellRef, cellRefLen, dataType, cellStyleId, in misc, cmIndex);

        if (formula.Type == FormulaType.DataTable)
        {
            WriteDataTableFormula(xml, formula);
        }
        else if (formula.Type == FormulaType.Array)
        {
            var isMasterCell = formula.Range.FirstPoint == point;
            if (isMasterCell)
            {
                xml.WriteStartElement("f", Main2006SsNs);
                xml.WriteAttributeString("t", "array");
                var rangeAddress = XLRangeAddress.FromSheetRange(xlWorksheet, formula.Range);
                xml.WriteAttributeString("ref", rangeAddress.ToStringRelative());
                xml.WriteString(formula.A1);
                xml.WriteEndElement(); // f
            }
        }
        else
        {
            xml.WriteStartElement("f", Main2006SsNs);
            xml.WriteString(formula.A1);
            xml.WriteEndElement(); // f
        }

        // Write cached value if present and the formula isn't dirty. Spilled (non-master)
        // array-formula cells also fall through here so their cached values round-trip.
        if (cachedValueType != XLDataType.Blank && formula.IsClean(xlWorksheet.Workbook))
        {
            WriteCachedFormulaValue(xml, cachedValue, ctx.Use1904DateSystem);
        }

        xml.WriteEndElement(); // cell
    }

    /// <summary>
    /// Variant of <see cref="WriteStartCellDirect"/> that takes a pre-computed <c>cm</c>
    /// attribute value. Needed for formula cells where the dynamic-array metadata index is
    /// applied as a fallback when <see cref="XLMiscSliceContent.CellMetaIndex"/> is null.
    /// </summary>
    private static void WriteStartFormulaCellDirect(XmlWriter w, char[] reference, int referenceLength,
        string? dataType, uint styleId, in XLMiscSliceContent misc, uint? cmIndex)
    {
        w.WriteStartElement("c", Main2006SsNs);

        w.WriteStartAttribute("r");
        w.WriteRaw(reference, 0, referenceLength);
        w.WriteEndAttribute();

        w.WriteAttribute("s", styleId);

        if (dataType is not null)
            w.WriteAttributeString("t", dataType);

        if (misc.HasPhonetic)
            w.WriteAttributeString("ph", TrueValue);

        if (cmIndex is not null)
            w.WriteAttribute("cm", cmIndex.Value);

        if (misc.ValueMetaIndex is not null)
            w.WriteAttribute("vm", misc.ValueMetaIndex.Value);
    }

    /// <summary>
    /// Write the cached value of a formula cell. Text is emitted inline; formulas can only
    /// store an inline-string text result, never a shared-string reference.
    /// </summary>
    private static void WriteCachedFormulaValue(XmlWriter w, XLCellValue cellValue, bool use1904DateSystem)
    {
        switch (cellValue.Type)
        {
            case XLDataType.Blank:
                return;
            case XLDataType.Text:
                WriteStringValue(w, cellValue.GetText());
                break;
            case XLDataType.TimeSpan:
                WriteNumberValue(w, cellValue.GetUnifiedNumber());
                break;
            case XLDataType.Number:
                WriteNumberValue(w, cellValue.GetNumber());
                break;
            case XLDataType.DateTime:
                {
                    var date = cellValue.GetDateTime();
                    if (use1904DateSystem)
                        date = date.AddDays(-Date1904OffsetDays);

                    WriteNumberValue(w, date.ToSerialDateTime());
                    break;
                }
            case XLDataType.Boolean:
                WriteStringValue(w, cellValue.GetBoolean() ? TrueValue : FalseValue);
                break;
            case XLDataType.Error:
                WriteStringValue(w, cellValue.GetError().ToDisplayString());
                break;
            default:
                throw new InvalidOperationException();
        }
    }

    /// <summary>
    /// Write a totals-row label cell directly from slice data. The cell is in
    /// <see cref="CellWriteContext.TableTotalCells"/> but has no formula — it carries either
    /// a label (e.g. "Total") or nothing.
    /// </summary>
    private static void WriteTotalLabelCellDirect(XmlWriter xml, ref CellWriteContext ctx,
        XLSheetPoint point, uint cellStyleId)
    {
        var cellsCollection = ctx.CellsCollection;
        var xlWorksheet = cellsCollection.Worksheet;

        XLTable? containingTable = null;
        foreach (var table in xlWorksheet.Tables)
        {
            if (table.Area.Contains(point))
            {
                containingTable = table;
                break;
            }
        }

        XLTableField? field = null;
        if (containingTable is not null)
        {
            foreach (var f in containingTable.Fields)
            {
                if (f.Column.ColumnNumber() == point.Column)
                {
                    field = (XLTableField)f;
                    break;
                }
            }
        }

        if (field is not null && !string.IsNullOrWhiteSpace(field.TotalsRowLabel))
        {
            var memorySstId = cellsCollection.ValueSlice.GetShareStringId(point);
            var sharedStringId = ctx.SaveContext.GetSharedStringId(memorySstId, point);

            Span<char> cellRefSpan = ctx.CellRef;
            var cellRefLen = point.Format(cellRefSpan);
            ref readonly var misc = ref cellsCollection.MiscSlice[point];

            WriteStartCellDirect(xml, ctx.CellRef, cellRefLen, "s", cellStyleId, in misc);
            WriteValue(xml, sharedStringId);
            xml.WriteEndElement(); // cell
        }
    }

    private static void WriteValueOnlyCell(XmlWriter xml, ref CellWriteContext ctx,
        XLSheetPoint point, uint cellStyleId, XLCellValue cellValue, bool shareString)
    {
        Span<char> cellRefSpan = ctx.CellRef;
        var cellRefLen = point.Format(cellRefSpan);
        var dataType = GetCellValueTypeDirect(cellValue.Type, shareString);
        ref readonly var misc = ref ctx.CellsCollection.MiscSlice[point];

        WriteStartCellDirect(xml, ctx.CellRef, cellRefLen, dataType, cellStyleId, in misc);
        WriteCellValueDirect(xml, cellValue, shareString, point, ctx.CellsCollection, ctx.Use1904DateSystem,
            ctx.SaveContext);
        xml.WriteEndElement(); // cell
    }

    private static void WriteBlankStyledCell(XmlWriter xml, XLCellsCollection cellsCollection,
        XLSheetPoint point, char[] cellRef, uint cellStyleId)
    {
        Span<char> cellRefSpan = cellRef;
        var cellRefLen = point.Format(cellRefSpan);
        ref readonly var misc = ref cellsCollection.MiscSlice[point];

        WriteStartCellDirect(xml, cellRef, cellRefLen, null, cellStyleId, in misc);
        xml.WriteEndElement(); // cell
    }

    private static void WriteTrailingRows(XmlWriter xml, XLWorksheet xlWorksheet,
        List<int> rows, int rowPropIndex, SaveContext context)
    {
        while (rowPropIndex < rows.Count)
        {
            var rowNumber = rows[rowPropIndex];
            var xlRow = xlWorksheet.Internals.RowsCollection[rowNumber];
            if (RowHasCustomProps(xlRow))
            {
                WriteStartRow(xml, xlRow, rowNumber, 0, context);
                xml.WriteEndElement(); // row
            }

            rowPropIndex++;
        }
    }

    private static void WriteStartRow(XmlWriter w, XLRow? xlRow, int rowNumber, int maxColumn, SaveContext context)
    {
        w.WriteStartElement("row", Main2006SsNs);

        w.WriteStartAttribute("r");
        w.WriteValue(rowNumber);
        w.WriteEndAttribute();

        if (maxColumn > 0)
        {
            w.WriteStartAttribute("spans");
            w.WriteString("1:");
            w.WriteValue(maxColumn);
            w.WriteEndAttribute();
        }

        if (xlRow is null)
            return;

        WriteRowAttributes(w, xlRow, context);
    }

    private static void WriteRowAttributes(XmlWriter w, XLRow xlRow, SaveContext context)
    {
        if (xlRow.HeightChanged)
        {
            var height = xlRow.Height.SaveRound();
            w.WriteStartAttribute("ht");
            w.WriteNumberValue(height);
            w.WriteEndAttribute();

            // Note that dyDescent automatically implies custom height
            w.WriteAttributeString("customHeight", TrueValue);
        }

        if (xlRow.IsHidden)
            w.WriteAttributeString("hidden", TrueValue);

        if (xlRow.StyleValue != xlRow.Worksheet.StyleValue)
        {
            var styleIndex = context.SharedStyles[xlRow.StyleValue].StyleId;
            w.WriteAttribute("s", styleIndex);
            w.WriteAttributeString("customFormat", TrueValue);
        }

        if (xlRow.Collapsed)
            w.WriteAttributeString("collapsed", TrueValue);

        if (xlRow.OutlineLevel > 0)
            w.WriteAttribute("outlineLevel", xlRow.OutlineLevel);

        if (xlRow.ShowPhonetic)
            w.WriteAttributeString("ph", TrueValue);

        if (xlRow.DyDescent is not null)
            w.WriteAttribute("dyDescent", X14Ac2009SsNs, xlRow.DyDescent.Value);

        // thickBot and thickTop attributes are not written, because Excel seems to determine adjustments
        // from cell borders on its own, and it would be rather costly to check each cell in each row.
        // If the row was adjusted when the cell had its border modified, then it would be fine to write
        // the thickBot/thickBot attributes.
    }

    private static void WriteDataTableFormula(XmlWriter xml, XLCellFormula xlFormula)
    {
        xml.WriteStartElement("f", Main2006SsNs);
        xml.WriteAttributeString("t", "dataTable");
        xml.WriteAttributeString("ref", xlFormula.Range.ToString());

        var is2D = xlFormula.Is2DDataTable;
        if (is2D)
            xml.WriteAttributeString("dt2D", TrueValue);

        if (xlFormula.IsRowDataTable)
            xml.WriteAttributeString("dtr", TrueValue);

        xml.WriteAttributeString("r1", xlFormula.Input1.ToString());
        if (xlFormula.Input1Deleted)
            xml.WriteAttributeString("del1", TrueValue);

        if (is2D)
            xml.WriteAttributeString("r2", xlFormula.Input2.ToString());

        if (xlFormula.Input2Deleted)
            xml.WriteAttributeString("del2", TrueValue);

        // Excel doesn't recalculate table formula on a load or on the click of a button or any kind of forced recalculation.
        // It is necessary to mark some precedent formula dirty (e.g., edit cell formula and enter in Excel).
        // By setting the CalculateCell, we ensure that Excel will calculate values of data table formula on load and
        // the user will see correct values.
        xml.WriteAttributeString("ca", TrueValue);

        xml.WriteEndElement(); // f
    }

    private static void WriteValue(XmlWriter xml, int sharedStringId)
    {
        xml.WriteStartElement("v", Main2006SsNs);
        xml.WriteValue(sharedStringId);
        xml.WriteEndElement();
    }

    private static string? GetCellValueTypeDirect(XLDataType dataType, bool shareString)
    {
        if (dataType == XLDataType.Text && !shareString)
            return "inlineStr";
        return ValueDataType[(int)dataType];
    }

    private static void WriteStartCellDirect(XmlWriter w, char[] reference, int referenceLength, string? dataType,
        uint styleId, in XLMiscSliceContent misc)
    {
        w.WriteStartElement("c", Main2006SsNs);

        w.WriteStartAttribute("r");
        w.WriteRaw(reference, 0, referenceLength);
        w.WriteEndAttribute();

        w.WriteAttribute("s", styleId);

        if (dataType is not null)
            w.WriteAttributeString("t", dataType);

        if (misc.HasPhonetic)
            w.WriteAttributeString("ph", TrueValue);

        if (misc.CellMetaIndex is not null)
            w.WriteAttribute("cm", misc.CellMetaIndex.Value);

        if (misc.ValueMetaIndex is not null)
            w.WriteAttribute("vm", misc.ValueMetaIndex.Value);
    }

    private static void WriteCellValueDirect(XmlWriter w, XLCellValue cellValue, bool shareString,
        XLSheetPoint point, XLCellsCollection cellsCollection, bool use1904DateSystem, SaveContext context)
    {
        switch (cellValue.Type)
        {
            case XLDataType.Blank:
                return;
            case XLDataType.Text:
                WriteCellValueDirectText(w, cellValue, shareString, point, cellsCollection, context);
                break;
            case XLDataType.TimeSpan:
                WriteNumberValue(w, cellValue.GetUnifiedNumber());
                break;
            case XLDataType.Number:
                WriteNumberValue(w, cellValue.GetNumber());
                break;
            case XLDataType.DateTime:
                {
                    var date = cellValue.GetDateTime();
                    if (use1904DateSystem)
                        date = date.AddDays(-Date1904OffsetDays);

                    WriteNumberValue(w, date.ToSerialDateTime());
                    break;
                }
            case XLDataType.Boolean:
                WriteStringValue(w, cellValue.GetBoolean() ? TrueValue : FalseValue);
                break;
            case XLDataType.Error:
                WriteStringValue(w, cellValue.GetError().ToDisplayString());
                break;
            default:
                throw new InvalidOperationException();
        }
    }

    private static void WriteCellValueDirectText(XmlWriter w, XLCellValue cellValue, bool shareString,
        XLSheetPoint point, XLCellsCollection cellsCollection, SaveContext context)
    {
        if (shareString)
        {
            var memorySstId = cellsCollection.ValueSlice.GetShareStringId(point);
            var sharedStringId = context.GetSharedStringId(memorySstId, point);
            WriteValue(w, sharedStringId);
        }
        else
        {
            w.WriteStartElement("is", Main2006SsNs);
            var richText = cellsCollection.ValueSlice.GetRichText(point);
            if (richText is not null)
            {
                TextSerializer.WriteRichTextElements(w, richText, context);
            }
            else
            {
                var text = cellValue.GetText();
                w.WriteStartElement("t", Main2006SsNs);
                if (text.PreserveSpaces())
                    w.WritePreserveSpaceAttr();

                w.WriteString(text);
                w.WriteEndElement();
            }

            w.WriteEndElement(); // is
        }
    }

    private static void WriteStringValue(XmlWriter w, string text)
    {
        w.WriteStartElement("v", Main2006SsNs);
        w.WriteString(text);
        w.WriteEndElement();
    }

    private static void WriteNumberValue(XmlWriter w, double value)
    {
        w.WriteStartElement("v", Main2006SsNs);
        w.WriteNumberValue(value);
        w.WriteEndElement();
    }

    private struct RowWriterState
    {
        public bool IsRowOpened;
        public int OpenedRowNumber;
        public int RowPropIndex;
    }

    private ref struct CellWriteContext
    {
        public XLCellsCollection CellsCollection;
        public char[] CellRef;
        public SaveContext SaveContext;
        public SaveOptions SaveOptions;
        public HashSet<XLSheetPoint>? TableTotalCells;
        public bool Use1904DateSystem;
    }

    internal static int GetMaxColumn(XLWorksheet xlWorksheet)
    {
        var maxColumn = 0;

        if (!xlWorksheet.Internals.CellsCollection.IsEmpty)
        {
            maxColumn = xlWorksheet.Internals.CellsCollection.MaxColumnUsed;
        }

        if (xlWorksheet.Internals.ColumnsCollection.Count <= 0) return maxColumn;
        var maxColCollection = xlWorksheet.Internals.ColumnsCollection.Keys.Max();
        if (maxColCollection > maxColumn)
            maxColumn = maxColCollection;

        return maxColumn;
    }
}
