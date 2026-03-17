using XLibur.Extensions;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Xml;
using XLibur.Excel.Coordinates;
using XLibur.Excel.Rows;
using XLibur.Excel.Tables;
using static XLibur.Excel.IO.OpenXmlConst;
using static XLibur.Excel.XLWorkbook;

namespace XLibur.Excel.IO;

internal static class SheetDataWriter
{
    /// <summary>
    /// Day offset between the 1900 and 1904 date systems used by Excel.
    /// </summary>
    private const int Date1904OffsetDays = 1462;

    private static readonly FieldInfo XmlWriterFieldInfo =
        typeof(OpenXmlPartWriter).GetField("_xmlWriter", BindingFlags.Instance | BindingFlags.NonPublic)!;

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

    internal static void StreamSheetData(OpenXmlWriter writer, XLWorksheet xlWorksheet, SaveContext context,
        SaveOptions options)
    {
        // Steal through reflection for now, the whole OpenXmlPartWriter will be replaced by XmlWriter soon. OpenXmlPartWriter has basically
        // no inner state, unless it is in a string leaf node. By writing SheetData through XmlWriter only, we bypass all that.
        var untypedXmlWriter = XmlWriterFieldInfo.GetValue(writer);
        var xml = (XmlWriter)untypedXmlWriter!;

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

            var cellStyleId = ResolveCellStyleId(xlWorksheet, point, ref lastCachedStyle, ref lastCachedStyleId, context);

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

        return new HashSet<XLSheetPoint>(
            xlWorksheet.Tables
                .Where<XLTable>(table => table.ShowTotalsRow)
                .SelectMany(table => table.TotalsRow()!.CellsUsed())
                .Select(cell => ((XLCell)cell).SheetPoint));
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
        var hasFormula = ctx.CellsCollection.FormulaSlice.IsUsed(point);
        if (hasFormula || (ctx.TableTotalCells is not null && ctx.TableTotalCells.Contains(point)))
        {
            var xlCell = ctx.CellsCollection.GetCell(point);
            WriteCell(xml, xlCell, ctx.CellRef, ctx.SaveContext, ctx.SaveOptions, ctx.TableTotalCells, rowStyleId, cellStyleId);
            return;
        }

        var cellValue = ctx.CellsCollection.ValueSlice.GetCellValue(point);
        if (cellValue.Type != XLDataType.Blank)
        {
            WriteValueOnlyCell(xml, ref ctx, point, cellStyleId, cellValue);
        }
        else if (rowStyleId != cellStyleId)
        {
            WriteBlankStyledCell(xml, ctx.CellsCollection, point, ctx.CellRef, cellStyleId);
        }
    }

    private static void WriteValueOnlyCell(XmlWriter xml, ref CellWriteContext ctx,
        XLSheetPoint point, uint cellStyleId, XLCellValue cellValue)
    {
        Span<char> cellRefSpan = ctx.CellRef;
        var cellRefLen = point.Format(cellRefSpan);
        var shareString = ctx.CellsCollection.ValueSlice.GetShareString(point);
        var dataType = GetCellValueTypeDirect(cellValue.Type, shareString);
        ref readonly var misc = ref ctx.CellsCollection.MiscSlice[point];

        WriteStartCellDirect(xml, ctx.CellRef, cellRefLen, dataType, cellStyleId, in misc);
        WriteCellValueDirect(xml, cellValue, shareString, point, ctx.CellsCollection, ctx.Use1904DateSystem, ctx.SaveContext);
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

    private static void WriteStartCell(XmlWriter w, XLCell xlCell, char[] reference, int referenceLength,
        string? dataType, uint styleId, SaveContext context)
    {
        w.WriteStartElement("c", Main2006SsNs);

        w.WriteStartAttribute("r");
        w.WriteRaw(reference, 0, referenceLength);
        w.WriteEndAttribute();

        w.WriteAttribute("s", styleId);

        if (dataType is not null)
            w.WriteAttributeString("t", dataType);

        if (xlCell.ShowPhonetic)
            w.WriteAttributeString("ph", TrueValue);

        var cmIndex = xlCell.CellMetaIndex;
        if (cmIndex is null && xlCell.Formula is { IsDynamicArray: true } && context.DynamicArrayMetaIndex is not null)
            cmIndex = context.DynamicArrayMetaIndex.Value;

        if (cmIndex is not null)
            w.WriteAttribute("cm", cmIndex.Value);

        if (xlCell.ValueMetaIndex is not null)
            w.WriteAttribute("vm", xlCell.ValueMetaIndex.Value);
    }

    private static void WriteCell(XmlWriter xml, XLCell xlCell, char[] cellRef, SaveContext context,
        SaveOptions options, HashSet<XLSheetPoint>? tableTotalCells, uint rowStyleId, uint styleId)
    {
        Span<char> cellRefSpan = cellRef;
        var cellRefLen = xlCell.SheetPoint.Format(cellRefSpan);

        if (xlCell.HasFormula)
        {
            WriteCellWithFormula(xml, xlCell, cellRef, cellRefLen, context, options, styleId);
        }
        else if (tableTotalCells is not null && tableTotalCells.Contains(xlCell.SheetPoint))
        {
            WriteCellWithTotalLabel(xml, xlCell, cellRef, cellRefLen, context, styleId);
        }
        else if (xlCell.DataType != XLDataType.Blank)
        {
            var dataType = GetCellValueType(xlCell);
            WriteStartCell(xml, xlCell, cellRef, cellRefLen, dataType, styleId, context);
            WriteCellValue(xml, xlCell, context);
            xml.WriteEndElement(); // cell
        }
        else if (rowStyleId != styleId)
        {
            WriteStartCell(xml, xlCell, cellRef, cellRefLen, null, styleId, context);
            xml.WriteEndElement(); // cell
        }
    }

    private static void WriteCellWithFormula(XmlWriter xml, XLCell xlCell, char[] cellRef, int cellRefLen,
        SaveContext context, SaveOptions options, uint styleId)
    {
        string? dataType = null;
        if (options.EvaluateFormulasBeforeSaving)
        {
            try
            {
                xlCell.Evaluate(false);
                dataType = FormulaDataType[(int)xlCell.DataType];
            }
            catch
            {
                // Do nothing, cell will be left blank. Unimplemented features or functions would stop trying to save a file.
            }
        }

        WriteStartCell(xml, xlCell, cellRef, cellRefLen, dataType, styleId, context);

        var xlFormula = xlCell.Formula!;
        if (xlFormula.Type == FormulaType.DataTable)
            WriteDataTableFormula(xml, xlFormula);
        else if (xlCell.HasArrayFormula)
            WriteArrayFormula(xml, xlCell);
        else
        {
            xml.WriteStartElement("f", Main2006SsNs);
            xml.WriteString(xlCell.FormulaA1);
            xml.WriteEndElement(); // f
        }

        if (options.EvaluateFormulasBeforeSaving && xlCell.CachedValue.Type != XLDataType.Blank &&
            !xlCell.NeedsRecalculation)
        {
            WriteCellValue(xml, xlCell, context);
        }

        xml.WriteEndElement(); // cell
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
        // It is necessary to mark some precedent formula dirty (e.g. edit cell formula and enter in Excel).
        // By setting the CalculateCell, we ensure that Excel will calculate values of data table formula on load and
        // user will see correct values.
        xml.WriteAttributeString("ca", TrueValue);

        xml.WriteEndElement(); // f
    }

    private static void WriteArrayFormula(XmlWriter xml, XLCell xlCell)
    {
        var isMasterCell = xlCell.Formula!.Range.FirstPoint == xlCell.SheetPoint;
        if (isMasterCell)
        {
            xml.WriteStartElement("f", Main2006SsNs);
            xml.WriteAttributeString("t", "array");
            xml.WriteAttributeString("ref", xlCell.FormulaReference!.ToStringRelative());
            xml.WriteString(xlCell.FormulaA1);
            xml.WriteEndElement(); // f
        }
    }

    private static void WriteCellWithTotalLabel(XmlWriter xml, XLCell xlCell, char[] cellRef, int cellRefLen,
        SaveContext context, uint styleId)
    {
        var table = xlCell.Worksheet.Tables.First<XLTable>(t => t.AsRange().Contains(xlCell));
        var field = (XLTableField)table.Fields.First(f => f.Column.ColumnNumber() == xlCell.SheetPoint.Column);

        if (!string.IsNullOrWhiteSpace(field.TotalsRowLabel))
        {
            var sharedStringId = context.GetSharedStringId(xlCell, field.TotalsRowLabel);
            WriteStartCell(xml, xlCell, cellRef, cellRefLen, "s", styleId, context);
            xml.WriteStartElement("v", Main2006SsNs);
            xml.WriteValue(sharedStringId);
            xml.WriteEndElement();
        }

        xml.WriteEndElement(); // cell
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
            w.WriteStartElement("v", Main2006SsNs);
            w.WriteValue(sharedStringId);
            w.WriteEndElement();
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

    internal static void WriteCellValue(XmlWriter w, XLCell xlCell, SaveContext context)
    {
        var dataType = xlCell.DataType;
        switch (dataType)
        {
            case XLDataType.Blank:
                return;
            case XLDataType.Text:
                WriteCellValueText(w, xlCell, context);
                break;
            case XLDataType.TimeSpan:
                WriteNumberValue(w, xlCell.Value.GetUnifiedNumber());
                break;
            case XLDataType.Number:
                WriteNumberValue(w, xlCell.Value.GetNumber());
                break;
            case XLDataType.DateTime:
                {
                    // OpenXML SDK validator requires a specific format, in addition to the spec, but can read many more
                    var date = xlCell.GetDateTime();
                    if (xlCell.Worksheet.Workbook.Use1904DateSystem)
                        date = date.AddDays(-Date1904OffsetDays);

                    WriteNumberValue(w, date.ToSerialDateTime());
                    break;
                }
            case XLDataType.Boolean:
                WriteStringValue(w, xlCell.GetBoolean() ? TrueValue : FalseValue);
                break;
            case XLDataType.Error:
                WriteStringValue(w, xlCell.Value.GetError().ToDisplayString());
                break;
            default:
                throw new InvalidOperationException();
        }
    }

    private static void WriteCellValueText(XmlWriter w, XLCell xlCell, SaveContext context)
    {
        var text = xlCell.GetText();
        if (xlCell.HasFormula)
        {
            WriteStringValue(w, text);
            return;
        }

        if (xlCell.ShareString)
        {
            var sharedStringId = context.GetSharedStringId(xlCell, text);
            w.WriteStartElement("v", Main2006SsNs);
            w.WriteValue(sharedStringId);
            w.WriteEndElement();
        }
        else
        {
            w.WriteStartElement("is", Main2006SsNs);
            var richText = xlCell.RichText;
            if (richText is not null)
            {
                TextSerializer.WriteRichTextElements(w, richText, context);
            }
            else
            {
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

    private static string? GetCellValueType(XLCell xlCell)
    {
        var dataType = xlCell.DataType;
        if (dataType == XLDataType.Text && !xlCell.ShareString)
            return "inlineStr";
        return ValueDataType[(int)dataType];
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
