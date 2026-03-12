using XLibur.Extensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Xml;
using static XLibur.Excel.IO.OpenXmlConst;
using static XLibur.Excel.XLWorkbook;

namespace XLibur.Excel.IO;

internal sealed class SheetDataWriter
{
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

        HashSet<XLSheetPoint>? tableTotalCells = null;
        if (xlWorksheet.Tables.Count > 0)
        {
            tableTotalCells = new HashSet<XLSheetPoint>(
                xlWorksheet.Tables
                    .Where<XLTable>(table => table.ShowTotalsRow)
                    .SelectMany(table =>
                        table.TotalsRow()!.CellsUsed())
                    .Select(cell => ((XLCell)cell).SheetPoint));
        }

        // A rather complicated state machine, so rows and cells can be written in a single loop
        var openedRowNumber = 0;
        var isRowOpened = false;
        var cellRef = new char[10]; // Buffer must be enough to hold span and rowNumber as strings
        List<int> rows;
        if (xlWorksheet.Internals.RowsCollection.Count > 0)
        {
            rows = xlWorksheet.Internals.RowsCollection.Keys.ToList();
            rows.Sort();
        }
        else
        {
            rows = [];
        }

        var rowPropIndex = 0;
        uint rowStyleId = 0;
        XLStyleValue? lastCachedStyle = null;
        uint lastCachedStyleId = 0;
        var cellsCollection = xlWorksheet.Internals.CellsCollection;
        var use1904DateSystem = xlWorksheet.Workbook.Use1904DateSystem;
        var enumerator = new XLCellsCollection.SlicesEnumerator(XLSheetRange.Full, cellsCollection);
        while (enumerator.MoveNext())
        {
            var point = enumerator.Current;
            var currentRowNumber = point.Row;

            // A space between cells can have several rows that don't contain cells
            // but have custom properties (e.g., height). Write them out.
            while (rowPropIndex < rows.Count && rows[rowPropIndex] < currentRowNumber)
            {
                if (isRowOpened)
                {
                    xml.WriteEndElement(); // row
                    isRowOpened = false;
                }

                var rowNumber = rows[rowPropIndex];
                var xlRow = xlWorksheet.Internals.RowsCollection[rowNumber];
                if (RowHasCustomProps(xlRow))
                {
                    WriteStartRow(xml, xlRow, rowNumber, maxColumn, context);

                    isRowOpened = true;
                    openedRowNumber = rowNumber;
                }

                rowPropIndex++;
            }

            // Quick check from value slice - avoids XLCell creation for non-blank cells
            var cellValue = cellsCollection.ValueSlice.GetCellValue(point);
            if (cellValue.Type == XLDataType.Blank)
            {
                // For blank cells, need full IsEmpty check which requires XLCell
                var xlCell = cellsCollection.GetCell(point);
                var isEmpty = xlCell.IsEmpty(XLCellsUsedOptions.All
                                             & ~XLCellsUsedOptions.ConditionalFormats
                                             & ~XLCellsUsedOptions.DataValidation
                                             & ~XLCellsUsedOptions.MergedRanges);
                if (isEmpty)
                    continue;
            }

            if (openedRowNumber != currentRowNumber)
            {
                if (isRowOpened)
                    xml.WriteEndElement(); // row

                if (xlWorksheet.Internals.RowsCollection.TryGetValue(currentRowNumber, out var row))
                {
                    rowPropIndex++;
                    rowStyleId = context.SharedStyles[row.StyleValue].StyleId;
                }
                else
                {
                    rowStyleId = 0;
                }

                WriteStartRow(xml, row, currentRowNumber, maxColumn, context);

                isRowOpened = true;
                openedRowNumber = currentRowNumber;
            }

            var cellStyleValue = xlWorksheet.GetStyleValue(point);
            uint cellStyleId;
            if (ReferenceEquals(cellStyleValue, lastCachedStyle))
                cellStyleId = lastCachedStyleId;
            else
            {
                lastCachedStyle = cellStyleValue;
                lastCachedStyleId = cellStyleId = context.SharedStyles[cellStyleValue].StyleId;
            }

            // Formula and table-total paths are rare — only create XLCell for them
            var hasFormula = cellsCollection.FormulaSlice.IsUsed(point);
            if (hasFormula || (tableTotalCells is not null && tableTotalCells.Contains(point)))
            {
                var xlCell = cellsCollection.GetCell(point);
                WriteCell(xml, xlCell, cellRef, context, options, tableTotalCells, rowStyleId, cellStyleId);
            }
            else if (cellValue.Type != XLDataType.Blank)
            {
                // Common value-only path — no XLCell allocation
                Span<char> cellRefSpan = cellRef;
                var cellRefLen = point.Format(cellRefSpan);
                var shareString = cellsCollection.ValueSlice.GetShareString(point);
                var dataType = GetCellValueTypeDirect(cellValue.Type, shareString);
                ref readonly var misc = ref cellsCollection.MiscSlice[point];

                WriteStartCellDirect(xml, cellRef, cellRefLen, dataType, cellStyleId, in misc);
                WriteCellValueDirect(xml, cellValue, shareString, point, cellsCollection, use1904DateSystem, context);
                xml.WriteEndElement(); // cell
            }
            else if (rowStyleId != cellStyleId)
            {
                // Blank cell with different style from row
                Span<char> cellRefSpan = cellRef;
                var cellRefLen = point.Format(cellRefSpan);
                ref readonly var misc = ref cellsCollection.MiscSlice[point];

                WriteStartCellDirect(xml, cellRef, cellRefLen, null, cellStyleId, in misc);
                xml.WriteEndElement(); // cell
            }
        }

        if (isRowOpened)
            xml.WriteEndElement(); // row

        // Write rows with custom properties after last cell.
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

        xml.WriteEndElement(); // SheetData
        return;

        static bool RowHasCustomProps(XLRow xlRow)
        {
            return xlRow.HeightChanged ||
                   xlRow.IsHidden ||
                   xlRow.StyleValue != xlRow.Worksheet.StyleValue ||
                   xlRow.Collapsed ||
                   xlRow.OutlineLevel > 0;
        }

        static void WriteStartRow(XmlWriter w, XLRow? xlRow, int rowNumber, int maxColumn, SaveContext context)
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
            {
                w.WriteAttributeString("hidden", TrueValue);
            }

            if (xlRow.StyleValue != xlRow.Worksheet.StyleValue)
            {
                var styleIndex = context.SharedStyles[xlRow.StyleValue].StyleId;
                w.WriteAttribute("s", styleIndex);
                w.WriteAttributeString("customFormat", TrueValue);
            }

            if (xlRow.Collapsed)
            {
                w.WriteAttributeString("collapsed", TrueValue);
            }

            if (xlRow.OutlineLevel > 0)
            {
                w.WriteAttribute("outlineLevel", xlRow.OutlineLevel);
            }

            if (xlRow.ShowPhonetic)
            {
                w.WriteAttributeString("ph", TrueValue);
            }

            if (xlRow.DyDescent is not null)
            {
                w.WriteAttribute("dyDescent", X14Ac2009SsNs, xlRow.DyDescent.Value);
            }

            // thickBot and thickTop attributes are not written, because Excel seems to determine adjustments
            // from cell borders on its own, and it would be rather costly to check each cell in each row.
            // If the row was adjusted when the cell had its border modified, then it would be fine to write
            // the thickBot/thickBot attributes.
        }

        static void WriteStartCell(XmlWriter w, XLCell xlCell, char[] reference, int referenceLength, string? dataType,
            uint styleId)
        {
            w.WriteStartElement("c", Main2006SsNs);

            w.WriteStartAttribute("r");
            w.WriteRaw(reference, 0, referenceLength);
            w.WriteEndAttribute();

            // TODO: if (styleId != 0) Test files have style even for 0, fix later
            w.WriteAttribute("s", styleId);

            if (dataType is not null)
                w.WriteAttributeString("t", dataType);

            if (xlCell.ShowPhonetic)
                w.WriteAttributeString("ph", TrueValue);

            if (xlCell.CellMetaIndex is not null)
                w.WriteAttribute("cm", xlCell.CellMetaIndex.Value);

            if (xlCell.ValueMetaIndex is not null)
                w.WriteAttribute("vm", xlCell.ValueMetaIndex.Value);
        }

        static void WriteCell(XmlWriter xml, XLCell xlCell, char[] cellRef, SaveContext context, SaveOptions options,
            HashSet<XLSheetPoint>? tableTotalCells, uint rowStyleId, uint styleId)
        {

            Span<char> cellRefSpan = cellRef;
            var cellRefLen = xlCell.SheetPoint.Format(cellRefSpan);

            if (xlCell.HasFormula)
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

                WriteStartCell(xml, xlCell, cellRef, cellRefLen, dataType, styleId);

                var xlFormula = xlCell.Formula!;
                if (xlFormula.Type == FormulaType.DataTable)
                {
                    // Data table doesn't write actual text of formula, that is referenced by context
                    xml.WriteStartElement("f", Main2006SsNs);
                    xml.WriteAttributeString("t", "dataTable");
                    xml.WriteAttributeString("ref", xlFormula.Range.ToString());

                    var is2D = xlFormula.Is2DDataTable;
                    if (is2D)
                        xml.WriteAttributeString("dt2D", TrueValue);

                    var isDataRowTable = xlFormula.IsRowDataTable;
                    if (isDataRowTable)
                        xml.WriteAttributeString("dtr", TrueValue);

                    xml.WriteAttributeString("r1", xlFormula.Input1.ToString());
                    var input1Deleted = xlFormula.Input1Deleted;
                    if (input1Deleted)
                        xml.WriteAttributeString("del1", TrueValue);

                    if (is2D)
                        xml.WriteAttributeString("r2", xlFormula.Input2.ToString());

                    var input2Deleted = xlFormula.Input2Deleted;
                    if (input2Deleted)
                        xml.WriteAttributeString("del2", TrueValue);

                    // Excel doesn't recalculate table formula on a load or on the click of a button or any kind of forced recalculation.
                    // It is necessary to mark some precedent formula dirty (e.g. edit cell formula and enter in Excel).
                    // By setting the CalculateCell, we ensure that Excel will calculate values of data table formula on load and
                    // user will see correct values.
                    xml.WriteAttributeString("ca", TrueValue);

                    xml.WriteEndElement(); // f
                }
                else if (xlCell.HasArrayFormula)
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
            else if (tableTotalCells is not null && tableTotalCells.Contains(xlCell.SheetPoint))
            {
                var table = xlCell.Worksheet.Tables.First<XLTable>(t => t.AsRange().Contains(xlCell));
                var field =
                    (XLTableField)table.Fields.First(f => f.Column.ColumnNumber() == xlCell.SheetPoint.Column);

                // If this is a cell in the total row that contains a label (xor with function), write label
                // Only label can be written. Total functions are basically formulas that use structured
                // references, and SR are not yet supported, so not yet possible to calculate total values.
                if (!string.IsNullOrWhiteSpace(field.TotalsRowLabel))
                {
                    // Excel requires that table totals row label attribute in tableColumn must match the cell
                    // string from SST. If they don't match, Excel will consider it a corrupt workbook.
                    var sharedStringId = context.GetSharedStringId(xlCell, field.TotalsRowLabel);
                    WriteStartCell(xml, xlCell, cellRef, cellRefLen, "s", styleId);
                    xml.WriteStartElement("v", Main2006SsNs);
                    xml.WriteValue(sharedStringId);
                    xml.WriteEndElement();
                }

                xml.WriteEndElement(); // cell
            }
            else if (xlCell.DataType != XLDataType.Blank)
            {
                // Cell contains only a value
                var dataType = GetCellValueType(xlCell);
                WriteStartCell(xml, xlCell, cellRef, cellRefLen, dataType, styleId);

                WriteCellValue(xml, xlCell, context);
                xml.WriteEndElement(); // cell
            }
            else if (rowStyleId != styleId)
            {
                // Cell is blank and should be written only if it has different style from parent.
                // Non-written cells use inherited style of a row.
                WriteStartCell(xml, xlCell, cellRef, cellRefLen, null, styleId);
                xml.WriteEndElement(); // cell
            }
        }

        static string? GetCellValueTypeDirect(XLDataType dataType, bool shareString)
        {
            if (dataType == XLDataType.Text && !shareString)
                return "inlineStr";
            return ValueDataType[(int)dataType];
        }

        static void WriteStartCellDirect(XmlWriter w, char[] reference, int referenceLength, string? dataType,
            uint styleId, in XLMiscSliceContent misc)
        {
            w.WriteStartElement("c", Main2006SsNs);

            w.WriteStartAttribute("r");
            w.WriteRaw(reference, 0, referenceLength);
            w.WriteEndAttribute();

            // TODO: if (styleId != 0) Test files have style even for 0, fix later
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

        static void WriteCellValueDirect(XmlWriter w, XLCellValue cellValue, bool shareString,
            XLSheetPoint point, XLCellsCollection cellsCollection, bool use1904DateSystem, SaveContext context)
        {
            switch (cellValue.Type)
            {
                case XLDataType.Blank:
                    return;
                case XLDataType.Text:
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

                    break;
                }
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
                        date = date.AddDays(-1462);

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

            return;

            static void WriteStringValue(XmlWriter w, string text)
            {
                w.WriteStartElement("v", Main2006SsNs);
                w.WriteString(text);
                w.WriteEndElement();
            }

            static void WriteNumberValue(XmlWriter w, double value)
            {
                w.WriteStartElement("v", Main2006SsNs);
                w.WriteNumberValue(value);
                w.WriteEndElement();
            }
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
            {
                var text = xlCell.GetText();
                if (xlCell.HasFormula)
                {
                    WriteStringValue(w, text);
                }
                else
                {
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

                break;
            }
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
                    date = date.AddDays(-1462);

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

        return;

        static void WriteStringValue(XmlWriter w, string text)
        {
            w.WriteStartElement("v", Main2006SsNs);
            w.WriteString(text);
            w.WriteEndElement();
        }

        static void WriteNumberValue(XmlWriter w, double value)
        {
            w.WriteStartElement("v", Main2006SsNs);
            w.WriteNumberValue(value);
            w.WriteEndElement();
        }
    }

    private static string? GetCellValueType(XLCell xlCell)
    {
        var dataType = xlCell.DataType;
        if (dataType == XLDataType.Text && !xlCell.ShareString)
            return "inlineStr";
        return ValueDataType[(int)dataType];
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
