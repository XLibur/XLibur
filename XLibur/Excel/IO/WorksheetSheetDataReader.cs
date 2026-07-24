using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.CalcEngine.Visitors;
using XLibur.Excel.Coordinates;
using XLibur.Extensions;
using XLibur.Utils;
using static XLibur.Excel.XLPredefinedFormat.DateTime;

namespace XLibur.Excel.IO;

/// <summary>
/// Reads cell, row, and column data from a worksheet part, including style application and formula handling.
/// </summary>
internal static class WorksheetSheetDataReader
{
    /// <summary>
    /// Loop-invariant parameters for sheet data reading.
    /// </summary>
    internal readonly struct SheetDataReadContext(
        StylesheetData styles,
        XLWorksheet worksheet,
        SharedStringEntry[]? sharedStrings,
        Dictionary<uint, string> sharedFormulasR1C1,
        Dictionary<int, XLStyleValue> styleList,
        Dictionary<XLNumberFormatValue, XLDataType> numberDataTypeCache,
        bool use1904DateSystem,
        HashSet<uint>? dynamicArrayCmIndexes = null)
    {
        public readonly StylesheetData Styles = styles;
        public readonly XLWorksheet Worksheet = worksheet;
        public readonly SharedStringEntry[]? SharedStrings = sharedStrings;
        public readonly Dictionary<uint, string> SharedFormulasR1C1 = sharedFormulasR1C1;
        public readonly Dictionary<int, XLStyleValue> StyleList = styleList;

        /// <summary>
        /// Memoizes the <see cref="XLDataType"/> derived from a cell's number format
        /// (see <see cref="GetNumberDataType"/>). Keyed by the interned
        /// <see cref="XLNumberFormatValue"/>, so it holds one entry per distinct format
        /// rather than recomputing the format-string scan for every numeric cell.
        /// </summary>
        public readonly Dictionary<XLNumberFormatValue, XLDataType> NumberDataTypeCache = numberDataTypeCache;

        public readonly bool Use1904DateSystem = use1904DateSystem;

        /// <summary>
        /// 1-based cell-metadata (<c>cm</c>) indexes that mark a formula as a dynamic array.
        /// <c>null</c> when the workbook has no dynamic-array metadata.
        /// </summary>
        public readonly HashSet<uint>? DynamicArrayCmIndexes = dynamicArrayCmIndexes;

        /// <summary>
        /// Whether the worksheet has any custom column styles. When <c>false</c>,
        /// the inherited style for any cell equals the row-level style, avoiding
        /// per-cell column dictionary lookups during loading.
        /// Evaluated live (not snapshot) because <c>&lt;cols&gt;</c> is parsed
        /// between context construction and the first <c>&lt;row&gt;</c>.
        /// </summary>
        public bool HasColumnStyles => Worksheet.Internals.ColumnsCollection.Count > 0;
    }

    /// <summary>
    /// Mutable tracking state across rows during sheet data reading.
    /// </summary>
    internal struct SheetDataReadState
    {
        public int LastRow;
        public int LastColumnNumber;

        /// <summary>
        /// Cached inherited style for the current row (combines sheet + row styles).
        /// Recomputed once per row in <see cref="LoadRowXml"/> to avoid per-cell
        /// dictionary lookups into <c>RowsCollection</c>.
        /// </summary>
        public XLStyleValue? CachedRowInheritedStyle;
    }

    /// <summary>
    /// Parsed row-level attributes from a &lt;row&gt; element, used to transport
    /// custom properties to <see cref="ApplyRowCustomProps"/> without a long parameter list.
    /// Stack-allocated (readonly record struct) to avoid per-row GC pressure.
    /// </summary>
    private readonly record struct RowProperties(
        double? Height,
        double? DyDescent,
        bool Hidden,
        bool Collapsed,
        int? OutlineLevel,
        bool ShowPhonetic,
        bool CustomFormat,
        int? StyleIndex)
    {
        public bool HasCustomProps =>
            Height is not null || DyDescent is not null || Hidden || Collapsed
            || OutlineLevel > 0 || ShowPhonetic || CustomFormat;
    }

    private static readonly string[] DateCellFormats =
    [
        "yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff", // Format accepted by OpenXML SDK
        "yyyy-MM-ddTHH:mm", "yyyy-MM-dd" // Formats accepted by Excel.
    ];

    // ---------------------------------------------------------------------------------------
    // Raw System.Xml.XmlReader sheet-data reading path.
    //
    // The DocumentFormat.OpenXml OpenXmlPartReader rebuilds a ReadOnlyCollection<OpenXmlAttribute>
    // for every <c>/<row>/<f> element and materializes text nodes through its object model, which
    // dominates load time and allocations for large sheets (~4x slower / ~5x more garbage than a
    // raw XmlReader doing the equivalent traversal). These methods read the <sheetData> hot path
    // directly from a System.Xml.XmlReader while reusing the reader-agnostic value/style/formula
    // helpers below. Structural elements (cols, merges, views, ...) still load via the SDK DOM
    // path in XLWorkbook_Load, which reads them before this runs, so column styles are available.
    //
    // Positioning contract: each LoadXxxXml enters positioned on the element's start node and
    // returns positioned on the node immediately after that element's end tag.
    // ---------------------------------------------------------------------------------------

    /// <summary>
    /// Reads all <c>&lt;row&gt;</c> children of <c>&lt;sheetData&gt;</c> from a raw
    /// <see cref="XmlReader"/>. The reader must be positioned on the first <c>&lt;row&gt;</c>
    /// start element (or on <c>&lt;/sheetData&gt;</c> when there are no rows).
    /// </summary>
    internal static void LoadSheetDataRows(XmlReader reader, in SheetDataReadContext context,
        ref SheetDataReadState state)
    {
        while (IsMainElement(reader, "row"))
            LoadRowXml(reader, in context, ref state);
    }

    /// <summary>
    /// Whether the reader is on a start element with the given local name in the SpreadsheetML
    /// main namespace. Matches the namespace strictness of the SDK reader path so foreign-namespace
    /// elements (e.g. markup-compatibility content) are never mistaken for cells/rows/values.
    /// </summary>
    private static bool IsMainElement(XmlReader reader, string localName)
        => reader.NodeType == XmlNodeType.Element
           && reader.LocalName == localName
           && reader.NamespaceURI == OpenXmlConst.Main2006SsNs;

    private static void LoadRowXml(XmlReader reader, in SheetDataReadContext context,
        ref SheetDataReadState state)
    {
        Debug.Assert(reader is { NodeType: XmlNodeType.Element, LocalName: "row" });

        var rowIndex = 0;
        double? height = null;
        double? dyDescent = null;
        bool hidden = false, collapsed = false, showPhonetic = false, customFormat = false;
        int? outlineLevel = null;
        int? styleIndex = null;
        var isEmptyRow = reader.IsEmptyElement;

        if (reader.HasAttributes)
        {
            while (reader.MoveToNextAttribute())
            {
                var ns = reader.NamespaceURI;
                switch (reader.LocalName)
                {
                    case "r" when ns.Length == 0:
                        TryParseOoxmlNonNegativeInt(reader.Value, out rowIndex);
                        break;
                    case "ht" when ns.Length == 0:
                        height = double.Parse(reader.Value, NumberStyles.Float, XLHelper.ParseCulture);
                        break;
                    case "dyDescent" when ns == OpenXmlConst.X14Ac2009SsNs:
                        dyDescent = double.Parse(reader.Value, NumberStyles.Float, XLHelper.ParseCulture);
                        break;
                    case "hidden" when ns.Length == 0:
                        hidden = ParseXmlBool(reader.Value);
                        break;
                    case "collapsed" when ns.Length == 0:
                        collapsed = ParseXmlBool(reader.Value);
                        break;
                    case "outlineLevel" when ns.Length == 0:
                        outlineLevel = int.Parse(reader.Value);
                        break;
                    case "ph" when ns.Length == 0:
                        showPhonetic = ParseXmlBool(reader.Value);
                        break;
                    case "customFormat" when ns.Length == 0:
                        customFormat = ParseXmlBool(reader.Value);
                        break;
                    case "s" when ns.Length == 0:
                        styleIndex = int.Parse(reader.Value);
                        break;
                }
            }

            reader.MoveToElement();
        }

        if (rowIndex == 0)
            rowIndex = ++state.LastRow;
        state.LastRow = rowIndex;

        var rowProps = new RowProperties(height, dyDescent, hidden, collapsed, outlineLevel, showPhonetic, customFormat, styleIndex);
        if (rowProps.HasCustomProps)
            ApplyRowCustomProps(in rowProps, context.Worksheet, rowIndex, context.Styles);

        var ws = context.Worksheet;
        var sheetStyle = ws.StyleValue;
        var rowStyle = ws.Internals.RowsCollection.TryGetValue(rowIndex, out var r)
            ? r.StyleValue
            : sheetStyle;
        state.CachedRowInheritedStyle = rowStyle;
        state.LastColumnNumber = 0;

        if (isEmptyRow)
        {
            reader.Read();
            return;
        }

        reader.Read(); // Move into the row's children (first <c> or </row>).

        while (IsMainElement(reader, "c"))
            LoadCellXml(reader, in context, rowIndex, ref state);

        // A row can also contain extLst; skip any remaining children.
        while (reader.NodeType == XmlNodeType.Element)
            reader.Skip();

        reader.Read(); // Move past </row>.
    }

    private static void LoadCellXml(XmlReader reader, in SheetDataReadContext context, int rowIndex,
        ref SheetDataReadState state)
    {
        Debug.Assert(reader is { NodeType: XmlNodeType.Element, LocalName: "c" });

        int styleIndex = 0;
        XLSheetPoint? cellRef = null;
        string? typeAttr = null;
        bool showPhonetic = false;
        uint? cellMetaIndex = null;
        uint? valueMetaIndex = null;
        bool hasMisc = false;
        var isEmptyCell = reader.IsEmptyElement;

        if (reader.HasAttributes)
        {
            while (reader.MoveToNextAttribute())
            {
                if (reader.NamespaceURI.Length != 0)
                    continue;

                switch (reader.LocalName)
                {
                    case "r":
                        cellRef = XLSheetPoint.Parse(reader.Value);
                        break;
                    case "s":
                        TryParseOoxmlNonNegativeInt(reader.Value, out styleIndex);
                        break;
                    case "t":
                        typeAttr = reader.Value;
                        break;
                    case "ph":
                        showPhonetic = ParseXmlBool(reader.Value);
                        if (showPhonetic) hasMisc = true;
                        break;
                    case "cm":
                        cellMetaIndex = uint.Parse(reader.Value);
                        hasMisc = true;
                        break;
                    case "vm":
                        valueMetaIndex = uint.Parse(reader.Value);
                        hasMisc = true;
                        break;
                }
            }

            reader.MoveToElement();
        }

        var cellAddress = cellRef ?? new XLSheetPoint(rowIndex, state.LastColumnNumber + 1);
        state.LastColumnNumber = cellAddress.Column;
        var dataType = ParseCellDataType(typeAttr);

        var cellStyleValue = ResolveCachedStyleValue(styleIndex, context.Styles, context.StyleList);

        var ws = context.Worksheet;
        var cellsCollection = ws.Internals.CellsCollection;
        var inherited = GetInheritedStyleFast(ws, state.CachedRowInheritedStyle!, cellAddress.Column, context.HasColumnStyles);
        var styleMatchesInherited = ReferenceEquals(cellStyleValue, inherited);
        if (!styleMatchesInherited)
            cellsCollection.StyleSlice.SetNonDefault(cellAddress.Row, cellAddress.Column, cellStyleValue);

        if (hasMisc)
        {
            var misc = new XLMiscSliceContent
            {
                HasPhonetic = showPhonetic,
                CellMetaIndex = cellMetaIndex,
                ValueMetaIndex = valueMetaIndex
            };
            cellsCollection.MiscSlice.Set(cellAddress, in misc);
        }

        if (isEmptyCell)
        {
            // <c/> with no content. Only string-typed cells materialize an empty value.
            if (dataType == CellValues.SharedString || dataType == CellValues.String)
                cellsCollection.ValueSlice.SetCellValueDuringLoad(cellAddress, string.Empty, false);

            if (styleMatchesInherited)
                EnsureStyleForBlankCell(cellsCollection, cellAddress, cellStyleValue);

            reader.Read(); // Move past <c/>.
            return;
        }

        reader.Read(); // Move into the cell's children (first <f>/<v>/<is> or </c>).
        LoadCellContentXml(reader, in context, dataType, cellAddress, cellStyleValue, ws, cellsCollection, cellMetaIndex);

        if (styleMatchesInherited)
            EnsureStyleForBlankCell(cellsCollection, cellAddress, cellStyleValue);

        reader.Read(); // Move past </c>.
    }

    private static void LoadCellContentXml(XmlReader reader, in SheetDataReadContext context,
        CellValues dataType, XLSheetPoint cellAddress, XLStyleValue cellStyleValue,
        XLWorksheet ws, XLCellsCollection cellsCollection, uint? cellMetaIndex)
    {
        // Positioned on the first child of <c> (an Element) or on </c> (an EndElement).
        var formula = IsMainElement(reader, "f")
            ? SetCellFormulaXml(reader, ws, cellAddress, context.SharedFormulasR1C1, cellMetaIndex, context.DynamicArrayCmIndexes)
            : null;

        var formulaInline = formula is not null;

        var cellHasValue = IsMainElement(reader, "v");
        var cellWasSetWithEmptyValue = false;
        if (cellHasValue)
        {
            var text = reader.ReadElementContentAsString(); // Reads <v> text and moves past </v>.
            SetCellValue(dataType, text, cellsCollection, cellAddress, cellStyleValue, ws,
                context.SharedStrings, formulaInline, context.NumberDataTypeCache);
        }
        else if (dataType == CellValues.SharedString || dataType == CellValues.String)
        {
            cellsCollection.ValueSlice.SetCellValueDuringLoad(cellAddress, string.Empty, formulaInline);
            cellWasSetWithEmptyValue = true;
        }

        if (formulaInline && (cellHasValue || cellWasSetWithEmptyValue))
            formula!.MarkClean(ws.Workbook);

        if (IsMainElement(reader, "is"))
            LoadInlineStringXml(reader, dataType, cellsCollection, cellAddress, ws);

        if (context.Use1904DateSystem && dataType == CellValues.Number)
            Adjust1904DateSystem(cellsCollection, cellAddress);

        // Ensure we land on </c>: skip any unrecognized trailing child elements.
        while (reader.NodeType == XmlNodeType.Element)
            reader.Skip();
    }

    private static XLCellFormula? SetCellFormulaXml(XmlReader reader, XLWorksheet ws, XLSheetPoint cellAddress,
        Dictionary<uint, string> sharedFormulasR1C1, uint? cellMetaIndex, HashSet<uint>? dynamicArrayCmIndexes)
    {
        string? typeAttr = null;
        string? refAttr = null;
        string? r1Attr = null;
        string? r2Attr = null;
        bool aca = false, dt2D = false, del1 = false, del2 = false, dtr = false;
        uint? sharedIndex = null;

        if (reader.HasAttributes)
        {
            while (reader.MoveToNextAttribute())
            {
                if (reader.NamespaceURI.Length != 0)
                    continue;

                switch (reader.LocalName)
                {
                    // bx attribute of cell formula is never used, per MS-OI29500 2.1.620.
                    case "t": typeAttr = reader.Value; break;
                    case "ref": refAttr = reader.Value; break;
                    case "si": sharedIndex = uint.Parse(reader.Value); break;
                    case "aca": aca = ParseXmlBool(reader.Value); break;
                    case "dt2D": dt2D = ParseXmlBool(reader.Value); break;
                    case "del1": del1 = ParseXmlBool(reader.Value); break;
                    case "del2": del2 = ParseXmlBool(reader.Value); break;
                    case "r1": r1Attr = reader.Value; break;
                    case "r2": r2Attr = reader.Value; break;
                    case "dtr": dtr = ParseXmlBool(reader.Value); break;
                }
            }

            reader.MoveToElement();
        }

        var formulaText = reader.ReadElementContentAsString(); // Reads <f> text and moves past </f>.

        var formulaType = typeAttr switch
        {
            "normal" => CellFormulaValues.Normal,
            "array" => CellFormulaValues.Array,
            "dataTable" => CellFormulaValues.DataTable,
            "shared" => CellFormulaValues.Shared,
            null => CellFormulaValues.Normal,
            _ => throw new NotSupportedException("Unknown formula type.")
        };

        var formulaSlice = ws.Internals.CellsCollection.FormulaSlice;
        XLCellFormula? formula = null;
        if (formulaType == CellFormulaValues.Normal)
        {
            formula = XLCellFormula.NormalA1(formulaText);
            formulaSlice.SetDuringLoad(cellAddress, formula);
        }
        else if (formulaType == CellFormulaValues.Array && refAttr is not null)
        {
            // Child cells of an array may have an array type but no ref (reserved for the master cell).
            var arrayArea = XLSheetRange.Parse(refAttr);
            var isDynamicArray = cellMetaIndex is { } cm &&
                                 dynamicArrayCmIndexes is not null &&
                                 dynamicArrayCmIndexes.Contains(cm);
            if (isDynamicArray)
            {
                formula = XLCellFormula.DynamicArrayA1(formulaText);
                formula.Range = arrayArea;
                formulaSlice.SetDuringLoad(cellAddress, formula);
            }
            else
            {
                formula = XLCellFormula.Array(formulaText, arrayArea, aca);
                formulaSlice.SetArray(arrayArea, formula);
            }
        }
        else if (formulaType == CellFormulaValues.Shared && sharedIndex is { } si)
        {
            formula = LoadSharedFormula(formulaText, cellAddress, si, sharedFormulasR1C1, formulaSlice);
        }
        else if (formulaType == CellFormulaValues.DataTable && refAttr is not null)
        {
            formula = LoadDataTableFormulaXml(refAttr, r1Attr, r2Attr, dt2D, del1, del2, dtr, cellAddress, formulaSlice);
        }

        return formula;
    }

    private static XLCellFormula LoadDataTableFormulaXml(string refAttr, string? r1Attr, string? r2Attr,
        bool is2D, bool input1Deleted, bool input2Deleted, bool isRowDataTable,
        XLSheetPoint cellAddress, FormulaSlice formulaSlice)
    {
        var dataTableArea = XLSheetRange.Parse(refAttr);
        var input1 = r1Attr is not null ? XLSheetPoint.Parse(r1Attr) : throw MissingRequiredAttr("r1");
        XLCellFormula formula;
        if (is2D)
        {
            var input2 = r2Attr is not null ? XLSheetPoint.Parse(r2Attr) : throw MissingRequiredAttr("r2");
            formula = XLCellFormula.DataTable2D(dataTableArea, input1, input1Deleted, input2, input2Deleted);
        }
        else
        {
            formula = XLCellFormula.DataTable1D(dataTableArea, input1, input1Deleted, isRowDataTable);
        }

        formulaSlice.SetDuringLoad(cellAddress, formula);
        return formula;
    }

    private static void LoadInlineStringXml(XmlReader reader, CellValues dataType,
        XLCellsCollection cellsCollection, XLSheetPoint cellAddress, XLWorksheet ws)
    {
        if (dataType != CellValues.InlineString)
        {
            reader.Skip(); // Moves past </is>.
            return;
        }

        cellsCollection.ValueSlice.SetShareString(cellAddress, false);

        // Rich text / phonetics are rare; reuse the SDK DOM parsing by materializing the <is>
        // subtree. ReadOuterXml emits the in-scope default namespace, so the fragment parses in
        // the spreadsheet-main namespace. ReadOuterXml also moves the reader past </is>.
        var inlineString = new InlineString(reader.ReadOuterXml());
        if (inlineString.Text is not null)
        {
            cellsCollection.ValueSlice.SetCellValue(cellAddress, inlineString.Text.Text.FixNewLines());
        }
        else if (inlineString.HasChildren)
        {
            var xlCell = new XLCell(ws, cellAddress);
            SetCellText(xlCell, inlineString);
        }
        else
        {
            cellsCollection.ValueSlice.SetCellValue(cellAddress, string.Empty);
        }
    }

    private static bool ParseXmlBool(string value)
        => value == "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Fast inherited style lookup using pre-cached row style. When the worksheet has no
    /// custom column styles (the common case for data sheets), returns the row style directly,
    /// avoiding a per-cell dictionary lookup into <c>ColumnsCollection</c>.
    /// </summary>
    private static XLStyleValue GetInheritedStyleFast(XLWorksheet ws, XLStyleValue rowStyle, int column, bool hasColumnStyles)
    {
        if (!hasColumnStyles)
            return rowStyle;

        var sheetStyle = ws.StyleValue;
        var colStyle = ws.Internals.ColumnsCollection.TryGetValue(column, out var c)
            ? c.StyleValue
            : sheetStyle;

        return XLStyleValue.Combine(sheetStyle, rowStyle, colStyle);
    }

    private static CellValues ParseCellDataType(string? typeAttribute)
    {
        return typeAttribute switch
        {
            "b" => CellValues.Boolean,
            "n" => CellValues.Number,
            "e" => CellValues.Error,
            "s" => CellValues.SharedString,
            "str" => CellValues.String,
            "inlineStr" => CellValues.InlineString,
            "d" => CellValues.Date,
            null => CellValues.Number,
            _ => throw new FormatException("Unknown cell type.")
        };
    }

    private static XLStyleValue ResolveCachedStyleValue(int styleIndex, StylesheetData styles,
        Dictionary<int, XLStyleValue> styleList)
    {
        if (!styleList.TryGetValue(styleIndex, out var cellStyleValue))
        {
            cellStyleValue = ResolveStyleValue(styleIndex, styles);
            styleList[styleIndex] = cellStyleValue;
        }

        return cellStyleValue;
    }

    /// <summary>
    /// Write cell value directly to <see cref="ValueSlice"/> during loading,
    /// bypassing <see cref="XLCell"/> allocation and <c>CalcEngine.MarkDirty</c>.
    /// An <see cref="XLCell"/> is only created for the rare rich-text shared-string path.
    /// </summary>
    internal static void SetCellValue(CellValues dataType, string? cellValue,
        XLCellsCollection cellsCollection, XLSheetPoint cellAddress, XLStyleValue cellStyleValue,
        XLWorksheet ws, SharedStringEntry[]? sharedStrings, bool inline,
        Dictionary<XLNumberFormatValue, XLDataType>? numberDataTypeCache = null)
    {
        // Only String writes an empty value when v is null.
        if (cellValue is null)
        {
            if (dataType == CellValues.String)
                cellsCollection.ValueSlice.SetCellValueDuringLoad(cellAddress, string.Empty, inline);
            return;
        }

        if (dataType == CellValues.Number)
            SetNumberCellValue(cellValue, cellsCollection, cellAddress, cellStyleValue, inline, numberDataTypeCache);
        else if (dataType == CellValues.SharedString)
            SetSharedStringCellValue(cellValue, cellsCollection, cellAddress, ws, sharedStrings, inline);
        else if (dataType == CellValues.String)
            cellsCollection.ValueSlice.SetCellValueDuringLoad(cellAddress, cellValue, inline);
        else if (dataType == CellValues.Boolean)
            SetBooleanCellValue(cellValue, cellsCollection, cellAddress, inline);
        else if (dataType == CellValues.Error)
            SetErrorCellValue(cellValue, cellsCollection, cellAddress, inline);
        else if (dataType == CellValues.Date)
            SetDateCellValue(cellValue, cellsCollection, cellAddress, inline);
    }

    /// <summary>
    /// Parses the cell value for normal or rich text.
    /// Input element should either be a shared string or inline string.
    /// </summary>
    internal static void SetCellText(XLCell xlCell, RstType element)
    {
        var runs = element.Elements<Run>();
        var hasRuns = false;
        foreach (var run in runs)
        {
            hasRuns = true;
            var runProperties = run.RunProperties;
            var text = run.Text!.InnerText.FixNewLines();

            if (runProperties == null)
                xlCell.GetRichText().AddText(text, xlCell.Style.Font);
            else
            {
                var rt = xlCell.GetRichText().AddText(text);
                var fontScheme = runProperties.Elements<FontScheme>().FirstOrDefault();
                if (fontScheme is { Val: not null })
                    rt.SetFontScheme(fontScheme.Val.Value.ToXLibur());

                OpenXmlHelper.LoadFont(runProperties, rt);
            }
        }

        if (!hasRuns)
            xlCell.SetOnlyValue(XmlEncoder.DecodeString(element.Text?.InnerText));

        LoadPhonetics(xlCell, element);
    }

    internal static void LoadColumns(StylesheetData styles, XLWorksheet ws, Columns columns)
    {
        var wsDefaultColumn =
            columns.Elements<Column>().FirstOrDefault(c => c.Max?.Value == XLHelper.MaxColumnNumber);

        if (wsDefaultColumn != null && wsDefaultColumn.Width != null)
            ws.ColumnWidth = wsDefaultColumn.Width - XLConstants.ColumnWidthOffset;

        var styleIndexDefault = wsDefaultColumn != null && wsDefaultColumn.Style != null
            ? int.Parse(wsDefaultColumn.Style!.InnerText!)
            : -1;
        if (styleIndexDefault >= 0)
            ApplyStyle(ws, styleIndexDefault, styles);

        foreach (var col in columns.Elements<Column>())
        {
            if (col.Max?.Value == XLHelper.MaxColumnNumber) continue;

            LoadColumn(col, ws, styles);
        }
    }

    internal static void ApplyStyle(IXLStylized xlStylized, int styleIndex, StylesheetData styles)
    {
        var xlStyleKey = XLStyle.Default.Key;
        LoadStyle(ref xlStyleKey, styleIndex, styles);

        // When loading columns, we must propagate the style to each column but not deeper. In other cases we do not propagate at all.
        if (xlStylized is IXLColumns columns)
        {
            columns.Cast<XLColumn>().ForEach(col => col.InnerStyle = new XLStyle(col, xlStyleKey));
        }
        else
        {
            xlStylized.InnerStyle = new XLStyle(xlStylized, xlStyleKey);
        }
    }

    /// <summary>
    /// Resolve a CellFormats style index to an interned <see cref="XLStyleValue"/>
    /// without creating an <see cref="XLStyle"/> wrapper or writing to any slice.
    /// </summary>
    internal static XLStyleValue ResolveStyleValue(int styleIndex, StylesheetData styles)
    {
        var xlStyleKey = XLStyle.Default.Key;
        LoadStyle(ref xlStyleKey, styleIndex, styles);
        return XLStyleValue.FromKey(ref xlStyleKey);
    }

    internal static void LoadStyle(ref XLStyleKey xlStyle, int styleIndex, StylesheetData styles)
    {
        if (styles.Stylesheet is not { CellFormats: not null } s)
            return; //No Stylesheet, no Styles

        var fills = styles.Fills!;
        var borders = styles.Borders!;
        var fonts = styles.Fonts!;
        var numberingFormats = styles.NumberingFormats;

        var cellFormat = (CellFormat)s.CellFormats.ElementAt(styleIndex);

        var xlIncludeQuotePrefix = OpenXmlHelper.GetBooleanValueAsBool(cellFormat.QuotePrefix, false);
        xlStyle = xlStyle with { IncludeQuotePrefix = xlIncludeQuotePrefix };

        if (cellFormat.ApplyProtection != null)
        {
            var protection = cellFormat.Protection;
            var xlProtection = XLProtectionValue.Default.Key;
            if (protection is not null)
                xlProtection = OpenXmlHelper.ProtectionToXLibur(protection, xlProtection);

            xlStyle = xlStyle with { Protection = xlProtection };
        }

        if (UInt32HasValue(cellFormat.FillId))
            xlStyle = LoadStyleFill(cellFormat, fills, xlStyle);

        var alignment = cellFormat.Alignment;
        if (alignment != null)
        {
            var xlAlignment = OpenXmlHelper.AlignmentToXLibur(alignment, xlStyle.Alignment);
            xlStyle = xlStyle with { Alignment = xlAlignment };
        }

        if (UInt32HasValue(cellFormat.BorderId))
            xlStyle = LoadStyleBorder(cellFormat.BorderId!.Value, borders, xlStyle);

        if (UInt32HasValue(cellFormat.FontId))
            xlStyle = LoadStyleFont(cellFormat.FontId!.Value, fonts, xlStyle);

        if (UInt32HasValue(cellFormat.NumberFormatId))
            xlStyle = LoadStyleNumberFormat(cellFormat, numberingFormats, xlStyle);
    }

    /// <summary>
    /// Resolves the <see cref="XLDataType"/> for a numeric cell's number format, memoizing the
    /// result per interned <see cref="XLNumberFormatValue"/>. <see cref="GetNumberDataType"/>
    /// scans the format string, so caching avoids repeating that work for every numeric cell that
    /// shares a format (the common case for data sheets). A <c>null</c> cache falls back to a
    /// direct computation.
    /// </summary>
    private static XLDataType GetCachedNumberDataType(XLNumberFormatValue numberFormat,
        Dictionary<XLNumberFormatValue, XLDataType>? cache)
    {
        if (cache is null)
            return GetNumberDataType(numberFormat);

        if (!cache.TryGetValue(numberFormat, out var dataType))
        {
            dataType = GetNumberDataType(numberFormat);
            cache[numberFormat] = dataType;
        }

        return dataType;
    }

    internal static XLDataType GetNumberDataType(XLNumberFormatValue numberFormat)
    {
        var numberFormatId = (XLPredefinedFormat.DateTime)numberFormat.NumberFormatId;
        var isTimeOnlyFormat = numberFormatId is
            Hour12MinutesAmPm or
            Hour12MinutesSecondsAmPm or
            Hour24Minutes or
            Hour24MinutesSeconds or
            MinutesSeconds or
            Hour12MinutesSeconds or
            MinutesSecondsMillis1;

        if (isTimeOnlyFormat)
            return XLDataType.TimeSpan;

        var isDateTimeFormat = numberFormatId is
            DayMonthYear4WithSlashes or
            DayMonthAbbrYear2WithDashes or
            DayMonthAbbrWithDash or
            MonthDayYear4WithDashesHour24Minutes;

        if (isDateTimeFormat)
            return XLDataType.DateTime;

        if (string.IsNullOrWhiteSpace(numberFormat.Format)) return XLDataType.Number;

        var dataType = GetDataTypeFromFormat(numberFormat.Format);
        return dataType ?? XLDataType.Number;
    }

    internal static XLDataType? GetDataTypeFromFormat(string format)
    {
        var length = format.Length;
        var i = 0;
        while (i < length)
        {
            var c = char.ToLowerInvariant(format[i]);
            switch (c)
            {
                case '"':
                    {
                        var closeIndex = format.IndexOf('"', i + 1);
                        if (closeIndex == -1)
                            return null;
                        i = closeIndex + 1;
                        break;
                    }
                case '[':
                    {
                        // #1742 We need to skip locale prefixes in DateTime formats [...]
                        var closeIndex = format.IndexOf(']', i + 1);
                        if (closeIndex == -1)
                            return null;
                        i = closeIndex + 1;
                        break;
                    }
                default:
                    {
                        var result = ClassifyFormatChar(c, format, i, length);
                        if (result.HasValue)
                            return result.Value;
                        i++;
                        break;
                    }
            }
        }

        return null;
    }

    private static XLDataType? ClassifyFormatChar(char c, string format, int i, int length)
    {
        return c switch
        {
            '0' or '#' or '?' => XLDataType.Number,
            'y' or 'd' => XLDataType.DateTime,
            'h' or 's' => XLDataType.TimeSpan,
            'm' => ResolveMonthOrMinute(format, i, length),
            _ => null
        };
    }

    internal static bool UInt32HasValue(UInt32Value? value)
    {
        return value != null && value.HasValue;
    }

    private static Exception MissingRequiredAttr(string attributeName)
    {
        throw new InvalidOperationException($"XML doesn't contain required attribute '{attributeName}'.");
    }

    private static void ApplyRowCustomProps(in RowProperties props, XLWorksheet ws,
        int rowIndex, StylesheetData styles)
    {
        var xlRow = ws.Row(rowIndex, false);

        if (props.Height is not null)
        {
            xlRow.Height = props.Height.Value;
        }
        else
        {
            xlRow.SetHeightNoFlag(ws.RowHeight);
        }

        if (props.DyDescent is not null)
            xlRow.DyDescent = props.DyDescent.Value;

        if (props.Hidden)
            xlRow.Hide();

        if (props.Collapsed)
            xlRow.Collapsed = true;

        if (props.OutlineLevel > 0)
            xlRow.OutlineLevel = props.OutlineLevel.Value;

        if (props.ShowPhonetic)
            xlRow.ShowPhonetic = true;

        if (props.CustomFormat)
        {
            if (props.StyleIndex is not null)
            {
                ApplyStyle(xlRow, props.StyleIndex.Value, styles);
            }
            else
            {
                xlRow.Style = ws.Style;
            }
        }
    }

    /// <summary>
    /// Adjusts the cell value for the 1904 date system by adding 1462 days.
    /// Must only be called for numeric serial date cells (t="n" or absent type attribute).
    /// ISO 8601 date cells (t="d") are absolute and must not be adjusted.
    /// </summary>
    private static void Adjust1904DateSystem(XLCellsCollection cellsCollection, XLSheetPoint cellAddress)
    {
        var cellValue = cellsCollection.ValueSlice.GetCellValue(cellAddress);
        if (cellValue.Type == XLDataType.DateTime)
        {
            cellsCollection.ValueSlice.SetCellValue(cellAddress, cellValue.GetDateTime().AddDays(1462));
        }
    }

    private static void EnsureStyleForBlankCell(XLCellsCollection cellsCollection, XLSheetPoint cellAddress,
        XLStyleValue cellStyleValue)
    {
        var hasOtherData = cellsCollection.ValueSlice.IsUsed(cellAddress)
                           || cellsCollection.FormulaSlice.IsUsed(cellAddress)
                           || cellsCollection.MiscSlice.IsUsed(cellAddress);

        if (!hasOtherData)
            cellsCollection.StyleSlice.SetNonDefault(cellAddress.Row, cellAddress.Column, cellStyleValue);
    }

    private static XLCellFormula LoadSharedFormula(string formulaText, XLSheetPoint cellAddress,
        uint sharedIndex, Dictionary<uint, string> sharedFormulasR1C1, FormulaSlice formulaSlice)
    {
        XLCellFormula formula;
        if (!sharedFormulasR1C1.TryGetValue(sharedIndex, out var sharedR1C1Formula))
        {
            formula = XLCellFormula.NormalA1(formulaText);
            formulaSlice.SetDuringLoad(cellAddress, formula);

            var formulaR1C1 = FormulaTransformation.SafeToR1C1(formulaText, cellAddress.Row, cellAddress.Column);
            sharedFormulasR1C1.Add(sharedIndex, formulaR1C1);
        }
        else
        {
            var sharedFormulaA1 =
                FormulaTransformation.SafeToA1(sharedR1C1Formula, cellAddress.Row, cellAddress.Column);
            formula = XLCellFormula.NormalA1(sharedFormulaA1);
            formulaSlice.SetDuringLoad(cellAddress, formula);
        }

        return formula;
    }

    private static void SetNumberCellValue(string cellValue, XLCellsCollection cellsCollection,
        XLSheetPoint cellAddress, XLStyleValue cellStyleValue, bool inline,
        Dictionary<XLNumberFormatValue, XLDataType>? numberDataTypeCache = null)
    {
        if (!TryParseOoxmlDouble(cellValue, out var number)) return;
        var numberDataType = GetCachedNumberDataType(cellStyleValue.NumberFormat, numberDataTypeCache);
        var cellNumber = numberDataType switch
        {
            XLDataType.DateTime => XLCellValue.FromSerialDateTime(number),
            XLDataType.TimeSpan => XLCellValue.FromSerialTimeSpan(number),
            _ => number
        };
        cellsCollection.ValueSlice.SetCellValueDuringLoad(cellAddress, cellNumber, inline);
    }

    private static void SetSharedStringCellValue(string cellValue, XLCellsCollection cellsCollection,
        XLSheetPoint cellAddress, XLWorksheet ws, SharedStringEntry[]? sharedStrings, bool inline)
    {
        if (TryParseOoxmlNonNegativeInt(cellValue, out var sharedStringId)
            && sharedStrings is not null && sharedStringId < sharedStrings.Length)
        {
            var entry = sharedStrings[sharedStringId];
            if (entry.IsRichText)
            {
                var xlCell = new XLCell(ws, cellAddress);
                SetCellText(xlCell, entry.RichText);
            }
            else
                cellsCollection.ValueSlice.SetCellValueDuringLoad(cellAddress, entry.PlainText, inline);
        }
        else
            cellsCollection.ValueSlice.SetCellValueDuringLoad(cellAddress, string.Empty, inline);
    }

    private static void SetBooleanCellValue(string cellValue, XLCellsCollection cellsCollection,
        XLSheetPoint cellAddress, bool inline)
    {
        var isTrue = string.Equals(cellValue, "1", StringComparison.Ordinal) ||
                     string.Equals(cellValue, "TRUE", StringComparison.OrdinalIgnoreCase);
        cellsCollection.ValueSlice.SetCellValueDuringLoad(cellAddress, isTrue, inline);
    }

    private static void SetErrorCellValue(string cellValue, XLCellsCollection cellsCollection,
        XLSheetPoint cellAddress, bool inline)
    {
        if (XLErrorParser.TryParseError(cellValue, out var error))
            cellsCollection.ValueSlice.SetCellValueDuringLoad(cellAddress, error, inline);
    }

    private static void SetDateCellValue(string cellValue, XLCellsCollection cellsCollection,
        XLSheetPoint cellAddress, bool inline)
    {
        var date = DateTime.ParseExact(cellValue, DateCellFormats,
            XLHelper.ParseCulture,
            DateTimeStyles.AllowLeadingWhite | DateTimeStyles.AllowTrailingWhite);
        cellsCollection.ValueSlice.SetCellValueDuringLoad(cellAddress, date, inline);
    }

    private static void LoadPhonetics(XLCell xlCell, RstType element)
    {
        var pp = element.Elements<PhoneticProperties>().FirstOrDefault();
        if (pp != null)
        {
            if (pp.Alignment != null)
                xlCell.GetRichText().Phonetics.Alignment = pp.Alignment.Value.ToXLibur();
            if (pp.Type != null)
                xlCell.GetRichText().Phonetics.Type = pp.Type.Value.ToXLibur();

            OpenXmlHelper.LoadFont(pp, xlCell.GetRichText().Phonetics);
        }

        foreach (var pr in element.Elements<PhoneticRun>())
        {
            var phoneticText = pr.Text!.InnerText.FixNewLines();
            var sb = (int)pr.BaseTextStartIndex!.Value;
            var eb = (int)pr.EndingBaseIndex!.Value;

            if (phoneticText.Length == 0 || sb >= eb)
                continue;

            xlCell.GetRichText().Phonetics.Add(phoneticText, sb, eb);
        }
    }

    private static void LoadColumn(Column col, XLWorksheet ws, StylesheetData styles)
    {
        var xlColumns = (XLColumns)ws.Columns((int)col.Min!.Value, (int)col.Max!.Value);
        if (col.Width != null)
        {
            var width = col.Width - XLConstants.ColumnWidthOffset;
            xlColumns.Width = width;
        }
        else
            xlColumns.Width = ws.ColumnWidth;

        if (col.Hidden != null && col.Hidden)
            xlColumns.Hide();

        if (col.Collapsed != null && col.Collapsed)
            xlColumns.CollapseOnly();

        if (col.OutlineLevel != null)
        {
            var outlineLevel = col.OutlineLevel;
            xlColumns.ForEach(c => c.OutlineLevel = outlineLevel);
        }

        var styleIndex = col.Style != null ? int.Parse(col.Style!.InnerText!) : -1;
        if (styleIndex >= 0)
        {
            ApplyStyle(xlColumns, styleIndex, styles);
        }
        else
        {
            xlColumns.Style = ws.Style;
        }
    }

    private static XLStyleKey LoadStyleFill(CellFormat cellFormat, Fills fills, XLStyleKey xlStyle)
    {
        var fill = (Fill)fills.ElementAt((int)cellFormat.FillId!.Value);
        if (fill.PatternFill == null) return xlStyle;
        var xlFill = new XLFill();
        OpenXmlHelper.LoadFill(fill, xlFill, differentialFillFormat: false);
        xlStyle = xlStyle with { Fill = xlFill.Key };

        return xlStyle;
    }

    private static XLStyleKey LoadStyleBorder(uint borderId, Borders borders, XLStyleKey xlStyle)
    {
        var border = (Border)borders.ElementAt((int)borderId);
        var xlBorder = OpenXmlHelper.BorderToXLibur(border, xlStyle.Border);
        xlStyle = xlStyle with { Border = xlBorder };
        return xlStyle;
    }

    private static XLStyleKey LoadStyleFont(uint fontId, Fonts fonts, XLStyleKey xlStyle)
    {
        var font = (Font)fonts.ElementAt((int)fontId);
        var xlFont = OpenXmlHelper.FontToXLibur(font, xlStyle.Font);
        xlStyle = xlStyle with { Font = xlFont };
        return xlStyle;
    }

    private static XLStyleKey LoadStyleNumberFormat(CellFormat cellFormat, NumberingFormats? numberingFormats,
        XLStyleKey xlStyle)
    {
        var numberFormatId = cellFormat.NumberFormatId;

        var formatCode = string.Empty;
        var numberingFormat =
            numberingFormats?.FirstOrDefault(nf =>
                ((NumberingFormat)nf).NumberFormatId != null &&
                ((NumberingFormat)nf).NumberFormatId!.Value == numberFormatId!) as NumberingFormat;

        if (numberingFormat != null && numberingFormat.FormatCode != null)
            formatCode = numberingFormat.FormatCode.Value!;

        var xlNumberFormat = xlStyle.NumberFormat;
        if (formatCode.Length > 0)
        {
            xlNumberFormat = XLNumberFormatKey.ForFormat(formatCode);
        }
        else
            xlNumberFormat = xlNumberFormat with { NumberFormatId = (int)numberFormatId!.Value };

        return xlStyle with { NumberFormat = xlNumberFormat };
    }

    // Exact power-of-10 divisors for up to 18 fraction digits. Using these instead of
    // Math.Pow(10, -n) ensures bit-exact results matching double.TryParse for simple decimals.
    private static readonly double[] Pow10 =
    [
        1E0, 1E1, 1E2, 1E3, 1E4, 1E5, 1E6, 1E7, 1E8, 1E9,
        1E10, 1E11, 1E12, 1E13, 1E14, 1E15, 1E16, 1E17, 1E18
    ];

    /// <summary>
    /// Fast-path double parser for OOXML numeric cell values. Handles the common forms
    /// produced by Excel: optional minus, digits with optional decimal point.
    /// Falls back to <see cref="double.TryParse(string, NumberStyles, IFormatProvider, out double)"/>
    /// for exponents, leading whitespace, leading '+', overflow, or any non-standard format.
    /// </summary>
    private static bool TryParseOoxmlDouble(string s, out double result)
    {
        var len = s.Length;
        if (len == 0)
        {
            result = 0;
            return false;
        }

        var i = 0;
        var first = s[0];

        // Leading whitespace, '+', or exponent numbers → fall back (rare in OOXML).
        if (first == ' ' || first == '+')
            return double.TryParse(s, XLHelper.NumberStyle, XLHelper.ParseCulture, out result);

        // Optional leading minus.
        bool negative = false;
        if (first == '-')
        {
            negative = true;
            i++;
            if (i >= len) { result = 0; return false; }
        }

        // Integer part — accumulate in long for exact precision.
        long mantissa = 0;
        var integerDigits = ScanDigits(s, ref i, ref mantissa);

        var fractionDigits = 0;

        // Fractional part.
        if (i < len && s[i] == '.')
        {
            i++;
            fractionDigits = ScanDigits(s, ref i, ref mantissa);
        }

        var totalDigits = integerDigits + fractionDigits;

        // Must have consumed at least one digit and ALL characters.
        // Any remaining chars (exponent, whitespace, etc.) → fall back.
        if (totalDigits == 0 || i != len || totalDigits > 18)
            return double.TryParse(s, XLHelper.NumberStyle, XLHelper.ParseCulture, out result);

        // Assemble the double using exact division.
        double d = fractionDigits == 0
            ? mantissa
            : mantissa / Pow10[fractionDigits];

        result = negative ? -d : d;
        return true;
    }

    /// <summary>
    /// Consume the run of ASCII digits starting at <paramref name="i"/>, folding them into
    /// <paramref name="mantissa"/>. Advances <paramref name="i"/> past the digits and returns the
    /// number of digits consumed.
    /// </summary>
    private static int ScanDigits(string s, ref int i, ref long mantissa)
    {
        var start = i;
        var len = s.Length;
        while (i < len && (uint)(s[i] - '0') <= 9)
        {
            mantissa = mantissa * 10 + (s[i] - '0');
            i++;
        }

        return i - start;
    }

    /// <summary>
    /// Fast-path parser for non-negative integer strings as found in OOXML shared string
    /// indices. Only accepts pure ASCII digit sequences with no whitespace or signs.
    /// </summary>
    /// <summary>
    /// Parse a row index attribute value. Row indices in OOXML are always positive integers.
    /// </summary>
    private static int ParseRowIndex(string s)
    {
        TryParseOoxmlNonNegativeInt(s, out var result);
        return result;
    }

    private static bool TryParseOoxmlNonNegativeInt(string s, out int result)
    {
        result = 0;
        var len = s.Length;
        if (len == 0)
            return false;

        // Guard: strings with >9 digits cannot fit in a non-negative int (max 2,147,483,647 = 10 digits).
        // Fall back to the full parser which handles overflow correctly.
        if (len > 9)
            return int.TryParse(s, XLHelper.NumberStyle, XLHelper.ParseCulture, out result);

        for (var i = 0; i < len; i++)
        {
            var digit = (uint)(s[i] - '0');
            if (digit > 9)
            {
                // Not a pure digit string — fall back to full parser.
                return int.TryParse(s, XLHelper.NumberStyle, XLHelper.ParseCulture, out result);
            }

            result = result * 10 + (int)digit;
        }

        return true;
    }

    private static XLDataType ResolveMonthOrMinute(string format, int i, int length)
    {
        for (var j = i + 1; j < length; j++)
        {
            var cj = char.ToLowerInvariant(format[j]);
            switch (cj)
            {
                case 'm':
                    continue;
                case 's':
                    return XLDataType.TimeSpan;
                case >= 'a' and <= 'z' or >= '0' and <= '9':
                    return XLDataType.DateTime;
            }
        }

        return XLDataType.DateTime;
    }
}
