using XLibur.Extensions;
using XLibur.Excel.CalcEngine.Visitors;
using XLibur.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
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
        bool use1904DateSystem)
    {
        public readonly StylesheetData Styles = styles;
        public readonly XLWorksheet Worksheet = worksheet;
        public readonly SharedStringEntry[]? SharedStrings = sharedStrings;
        public readonly Dictionary<uint, string> SharedFormulasR1C1 = sharedFormulasR1C1;
        public readonly Dictionary<int, XLStyleValue> StyleList = styleList;
        public readonly bool Use1904DateSystem = use1904DateSystem;
    }

    /// <summary>
    /// Mutable tracking state across rows during sheet data reading.
    /// </summary>
    internal struct SheetDataReadState
    {
        public int LastRow;
        public int LastColumnNumber;
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

    internal static void LoadRow(in SheetDataReadContext context, OpenXmlPartReader reader,
        ref SheetDataReadState state)
    {
        Debug.Assert(reader.LocalName == "row");

        var attributes = reader.Attributes;
        var rowIndexAttr = attributes.GetAttribute("r");
        var rowIndex = string.IsNullOrEmpty(rowIndexAttr) ? ++state.LastRow : int.Parse(rowIndexAttr);
        state.LastRow = rowIndex;

        var rowProps = new RowProperties(
            Height: attributes.GetDoubleAttribute("ht"),
            DyDescent: attributes.GetDoubleAttribute("dyDescent", OpenXmlConst.X14Ac2009SsNs),
            Hidden: attributes.GetBoolAttribute("hidden", false),
            Collapsed: attributes.GetBoolAttribute("collapsed", false),
            OutlineLevel: attributes.GetIntAttribute("outlineLevel"),
            ShowPhonetic: attributes.GetBoolAttribute("ph", false),
            CustomFormat: attributes.GetBoolAttribute("customFormat", false),
            StyleIndex: attributes.GetIntAttribute("s"));

        if (rowProps.HasCustomProps)
        {
            ApplyRowCustomProps(in rowProps, context.Worksheet, rowIndex, context.Styles);
        }

        state.LastColumnNumber = 0;

        // Move from the start element of 'row' forward. We can get cell, extList or end of row.
        reader.MoveAhead();

        while (reader.IsStartElement("c"))
        {
            LoadCell(in context, reader, rowIndex, ref state.LastColumnNumber);

            // Move from an end element of 'cell' either to next cell, extList start or end of row.
            reader.MoveAhead();
        }

        // In theory, row can also contain extList, just skip them.
        while (reader.IsStartElement("extLst"))
            reader.Skip();
    }

    private static void LoadCell(in SheetDataReadContext context, OpenXmlPartReader reader, int rowIndex,
        ref int lastColumnNumber)
    {
        Debug.Assert(reader is { LocalName: "c", IsStartElement: true });

        var attributes = reader.Attributes;

        var styleIndex = attributes.GetIntAttribute("s") ?? 0;

        var cellAddress = attributes.GetCellRefAttribute("r") ?? new XLSheetPoint(rowIndex, lastColumnNumber + 1);
        lastColumnNumber = cellAddress.Column;

        var dataType = attributes.GetAttribute("t") switch
        {
            "b" => CellValues.Boolean,
            "n" => CellValues.Number,
            "e" => CellValues.Error,
            "s" => CellValues.SharedString,
            "str" => CellValues.String,
            "inlineStr" => CellValues.InlineString,
            "d" => CellValues.Date,
            null => CellValues.Number,
            _ => throw new FormatException($"Unknown cell type.")
        };

        // Resolve style index to interned XLStyleValue (cached after first encounter).
        // The actual writing to StyleSlice is deferred to the end of the method, so we can
        // check whether other slices (value/formula/misc) already make this cell visible
        // to the SlicesEnumerator.  When the resolved style matches the inherited style
        // AND the cell has data in another slice, we skip the StyleSlice write — avoiding
        // per-row Lut allocation in the style slice for large data sheets.
        if (!context.StyleList.TryGetValue(styleIndex, out var cellStyleValue))
        {
            cellStyleValue = ResolveStyleValue(styleIndex, context.Styles);
            context.StyleList[styleIndex] = cellStyleValue;
        }

        // Write style directly to the slice — no XLCell allocation needed.
        var ws = context.Worksheet;
        var cellsCollection = ws.Internals.CellsCollection;
        var inherited = ws.GetInheritedStyleValue(cellAddress.Row, cellAddress.Column);
        var styleMatchesInherited = ReferenceEquals(cellStyleValue, inherited);
        if (!styleMatchesInherited)
            cellsCollection.StyleSlice.Set(cellAddress.Row, cellAddress.Column, cellStyleValue);

        LoadCellMisc(attributes, cellsCollection, cellAddress);

        // Move from the cell start element onwards.
        reader.MoveAhead();

        var cellHasFormula = reader.IsStartElement("f");
        XLCellFormula? formula = null;
        if (cellHasFormula)
        {
            formula = SetCellFormula(ws, cellAddress, reader, context.SharedFormulasR1C1);

            // Move from the end of the 'f' element.
            reader.MoveAhead();
        }

        // Unified code to load value. Value can be empty and only type specified (e.g. when the formula doesn't save values)
        // The string type is only for formulas, while shared string/inline string/date is only for pure cell values.
        var cellHasValue = reader.IsStartElement("v");
        if (cellHasValue)
        {
            SetCellValue(dataType, reader.GetText(), cellsCollection, cellAddress, cellStyleValue, ws,
                context.SharedStrings);

            // Skips all nodes of the 'v' element (has no child nodes) and moves to the first element after.
            reader.Skip();
        }
        else
        {
            // A string cell must contain at least an empty string.
            if (dataType.Equals(CellValues.SharedString) || dataType.Equals(CellValues.String))
                cellsCollection.ValueSlice.SetCellValue(cellAddress, string.Empty);
        }

        // If the cell doesn't contain value, we should invalidate it, otherwise rely on the stored value.
        // The value is likely more reliable. It should be set when cellFormula.CalculateCell is set or
        // when the value is missing. Formula can be null in some cases, e.g., slave cells of array formula.
        if (formula is not null && !cellHasValue)
        {
            formula.IsDirty = true;
        }

        // Inline text is dealt separately, because it is in a separate element.
        if (reader.IsStartElement("is"))
            LoadInlineString(dataType, cellsCollection, cellAddress, ws, reader);

        if (context.Use1904DateSystem)
            Adjust1904DateSystem(cellsCollection, cellAddress);

        if (styleMatchesInherited)
            EnsureStyleForBlankCell(cellsCollection, cellAddress, cellStyleValue);
    }

    private static XLCellFormula? SetCellFormula(XLWorksheet ws, XLSheetPoint cellAddress, OpenXmlPartReader reader,
        Dictionary<uint, string> sharedFormulasR1C1)
    {
        var attributes = reader.Attributes;
        var formulaSlice = ws.Internals.CellsCollection.FormulaSlice;
        var valueSlice = ws.Internals.CellsCollection.ValueSlice;

        // bx attribute of cell formula is not ever used, per MS-OI29500 2.1.620
        var formulaText = reader.GetText();
        var formulaType = attributes.GetAttribute("t") switch
        {
            "normal" => CellFormulaValues.Normal,
            "array" => CellFormulaValues.Array,
            "dataTable" => CellFormulaValues.DataTable,
            "shared" => CellFormulaValues.Shared,
            null => CellFormulaValues.Normal,
            _ => throw new NotSupportedException("Unknown formula type.")
        };

        // Always set the shareString flag to `false`, because the text result of
        // the formula is stored directly in the sheet, not the shared string table.
        XLCellFormula? formula = null;
        if (formulaType == CellFormulaValues.Normal)
        {
            formula = XLCellFormula.NormalA1(formulaText);
            formulaSlice.Set(cellAddress, formula);
            valueSlice.SetShareString(cellAddress, false);
        }
        else if (formulaType == CellFormulaValues.Array &&
                 attributes.GetRefAttribute("ref") is
                 {
                 } arrayArea) // Child cells of an array may have an array type, but not ref, that is reserved for the master cell
        {
            var aca = attributes.GetBoolAttribute("aca", false);

            // Because cells are read from top-to-bottom, from left-to-right, none of child cells have
            // a formula yet. Also, Excel doesn't allow change of array data, only through the parent formula.
            formula = XLCellFormula.Array(formulaText, arrayArea, aca);
            formulaSlice.SetArray(arrayArea, formula);

            for (var col = arrayArea.FirstPoint.Column; col <= arrayArea.LastPoint.Column; ++col)
            {
                for (var row = arrayArea.FirstPoint.Row; row <= arrayArea.LastPoint.Row; ++row)
                {
                    valueSlice.SetShareString(cellAddress, false);
                }
            }
        }
        else if (formulaType == CellFormulaValues.Shared && attributes.GetUintAttribute("si") is { } sharedIndex)
        {
            formula = LoadSharedFormula(formulaText, cellAddress, sharedIndex, sharedFormulasR1C1, formulaSlice);
            valueSlice.SetShareString(cellAddress, false);
        }
        else if (formulaType == CellFormulaValues.DataTable && attributes.GetRefAttribute("ref") is { } dataTableArea)
        {
            formula = LoadDataTableFormula(attributes, cellAddress, dataTableArea, formulaSlice);
            valueSlice.SetShareString(cellAddress, false);
        }

        // Go from the start of the 'f' element to the end of the 'f' element.
        reader.MoveAhead();

        return formula;
    }

    /// <summary>
    /// Write cell value directly to <see cref="ValueSlice"/> during loading,
    /// bypassing <see cref="XLCell"/> allocation and <c>CalcEngine.MarkDirty</c>.
    /// An <see cref="XLCell"/> is only created for the rare rich-text shared-string path.
    /// </summary>
    internal static void SetCellValue(CellValues dataType, string? cellValue,
        XLCellsCollection cellsCollection, XLSheetPoint cellAddress, XLStyleValue cellStyleValue,
        XLWorksheet ws, SharedStringEntry[]? sharedStrings)
    {
        // Only String writes an empty value when v is null.
        if (cellValue is null)
        {
            if (dataType == CellValues.String)
                cellsCollection.ValueSlice.SetCellValue(cellAddress, string.Empty);
            return;
        }

        if (dataType == CellValues.Number)
            SetNumberCellValue(cellValue, cellsCollection, cellAddress, cellStyleValue);
        else if (dataType == CellValues.SharedString)
            SetSharedStringCellValue(cellValue, cellsCollection, cellAddress, ws, sharedStrings);
        else if (dataType == CellValues.String)
            cellsCollection.ValueSlice.SetCellValue(cellAddress, cellValue);
        else if (dataType == CellValues.Boolean)
            SetBooleanCellValue(cellValue, cellsCollection, cellAddress);
        else if (dataType == CellValues.Error)
            SetErrorCellValue(cellValue, cellsCollection, cellAddress);
        else if (dataType == CellValues.Date)
            SetDateCellValue(cellValue, cellsCollection, cellAddress);
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

    private static void LoadCellMisc(ReadOnlyCollection<OpenXmlAttribute> attributes,
        XLCellsCollection cellsCollection, XLSheetPoint cellAddress)
    {
        var showPhonetic = attributes.GetBoolAttribute("ph", false);
        var cellMetaIndex = attributes.GetUintAttribute("cm");
        var valueMetaIndex = attributes.GetUintAttribute("vm");
        if (!showPhonetic && cellMetaIndex is null && valueMetaIndex is null) return;
        var misc = new XLMiscSliceContent
        {
            HasPhonetic = showPhonetic,
            CellMetaIndex = cellMetaIndex,
            ValueMetaIndex = valueMetaIndex
        };
        cellsCollection.MiscSlice.Set(cellAddress, in misc);
    }

    private static void LoadInlineString(CellValues dataType, XLCellsCollection cellsCollection,
        XLSheetPoint cellAddress, XLWorksheet ws, OpenXmlPartReader reader)
    {
        if (dataType == CellValues.InlineString)
        {
            cellsCollection.ValueSlice.SetShareString(cellAddress, false);
            if (reader.LoadCurrentElement() is RstType inlineString)
            {
                if (inlineString.Text is not null)
                    cellsCollection.ValueSlice.SetCellValue(cellAddress, inlineString.Text.Text.FixNewLines());
                else
                {
                    var xlCell = new XLCell(ws, cellAddress);
                    SetCellText(xlCell, inlineString);
                }
            }
            else
            {
                cellsCollection.ValueSlice.SetCellValue(cellAddress, string.Empty);
            }

            reader.MoveAhead();
        }
        else
        {
            reader.Skip();
        }
    }

    /// <summary>
    /// Adjusts the cell value for the 1904 date system by adding 1462 days to any date values.
    /// </summary>
    /// <param name="cellsCollection"></param>
    /// <param name="cellAddress"></param>
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
            cellsCollection.StyleSlice.Set(cellAddress.Row, cellAddress.Column, cellStyleValue);
    }

    private static XLCellFormula LoadSharedFormula(string formulaText, XLSheetPoint cellAddress,
        uint sharedIndex, Dictionary<uint, string> sharedFormulasR1C1, FormulaSlice formulaSlice)
    {
        XLCellFormula formula;
        if (!sharedFormulasR1C1.TryGetValue(sharedIndex, out var sharedR1C1Formula))
        {
            formula = XLCellFormula.NormalA1(formulaText);
            formulaSlice.Set(cellAddress, formula);

            var formulaR1C1 = FormulaTransformation.SafeToR1C1(formulaText, cellAddress.Row, cellAddress.Column);
            sharedFormulasR1C1.Add(sharedIndex, formulaR1C1);
        }
        else
        {
            var sharedFormulaA1 =
                FormulaTransformation.SafeToA1(sharedR1C1Formula, cellAddress.Row, cellAddress.Column);
            formula = XLCellFormula.NormalA1(sharedFormulaA1);
            formulaSlice.Set(cellAddress, formula);
        }

        return formula;
    }

    private static XLCellFormula LoadDataTableFormula(ReadOnlyCollection<OpenXmlAttribute> attributes,
        XLSheetPoint cellAddress, XLSheetRange dataTableArea, FormulaSlice formulaSlice)
    {
        var is2D = attributes.GetBoolAttribute("dt2D", false);
        var input1Deleted = attributes.GetBoolAttribute("del1", false);
        var input1 = attributes.GetCellRefAttribute("r1") ?? throw MissingRequiredAttr("r1");
        XLCellFormula formula;
        if (is2D)
        {
            var input2Deleted = attributes.GetBoolAttribute("del2", false);
            var input2 = attributes.GetCellRefAttribute("r2") ?? throw MissingRequiredAttr("r2");
            formula = XLCellFormula.DataTable2D(dataTableArea, input1, input1Deleted, input2, input2Deleted);
        }
        else
        {
            var isRowDataTable = attributes.GetBoolAttribute("dtr", false);
            formula = XLCellFormula.DataTable1D(dataTableArea, input1, input1Deleted, isRowDataTable);
        }

        formulaSlice.Set(cellAddress, formula);

        return formula;
    }

    private static void SetNumberCellValue(string cellValue, XLCellsCollection cellsCollection,
        XLSheetPoint cellAddress, XLStyleValue cellStyleValue)
    {
        if (!double.TryParse(cellValue, XLHelper.NumberStyle, XLHelper.ParseCulture, out var number)) return;
        var numberDataType = GetNumberDataType(cellStyleValue.NumberFormat);
        var cellNumber = numberDataType switch
        {
            XLDataType.DateTime => XLCellValue.FromSerialDateTime(number),
            XLDataType.TimeSpan => XLCellValue.FromSerialTimeSpan(number),
            _ => number
        };
        cellsCollection.ValueSlice.SetCellValue(cellAddress, cellNumber);
    }

    private static void SetSharedStringCellValue(string cellValue, XLCellsCollection cellsCollection,
        XLSheetPoint cellAddress, XLWorksheet ws, SharedStringEntry[]? sharedStrings)
    {
        if (int.TryParse(cellValue, XLHelper.NumberStyle, XLHelper.ParseCulture, out var sharedStringId)
            && sharedStrings is not null && sharedStringId >= 0 && sharedStringId < sharedStrings.Length)
        {
            var entry = sharedStrings[sharedStringId];
            if (entry.IsRichText)
            {
                var xlCell = new XLCell(ws, cellAddress);
                SetCellText(xlCell, entry.RichText);
            }
            else
                cellsCollection.ValueSlice.SetCellValue(cellAddress, entry.PlainText);
        }
        else
            cellsCollection.ValueSlice.SetCellValue(cellAddress, string.Empty);
    }

    private static void SetBooleanCellValue(string cellValue, XLCellsCollection cellsCollection,
        XLSheetPoint cellAddress)
    {
        var isTrue = string.Equals(cellValue, "1", StringComparison.Ordinal) ||
                     string.Equals(cellValue, "TRUE", StringComparison.OrdinalIgnoreCase);
        cellsCollection.ValueSlice.SetCellValue(cellAddress, isTrue);
    }

    private static void SetErrorCellValue(string cellValue, XLCellsCollection cellsCollection,
        XLSheetPoint cellAddress)
    {
        if (XLErrorParser.TryParseError(cellValue, out var error))
            cellsCollection.ValueSlice.SetCellValue(cellAddress, error);
    }

    private static void SetDateCellValue(string cellValue, XLCellsCollection cellsCollection,
        XLSheetPoint cellAddress)
    {
        var date = DateTime.ParseExact(cellValue, DateCellFormats,
            XLHelper.ParseCulture,
            DateTimeStyles.AllowLeadingWhite | DateTimeStyles.AllowTrailingWhite);
        cellsCollection.ValueSlice.SetCellValue(cellAddress, date);
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
