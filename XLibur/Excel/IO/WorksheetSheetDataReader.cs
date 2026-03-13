using XLibur.Extensions;
using XLibur.Excel.CalcEngine.Visitors;
using ClosedXML.Parser;
using XLibur.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
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
    private static readonly string[] DateCellFormats =
    [
        "yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff", // Format accepted by OpenXML SDK
        "yyyy-MM-ddTHH:mm", "yyyy-MM-dd" // Formats accepted by Excel.
    ];

    internal static void LoadRow(Stylesheet s, NumberingFormats? numberingFormats, Fills fills, Borders borders,
        Fonts fonts, XLWorksheet ws, SharedStringItem[]? sharedStrings,
        Dictionary<uint, string> sharedFormulasR1C1, Dictionary<int, IXLStyle> styleList,
        OpenXmlPartReader reader, ref int lastRow, ref int lastColumnNumber, bool use1904DateSystem)
    {
        Debug.Assert(reader.LocalName == "row");

        var attributes = reader.Attributes;
        var rowIndexAttr = attributes.GetAttribute("r");
        var rowIndex = string.IsNullOrEmpty(rowIndexAttr) ? ++lastRow : int.Parse(rowIndexAttr);

        var xlRow = ws.Row(rowIndex, false);

        var height = attributes.GetDoubleAttribute("ht");
        if (height is not null)
        {
            xlRow.Height = height.Value;
        }
        else
        {
            xlRow.Loading = true;
            xlRow.Height = ws.RowHeight;
            xlRow.Loading = false;
        }

        var dyDescent = attributes.GetDoubleAttribute("dyDescent", OpenXmlConst.X14Ac2009SsNs);
        if (dyDescent is not null)
            xlRow.DyDescent = dyDescent.Value;

        var hidden = attributes.GetBoolAttribute("hidden", false);
        if (hidden)
            xlRow.Hide();

        var collapsed = attributes.GetBoolAttribute("collapsed", false);
        if (collapsed)
            xlRow.Collapsed = true;

        var outlineLevel = attributes.GetIntAttribute("outlineLevel");
        if (outlineLevel is not null && outlineLevel.Value > 0)
            xlRow.OutlineLevel = outlineLevel.Value;

        var showPhonetic = attributes.GetBoolAttribute("ph", false);
        if (showPhonetic)
            xlRow.ShowPhonetic = true;

        var customFormat = attributes.GetBoolAttribute("customFormat", false);
        if (customFormat)
        {
            var styleIndex = attributes.GetIntAttribute("s");
            if (styleIndex is not null)
            {
                ApplyStyle(xlRow, styleIndex.Value, s, fills, borders, fonts, numberingFormats);
            }
            else
            {
                xlRow.Style = ws.Style;
            }
        }

        lastColumnNumber = 0;

        // Move from the start element of 'row' forward. We can get cell, extList or end of row.
        reader.MoveAhead();

        while (reader.IsStartElement("c"))
        {
            LoadCell(sharedStrings, s, numberingFormats, fills, borders, fonts, sharedFormulasR1C1, ws, styleList,
                reader, rowIndex, ref lastColumnNumber, use1904DateSystem);

            // Move from end element of 'cell' either to next cell, extList start or end of row.
            reader.MoveAhead();
        }

        // In theory, row can also contain extList, just skip them.
        while (reader.IsStartElement("extLst"))
            reader.Skip();
    }

    internal static void LoadCell(SharedStringItem[]? sharedStrings, Stylesheet s, NumberingFormats? numberingFormats,
        Fills fills, Borders borders, Fonts fonts, Dictionary<uint, string> sharedFormulasR1C1,
        XLWorksheet ws, Dictionary<int, IXLStyle> styleList, OpenXmlPartReader reader, int rowIndex,
        ref int lastColumnNumber, bool use1904DateSystem)
    {
        Debug.Assert(reader.LocalName == "c" && reader.IsStartElement);

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

        var xlCell = ws.Cell(cellAddress.Row, cellAddress.Column);

        if (styleList.TryGetValue(styleIndex, out var style))
        {
            xlCell.InnerStyle = style;
        }
        else
        {
            ApplyStyle(xlCell, styleIndex, s, fills, borders, fonts, numberingFormats);
        }

        var showPhonetic = attributes.GetBoolAttribute("ph", false);
        if (showPhonetic)
            xlCell.ShowPhonetic = true;

        var cellMetaIndex = attributes.GetUintAttribute("cm");
        if (cellMetaIndex is not null)
            xlCell.CellMetaIndex = cellMetaIndex.Value;

        var valueMetaIndex = attributes.GetUintAttribute("vm");
        if (valueMetaIndex is not null)
            xlCell.ValueMetaIndex = valueMetaIndex.Value;

        // Move from cell start element onwards.
        reader.MoveAhead();

        var cellHasFormula = reader.IsStartElement("f");
        XLCellFormula? formula = null;
        if (cellHasFormula)
        {
            formula = SetCellFormula(ws, cellAddress, reader, sharedFormulasR1C1);

            // Move from end of 'f' element.
            reader.MoveAhead();
        }

        // Unified code to load value. Value can be empty and only type specified (e.g. when formula doesn't save values)
        // String type is only for formulas, while shared string/inline string/date is only for pure cell values.
        var cellHasValue = reader.IsStartElement("v");
        if (cellHasValue)
        {
            SetCellValue(dataType, reader.GetText(), xlCell, sharedStrings);

            // Skips all nodes of the 'v' element (has no child nodes) and moves to the first element after.
            reader.Skip();
        }
        else
        {
            // A string cell must contain at least empty string.
            if (dataType.Equals(CellValues.SharedString) || dataType.Equals(CellValues.String))
                xlCell.SetOnlyValue(string.Empty);
        }

        // If the cell doesn't contain value, we should invalidate it, otherwise rely on the stored value.
        // The value is likely more reliable. It should be set when cellFormula.CalculateCell is set or
        // when value is missing. Formula can be null in some cases, e.g. slave cells of array formula.
        if (formula is not null && !cellHasValue)
        {
            formula.IsDirty = true;
        }

        // Inline text is dealt separately, because it is in a separate element.
        var cellHasInlineString = reader.IsStartElement("is");
        if (cellHasInlineString)
        {
            if (dataType == CellValues.InlineString)
            {
                xlCell.ShareString = false;
                var inlineString = reader.LoadCurrentElement() as RstType;
                if (inlineString is not null)
                {
                    if (inlineString.Text is not null)
                        xlCell.SetOnlyValue(inlineString.Text.Text.FixNewLines());
                    else
                        SetCellText(xlCell, inlineString);
                }
                else
                {
                    xlCell.SetOnlyValue(string.Empty);
                }

                // Move from end 'is' element to the end of a 'c' element.
                reader.MoveAhead();
            }
            else
            {
                // Move to the first node after end of 'is' element, which should be end of cell.
                reader.Skip();
            }
        }

        if (use1904DateSystem && xlCell.DataType == XLDataType.DateTime)
        {
            // Internally XLibur stores cells as standard 1900-based style
            // so if a workbook is in 1904-format, we do that adjustment here and when saving.
            xlCell.SetOnlyValue(xlCell.GetDateTime().AddDays(1462));
        }

        if (!styleList.ContainsKey(styleIndex))
            styleList.Add(styleIndex, xlCell.Style);
    }

    internal static XLCellFormula? SetCellFormula(XLWorksheet ws, XLSheetPoint cellAddress, OpenXmlPartReader reader,
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

        // Always set shareString flag to `false`, because the text result of
        // formula is stored directly in the sheet, not shared string table.
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
                 } arrayArea) // Child cells of an array may have array type, but not ref, that is reserved for master cell
        {
            var aca = attributes.GetBoolAttribute("aca", false);

            // Because cells are read from top-to-bottom, from left-to-right, none of child cells have
            // a formula yet. Also, Excel doesn't allow change of array data, only through parent formula.
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
            // Shared formulas are rather limited in use and parsing, even by Excel
            // https://stackoverflow.com/questions/54654993. Therefore we accept them,
            // but don't output them. Shared formula is created, when user in Excel
            // takes a supported formula and drags it to more cells.
            if (!sharedFormulasR1C1.TryGetValue(sharedIndex, out var sharedR1C1Formula))
            {
                // Spec: The first formula in a group of shared formulas is saved
                // in the f element. This is considered the 'master' formula cell.
                formula = XLCellFormula.NormalA1(formulaText);
                formulaSlice.Set(cellAddress, formula);

                // The key reason why Excel hates shared formulas is likely relative addressing and the messy situation it creates
                var formulaR1C1 = FormulaTransformation.SafeToR1C1(formulaText, cellAddress.Row, cellAddress.Column);
                sharedFormulasR1C1.Add(sharedIndex, formulaR1C1);
            }
            else
            {
                // Spec: The formula expression for a cell that is specified to be part of a shared formula
                // (and is not the master) shall be ignored, and the master formula shall override.
                var sharedFormulaA1 = FormulaTransformation.SafeToA1(sharedR1C1Formula, cellAddress.Row, cellAddress.Column);
                formula = XLCellFormula.NormalA1(sharedFormulaA1);
                formulaSlice.Set(cellAddress, formula);
            }

            valueSlice.SetShareString(cellAddress, false);
        }
        else if (formulaType == CellFormulaValues.DataTable && attributes.GetRefAttribute("ref") is { } dataTableArea)
        {
            var is2D = attributes.GetBoolAttribute("dt2D", false);
            var input1Deleted = attributes.GetBoolAttribute("del1", false);
            var input1 = attributes.GetCellRefAttribute("r1") ?? throw MissingRequiredAttr("r1");
            if (is2D)
            {
                // Input 2 is only used for 2D tables
                var input2Deleted = attributes.GetBoolAttribute("del2", false);
                var input2 = attributes.GetCellRefAttribute("r2") ?? throw MissingRequiredAttr("r2");
                formula = XLCellFormula.DataTable2D(dataTableArea, input1, input1Deleted, input2, input2Deleted);
                formulaSlice.Set(cellAddress, formula);
            }
            else
            {
                var isRowDataTable = attributes.GetBoolAttribute("dtr", false);
                formula = XLCellFormula.DataTable1D(dataTableArea, input1, input1Deleted, isRowDataTable);
                formulaSlice.Set(cellAddress, formula);
            }

            valueSlice.SetShareString(cellAddress, false);
        }

        // Go from start of 'f' element to the end of 'f' element.
        reader.MoveAhead();

        return formula;
    }

    internal static void SetCellValue(CellValues dataType, string? cellValue, XLCell xlCell,
        SharedStringItem[]? sharedStrings)
    {
        if (dataType == CellValues.Number)
        {
            // XLCell is by default blank, so no need to set it.
            if (cellValue is not null &&
                double.TryParse(cellValue, XLHelper.NumberStyle, XLHelper.ParseCulture, out var number))
            {
                var numberDataType = GetNumberDataType(xlCell.StyleValue.NumberFormat);
                var cellNumber = numberDataType switch
                {
                    XLDataType.DateTime => XLCellValue.FromSerialDateTime(number),
                    XLDataType.TimeSpan => XLCellValue.FromSerialTimeSpan(number),
                    _ => number // Normal number
                };
                xlCell.SetOnlyValue(cellNumber);
            }
        }
        else if (dataType == CellValues.SharedString)
        {
            if (cellValue is not null
                && int.TryParse(cellValue, XLHelper.NumberStyle, XLHelper.ParseCulture, out var sharedStringId)
                && sharedStrings is not null && sharedStringId >= 0 && sharedStringId < sharedStrings.Length)
            {
                var sharedString = sharedStrings[sharedStringId];

                SetCellText(xlCell, sharedString);
            }
            else
                xlCell.SetOnlyValue(string.Empty);
        }
        else if (dataType == CellValues.String) // A plain string that is a result of a formula calculation
        {
            xlCell.SetOnlyValue(cellValue ?? string.Empty);
        }
        else if (dataType == CellValues.Boolean)
        {
            if (cellValue is not null)
            {
                var isTrue = string.Equals(cellValue, "1", StringComparison.Ordinal) ||
                             string.Equals(cellValue, "TRUE", StringComparison.OrdinalIgnoreCase);
                xlCell.SetOnlyValue(isTrue);
            }
        }
        else if (dataType == CellValues.Error)
        {
            if (cellValue is not null && XLErrorParser.TryParseError(cellValue, out var error))
                xlCell.SetOnlyValue(error);
        }
        else if (dataType == CellValues.Date)
        {
            // Technically, cell can contain date as ISO8601 string, but not rarely used due
            // to inconsistencies between ISO and serial date time representation.
            if (cellValue is not null)
            {
                var date = DateTime.ParseExact(cellValue, DateCellFormats,
                    XLHelper.ParseCulture,
                    DateTimeStyles.AllowLeadingWhite | DateTimeStyles.AllowTrailingWhite);
                xlCell.SetOnlyValue(date);
            }
        }
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

        // Load phonetic properties
        var phoneticProperties = element.Elements<PhoneticProperties>();
        var pp = phoneticProperties.FirstOrDefault();
        if (pp != null)
        {
            if (pp.Alignment != null)
                xlCell.GetRichText().Phonetics.Alignment = pp.Alignment.Value.ToXLibur();
            if (pp.Type != null)
                xlCell.GetRichText().Phonetics.Type = pp.Type.Value.ToXLibur();

            OpenXmlHelper.LoadFont(pp, xlCell.GetRichText().Phonetics);
        }

        // Load phonetic runs
        var phoneticRuns = element.Elements<PhoneticRun>();
        foreach (var pr in phoneticRuns)
        {
            xlCell.GetRichText().Phonetics.Add(pr.Text!.InnerText.FixNewLines(), (int)pr.BaseTextStartIndex!.Value,
                (int)pr.EndingBaseIndex!.Value);
        }
    }

    internal static void LoadColumns(Stylesheet s, NumberingFormats? numberingFormats, Fills fills, Borders borders,
        Fonts fonts, XLWorksheet ws, Columns columns)
    {
        if (columns == null) return;

        var wsDefaultColumn =
            columns.Elements<Column>().FirstOrDefault(c => c.Max?.Value == XLHelper.MaxColumnNumber);

        if (wsDefaultColumn != null && wsDefaultColumn.Width != null)
            ws.ColumnWidth = wsDefaultColumn.Width - XLConstants.ColumnWidthOffset;

        var styleIndexDefault = wsDefaultColumn != null && wsDefaultColumn.Style != null
            ? int.Parse(wsDefaultColumn.Style!.InnerText!)
            : -1;
        if (styleIndexDefault >= 0)
            ApplyStyle(ws, styleIndexDefault, s, fills, borders, fonts, numberingFormats);

        foreach (var col in columns.Elements<Column>())
        {
            //IXLStylized toApply;
            if (col.Max?.Value == XLHelper.MaxColumnNumber) continue;

            var xlColumns = (XLColumns)ws.Columns((int)col.Min!.Value, (int)col.Max!.Value);
            if (col.Width != null)
            {
                var width = col.Width - XLConstants.ColumnWidthOffset;
                //if (width < 0) width = 0;
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
                ApplyStyle(xlColumns, styleIndex, s, fills, borders, fonts, numberingFormats);
            }
            else
            {
                xlColumns.Style = ws.Style;
            }
        }
    }

    internal static void ApplyStyle(IXLStylized xlStylized, int styleIndex, Stylesheet s, Fills fills, Borders borders,
        Fonts fonts, NumberingFormats? numberingFormats)
    {
        var xlStyleKey = XLStyle.Default.Key;
        LoadStyle(ref xlStyleKey, styleIndex, s, fills, borders, fonts, numberingFormats);

        // When loading columns we must propagate style to each column but not deeper. In other cases we do not propagate at all.
        if (xlStylized is IXLColumns columns)
        {
            columns.Cast<XLColumn>().ForEach(col => col.InnerStyle = new XLStyle(col, xlStyleKey));
        }
        else
        {
            xlStylized.InnerStyle = new XLStyle(xlStylized, xlStyleKey);
        }
    }

    internal static void LoadStyle(ref XLStyleKey xlStyle, int styleIndex, Stylesheet s, Fills fills, Borders borders,
        Fonts fonts, NumberingFormats? numberingFormats)
    {
        if (s == null || s.CellFormats is null) return; //No Stylesheet, no Styles

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
        {
            var fill = (Fill)fills.ElementAt((int)cellFormat.FillId!.Value);
            if (fill.PatternFill != null)
            {
                var xlFill = new XLFill();
                OpenXmlHelper.LoadFill(fill, xlFill, differentialFillFormat: false);
                xlStyle = xlStyle with { Fill = xlFill.Key };
            }
        }

        var alignment = cellFormat.Alignment;
        if (alignment != null)
        {
            var xlAlignment = OpenXmlHelper.AlignmentToXLibur(alignment, xlStyle.Alignment);
            xlStyle = xlStyle with { Alignment = xlAlignment };
        }

        if (UInt32HasValue(cellFormat.BorderId))
        {
            var borderId = cellFormat.BorderId!.Value;
            var border = (Border)borders.ElementAt((int)borderId);
            if (border is not null)
            {
                var xlBorder = OpenXmlHelper.BorderToXLibur(border, xlStyle.Border);
                xlStyle = xlStyle with { Border = xlBorder };
            }
        }

        if (UInt32HasValue(cellFormat.FontId))
        {
            var fontId = cellFormat.FontId;
            var font = (Font)fonts.ElementAt((int)fontId!.Value);
            if (font is not null)
            {
                var xlFont = OpenXmlHelper.FontToXLibur(font, xlStyle.Font);
                xlStyle = xlStyle with { Font = xlFont };
            }
        }

        if (UInt32HasValue(cellFormat.NumberFormatId))
        {
            var numberFormatId = cellFormat.NumberFormatId;

            var formatCode = string.Empty;
            if (numberingFormats != null)
            {
                var numberingFormat =
                    numberingFormats.FirstOrDefault(nf =>
                        ((NumberingFormat)nf).NumberFormatId != null &&
                        ((NumberingFormat)nf).NumberFormatId!.Value == numberFormatId!) as NumberingFormat;

                if (numberingFormat != null && numberingFormat.FormatCode != null)
                    formatCode = numberingFormat.FormatCode.Value!;
            }

            var xlNumberFormat = xlStyle.NumberFormat;
            if (formatCode.Length > 0)
            {
                xlNumberFormat = XLNumberFormatKey.ForFormat(formatCode);
            }
            else
                xlNumberFormat = xlNumberFormat with { NumberFormatId = (int)numberFormatId!.Value };

            xlStyle = xlStyle with { NumberFormat = xlNumberFormat };
        }
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

        if (!string.IsNullOrWhiteSpace(numberFormat.Format))
        {
            var dataType = GetDataTypeFromFormat(numberFormat.Format);
            return dataType ?? XLDataType.Number;
        }

        return XLDataType.Number;
    }

    internal static XLDataType? GetDataTypeFromFormat(string format)
    {
        var length = format.Length;
        var f = format.ToLower();
        for (var i = 0; i < length; i++)
        {
            var c = f[i];
            if (c == '"')
                i = f.IndexOf('"', i + 1);
            else if (c == '[')
            {
                // #1742 We need to skip locale prefixes in DateTime formats [...]
                i = f.IndexOf(']', i + 1);
                if (i == -1)
                    return null;
            }
            else if (c == '0' || c == '#' || c == '?')
                return XLDataType.Number;
            else if (c == 'y' || c == 'd')
                return XLDataType.DateTime;
            else if (c == 'h' || c == 's')
                return XLDataType.TimeSpan;
            else if (c == 'm')
            {
                // Excel treats "m" immediately after "hh" or "h" or immediately before "ss" or "s" as minutes, otherwise as a month value
                // We can ignore the "hh" or "h" prefixes as these would have been detected by the preceding condition above.
                // So we just need to make sure any 'm' is followed immediately by "ss" or "s" (excluding placeholders) to detect a timespan value
                for (var j = i + 1; j < length; j++)
                {
                    if (f[j] == 'm')
                        continue;
                    if (f[j] == 's')
                        return XLDataType.TimeSpan;
                    if ((f[j] >= 'a' && f[j] <= 'z') || (f[j] >= '0' && f[j] <= '9'))
                        return XLDataType.DateTime;
                }

                return XLDataType.DateTime;
            }
        }

        return null;
    }

    internal static bool UInt32HasValue(UInt32Value? value)
    {
        return value != null && value.HasValue;
    }

    internal static Exception MissingRequiredAttr(string attributeName)
    {
        throw new InvalidOperationException($"XML doesn't contain required attribute '{attributeName}'.");
    }
}
