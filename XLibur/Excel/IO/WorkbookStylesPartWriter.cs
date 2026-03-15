using XLibur.Utils;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using static XLibur.Excel.XLWorkbook;
using EnumerableExtensions = XLibur.Extensions.EnumerableExtensions;

namespace XLibur.Excel.IO;

internal static class WorkbookStylesPartWriter
{
    internal static void GenerateContent(WorkbookStylesPart workbookStylesPart, XLWorkbook workbook,
        SaveContext context)
    {
        var defaultStyle = DefaultStyleValue;

        if (!context.SharedFonts.ContainsKey(defaultStyle.Font))
            context.SharedFonts.Add(defaultStyle.Font, new FontInfo { FontId = 0, Font = defaultStyle.Font });

        workbookStylesPart.Stylesheet ??= new Stylesheet();
        workbookStylesPart.Stylesheet!.CellStyles ??= new CellStyles();

        var defaultFormatId = ResolveDefaultFormatId(workbookStylesPart);

        context.SharedStyles.Add(defaultStyle,
            new StyleInfo
            {
                StyleId = defaultFormatId,
                Style = defaultStyle,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                IncludeQuotePrefix = false,
                NumberFormatId = 0
            });

        var (xlStyles, pivotCustomFormats) = CollectWorkbookStyles(workbook);

        uint fontCount = 1;
        uint fillCount = 3;
        uint borderCount = 1;

        foreach (var font in xlStyles.Select(s => s.Font).Distinct())
        {
            if (!context.SharedFonts.ContainsKey(font))
                context.SharedFonts.Add(font, new FontInfo { FontId = fontCount++, Font = font });
        }

        var sharedFills = xlStyles.Select(s => s.Fill).Distinct().ToDictionary(
            f => f, f => new FillInfo { FillId = fillCount++, Fill = f });

        var sharedBorders = xlStyles.Select(s => s.Border).Distinct().ToDictionary(
            b => b, b => new BorderInfo { BorderId = borderCount++, Border = b });

        var customNumberFormats = CollectCustomNumberFormats(xlStyles, pivotCustomFormats);

        var allSharedNumberFormats = ResolveNumberFormats(workbookStylesPart, customNumberFormats, defaultFormatId);
        foreach (var nf in allSharedNumberFormats)
            context.SharedNumberFormats.Add(nf.Key, nf.Value);

        ResolveFonts(workbookStylesPart, context);
        var allSharedFills = ResolveFills(workbookStylesPart, sharedFills);
        var allSharedBorders = ResolveBorders(workbookStylesPart, sharedBorders);

        BuildSharedStyleMappings(context, xlStyles, allSharedNumberFormats, allSharedFills, allSharedBorders);

        ResolveCellStyleFormats(workbookStylesPart, context);
        ResolveRest(workbookStylesPart, context);

        if (!workbookStylesPart.Stylesheet!.CellStyles.Elements<CellStyle>().Any(c =>
                c.BuiltinId != null && c.BuiltinId.HasValue && c.BuiltinId.Value == 0U))
            workbookStylesPart.Stylesheet!.CellStyles.AppendChild(new CellStyle
                { Name = "Normal", FormatId = defaultFormatId, BuiltinId = 0U });

        workbookStylesPart.Stylesheet!.CellStyles.Count = (uint)workbookStylesPart.Stylesheet!.CellStyles.Count();

        RemapStyleIds(workbookStylesPart, context);

        AddDifferentialFormats(workbookStylesPart, workbook, context);
    }

    /// <summary>
    /// Determine the default workbook style by looking for the style with builtInId = 0.
    /// </summary>
    private static uint ResolveDefaultFormatId(WorkbookStylesPart workbookStylesPart)
    {
        var cellStyles = workbookStylesPart.Stylesheet!.CellStyles!;

        if (cellStyles.Elements<CellStyle>()
            .Any(c => c.BuiltinId != null && c.BuiltinId.HasValue && c.BuiltinId.Value == 0))
        {
            // Possible to have duplicate default cell styles - occurs when file gets saved under different cultures.
            // We prefer the style that is named Normal
            var normalCellStyles = cellStyles.Elements<CellStyle>()
                .Where(c => c.BuiltinId != null && c.BuiltinId.HasValue && c.BuiltinId.Value == 0)
                .OrderBy(c => c.Name != null && c.Name.HasValue && c.Name.Value == "Normal");

            return normalCellStyles.Last().FormatId!.Value;
        }

        if (cellStyles.Elements<CellStyle>().Any())
            return cellStyles.Elements<CellStyle>().Max(c => c.FormatId!.Value) + 1;

        return 0;
    }

    /// <summary>
    /// Collect all distinct styles and pivot table custom number formats from every worksheet.
    /// </summary>
    private static (HashSet<XLStyleValue> styles, HashSet<string> pivotCustomFormats) CollectWorkbookStyles(
        XLWorkbook workbook)
    {
        var pivotCustomFormats = new HashSet<string>();
        var styles = new HashSet<XLStyleValue>();

        foreach (var worksheet in workbook.WorksheetsInternal)
        {
            styles.Add(worksheet.StyleValue);

            foreach (var s in worksheet.Internals.ColumnsCollection.Select(c => c.Value.StyleValue))
                styles.Add(s);

            foreach (var s in worksheet.Internals.RowsCollection.Select(r => r.Value.StyleValue))
                styles.Add(s);

            foreach (var c in worksheet.Internals.CellsCollection.GetCells())
                styles.Add(c.StyleValue);

            var xlPivotTableDataFieldFormats = worksheet.PivotTables
                .SelectMany<XLPivotTable, XLPivotDataField>(pt => pt.DataFields)
                .Where(x => x.NumberFormatValue is not null && !string.IsNullOrEmpty(x.NumberFormatValue.Format))
                .Select(x => x.NumberFormatValue!.Format);
            pivotCustomFormats.UnionWith(xlPivotTableDataFieldFormats);

            var xlPivotTableFieldFormats = worksheet.PivotTables
                .SelectMany<XLPivotTable, XLPivotTableField>(pt => pt.PivotFields)
                .Where(x => x.NumberFormatValue is not null && !string.IsNullOrEmpty(x.NumberFormatValue.Format))
                .Select(x => x.NumberFormatValue!.Format);
            pivotCustomFormats.UnionWith(xlPivotTableFieldFormats);
        }

        return (styles, pivotCustomFormats);
    }

    private static HashSet<XLNumberFormatValue> CollectCustomNumberFormats(
        HashSet<XLStyleValue> xlStyles,
        HashSet<string> pivotCustomFormats)
    {
        var customNumberFormats = xlStyles
            .Select(s => s.NumberFormat)
            .Distinct()
            .Where(nf => nf.NumberFormatId == -1)
            .ToHashSet();

        foreach (var pivotNumberFormat in pivotCustomFormats)
        {
            var numberFormatKey = XLNumberFormatKey.ForFormat(pivotNumberFormat);
            var numberFormat = XLNumberFormatValue.FromKey(ref numberFormatKey);
            customNumberFormats.Add(numberFormat);
        }

        return customNumberFormats;
    }

    /// <summary>
    /// Map each collected style to a <see cref="StyleInfo"/> with resolved font/fill/border/number-format IDs.
    /// </summary>
    private static void BuildSharedStyleMappings(
        SaveContext context,
        HashSet<XLStyleValue> xlStyles,
        Dictionary<XLNumberFormatValue, NumberFormatInfo> sharedNumberFormats,
        Dictionary<XLFillValue, FillInfo> sharedFills,
        Dictionary<XLBorderValue, BorderInfo> sharedBorders)
    {
        uint styleCount = 1;
        foreach (var xlStyle in xlStyles)
        {
            var numberFormatId = xlStyle.NumberFormat.NumberFormatId >= 0
                ? xlStyle.NumberFormat.NumberFormatId
                : sharedNumberFormats[xlStyle.NumberFormat].NumberFormatId;

            if (!context.SharedStyles.ContainsKey(xlStyle))
                context.SharedStyles.Add(xlStyle,
                    new StyleInfo
                    {
                        StyleId = styleCount++,
                        Style = xlStyle,
                        FontId = context.SharedFonts[xlStyle.Font].FontId,
                        FillId = sharedFills[xlStyle.Fill].FillId,
                        BorderId = sharedBorders[xlStyle.Border].BorderId,
                        NumberFormatId = numberFormatId,
                        IncludeQuotePrefix = xlStyle.IncludeQuotePrefix
                    });
        }
    }

    /// <summary>
    /// Match each shared style against the CellFormats in the part and assign the final style ID.
    /// </summary>
    private static void RemapStyleIds(WorkbookStylesPart workbookStylesPart, SaveContext context)
    {
        var newSharedStyles = new Dictionary<XLStyleValue, StyleInfo>();
        foreach (var ss in context.SharedStyles)
        {
            var styleId = -1;
            foreach (var openXmlElement in workbookStylesPart.Stylesheet!.CellFormats!)
            {
                var f = (CellFormat)openXmlElement;
                styleId++;
                if (CellFormatsAreEqual(f, ss.Value, compareAlignment: true))
                    break;
            }

            if (styleId == -1)
                styleId = 0;
            var si = ss.Value;
            si.StyleId = (uint)styleId;
            newSharedStyles.Add(ss.Key, si);
        }

        context.SharedStyles.Clear();
        EnumerableExtensions.ForEach(newSharedStyles, kp => context.SharedStyles.Add(kp.Key, kp.Value));
    }

    /// <summary>
    /// Populates the differential formats that are currently in the file to the SaveContext
    /// </summary>
    private static void AddDifferentialFormats(WorkbookStylesPart workbookStylesPart, XLWorkbook workbook,
        SaveContext context)
    {
        workbookStylesPart.Stylesheet!.DifferentialFormats ??= new DifferentialFormats();

        var differentialFormats = workbookStylesPart.Stylesheet!.DifferentialFormats;
        differentialFormats.RemoveAllChildren();
        FillDifferentialFormatsCollection(differentialFormats, context.DifferentialFormats);

        foreach (var ws in workbook.WorksheetsInternal)
        {
            AddConditionalFormatDxfs(differentialFormats, ws, context);
            AddTableFieldDxfs(differentialFormats, ws, context);
            AddPivotTableDxfs(differentialFormats, ws, context);
            AddAutoFilterColorFilterDxfs(differentialFormats, ws, context);
        }

        differentialFormats.Count = (uint)differentialFormats.Count();
        if (differentialFormats.Count == 0)
            workbookStylesPart.Stylesheet!.DifferentialFormats = null;
    }

    private static void AddConditionalFormatDxfs(DifferentialFormats differentialFormats, XLWorksheet ws,
        SaveContext context)
    {
        foreach (var cf in ws.ConditionalFormats)
        {
            var styleValue = ((XLStyle)cf.Style).Value;
            if (!styleValue.Equals(DefaultStyleValue) && !context.DifferentialFormats.ContainsKey(styleValue))
                AddConditionalDifferentialFormat(differentialFormats, cf, context);
        }
    }

    private static void AddTableFieldDxfs(DifferentialFormats differentialFormats, XLWorksheet ws,
        SaveContext context)
    {
        foreach (var tf in ws.Tables.SelectMany<XLTable, IXLTableField>(t => t.Fields))
        {
            if (!tf.IsConsistentStyle())
                continue;

            var style = ((XLStyle)tf.Column.Cells()
                .Skip(tf.Table.ShowHeaderRow ? 1 : 0)
                .First()
                .Style).Value;

            if (!style.Equals(DefaultStyleValue) && !context.DifferentialFormats.ContainsKey(style))
                AddStyleAsDifferentialFormat(differentialFormats, style, context);
        }
    }

    private static void AddPivotTableDxfs(DifferentialFormats differentialFormats, XLWorksheet ws,
        SaveContext context)
    {
        foreach (var pt in ws.PivotTables)
        {
            AddPivotTableStyleFormatDxfs(differentialFormats, pt, context);
            AddPivotTableFormatDxfs(differentialFormats, pt, context);
            AddPivotTableConditionalFormatDxfs(differentialFormats, pt, context);
        }
    }

    private static void AddPivotTableStyleFormatDxfs(DifferentialFormats differentialFormats,
        XLPivotTable pt, SaveContext context)
    {
        foreach (var styleFormat in pt.AllStyleFormats)
        {
            var xlStyle = (XLStyle)styleFormat.Style;
            if (!xlStyle.Value.Equals(DefaultStyleValue) &&
                !context.DifferentialFormats.ContainsKey(xlStyle.Value))
                AddStyleAsDifferentialFormat(differentialFormats, xlStyle.Value, context);
        }
    }

    private static void AddPivotTableFormatDxfs(DifferentialFormats differentialFormats,
        XLPivotTable pt, SaveContext context)
    {
        foreach (var xlPivotFormat in pt.Formats)
        {
            var xlStyleValue = xlPivotFormat.DxfStyleValue;
            if (!xlStyleValue.Equals(XLStyleValue.Default) &&
                !context.DifferentialFormats.ContainsKey(xlStyleValue))
                AddStyleAsDifferentialFormat(differentialFormats, xlStyleValue, context);
        }
    }

    private static void AddPivotTableConditionalFormatDxfs(DifferentialFormats differentialFormats,
        XLPivotTable pt, SaveContext context)
    {
        foreach (var xlConditionalStyle in pt.ConditionalFormats)
        {
            var xlStyle = (XLStyle)xlConditionalStyle.Format.Style;
            if (!xlStyle.Value.Equals(XLStyleValue.Default) &&
                !context.DifferentialFormats.ContainsKey(xlStyle.Value))
                AddStyleAsDifferentialFormat(differentialFormats, xlStyle.Value, context);
        }
    }

    private static void AddAutoFilterColorFilterDxfs(DifferentialFormats differentialFormats, XLWorksheet ws,
        SaveContext context)
    {
        void AddColorFilterDxf(XLAutoFilter autoFilter)
        {
            foreach (var (_, xlFilterColumn) in autoFilter.Columns)
            {
                if (xlFilterColumn.FilterType != XLFilterType.Color)
                    continue;

                var key = (xlFilterColumn.FilterColor.Key, xlFilterColumn.FilterByCellColor);
                if (context.ColorFilterDxfIds.ContainsKey(key))
                    continue;

                var differentialFormat = new DifferentialFormat();
                if (xlFilterColumn.FilterByCellColor)
                {
                    var fillKey = new XLFillKey
                    {
                        PatternType = XLFillPatternValues.Solid,
                        BackgroundColor = xlFilterColumn.FilterColor.Key,
                        PatternColor = XLColor.FromIndex(64).Key,
                    };
                    var fillValue = XLFillValue.FromKey(ref fillKey);
                    var fill = GetNewFill(new FillInfo { Fill = fillValue }, differentialFillFormat: true);
                    differentialFormat.Append(fill);
                }
                else
                {
                    var fontKey = XLFontValue.Default.Key with
                    {
                        FontColor = xlFilterColumn.FilterColor.Key,
                    };
                    var fontValue = XLFontValue.FromKey(ref fontKey);
                    var font = GetNewFont(new FontInfo { Font = fontValue }, false);
                    if (font?.HasChildren ?? false)
                        differentialFormat.Append(font);
                }

                differentialFormats.Append(differentialFormat);
                context.ColorFilterDxfIds.Add(key, differentialFormats.Count() - 1);
            }
        }

        AddColorFilterDxf(ws.AutoFilter);

        foreach (var table in ws.Tables.Cast<XLTable>())
        {
            AddColorFilterDxf(table.AutoFilter);
        }
    }

    private static void FillDifferentialFormatsCollection(DifferentialFormats differentialFormats,
        Dictionary<XLStyleValue, int> dictionary)
    {
        dictionary.Clear();
        var id = 0;

        foreach (var df in differentialFormats.Elements<DifferentialFormat>())
        {
            var emptyContainer = new XLStylizedEmpty(DefaultStyle);
            OpenXmlHelper.LoadFont(df.Font, emptyContainer.Style.Font);
            OpenXmlHelper.LoadBorder(df.Border, emptyContainer.Style.Border);
            OpenXmlHelper.LoadNumberFormat(df.NumberingFormat, emptyContainer.Style.NumberFormat);
            OpenXmlHelper.LoadFill(df.Fill, emptyContainer.Style.Fill, differentialFillFormat: true);

            if (!dictionary.ContainsKey(emptyContainer.StyleValue))
                dictionary.Add(emptyContainer.StyleValue, id++);
        }
    }

    private static void AddConditionalDifferentialFormat(DifferentialFormats differentialFormats,
        IXLConditionalFormat cf,
        SaveContext context)
    {
        var differentialFormat = new DifferentialFormat();
        var styleValue = ((XLStyle)cf.Style).Value;

        var diffFont = GetNewFont(new FontInfo { Font = styleValue.Font }, false);
        if (diffFont?.HasChildren ?? false)
            differentialFormat.Append(diffFont);

        if (!string.IsNullOrWhiteSpace(cf.Style.NumberFormat.Format))
        {
            var numberFormat = new NumberingFormat
            {
                NumberFormatId = (uint)(XLConstants.NumberOfBuiltInStyles + differentialFormats.Count()),
                FormatCode = cf.Style.NumberFormat.Format
            };
            differentialFormat.Append(numberFormat);
        }

        var diffFill = GetNewFill(new FillInfo { Fill = styleValue.Fill }, differentialFillFormat: true);
        if (diffFill?.HasChildren ?? false)
            differentialFormat.Append(diffFill);

        var diffBorder = GetNewBorder(new BorderInfo { Border = styleValue.Border }, false);
        if (diffBorder?.HasChildren ?? false)
            differentialFormat.Append(diffBorder);

        differentialFormats.Append(differentialFormat);

        context.DifferentialFormats.Add(styleValue, differentialFormats.Count() - 1);
    }

    private static void AddStyleAsDifferentialFormat(DifferentialFormats differentialFormats, XLStyleValue style,
        SaveContext context)
    {
        var differentialFormat = new DifferentialFormat();

        var diffFont = GetNewFont(new FontInfo { Font = style.Font }, false);
        if (diffFont?.HasChildren ?? false)
            differentialFormat.Append(diffFont);

        if (!string.IsNullOrWhiteSpace(style.NumberFormat.Format) || style.NumberFormat.NumberFormatId != 0)
        {
            var numberFormat = new NumberingFormat();

            if (style.NumberFormat.NumberFormatId == -1)
            {
                numberFormat.FormatCode = style.NumberFormat.Format;
                numberFormat.NumberFormatId = (uint)(XLConstants.NumberOfBuiltInStyles +
                                                     differentialFormats
                                                         .Descendants<DifferentialFormat>()
                                                         .Count(df =>
                                                             df.NumberingFormat != null &&
                                                             df.NumberingFormat.NumberFormatId != null &&
                                                             df.NumberingFormat.NumberFormatId.Value >=
                                                             XLConstants.NumberOfBuiltInStyles));
            }
            else
            {
                numberFormat.NumberFormatId = (uint)(style.NumberFormat.NumberFormatId);
                if (!string.IsNullOrEmpty(style.NumberFormat.Format))
                    numberFormat.FormatCode = style.NumberFormat.Format;
                else if (XLPredefinedFormat.FormatCodes.TryGetValue(style.NumberFormat.NumberFormatId,
                             out var formatCode))
                    numberFormat.FormatCode = formatCode;
            }

            differentialFormat.Append(numberFormat);
        }

        var diffFill = GetNewFill(new FillInfo { Fill = style.Fill }, differentialFillFormat: true);
        if (diffFill?.HasChildren ?? false)
            differentialFormat.Append(diffFill);

        var diffBorder = GetNewBorder(new BorderInfo { Border = style.Border }, false);
        if (diffBorder?.HasChildren ?? false)
            differentialFormat.Append(diffBorder);

        differentialFormats.Append(differentialFormat);

        context.DifferentialFormats.Add(style, differentialFormats.Count() - 1);
    }

    private static void ResolveRest(WorkbookStylesPart workbookStylesPart, SaveContext context)
    {
        workbookStylesPart.Stylesheet!.CellFormats ??= new CellFormats();

        foreach (var styleInfo in context.SharedStyles.Values)
        {
            var info = styleInfo;
            var foundOne =
                workbookStylesPart.Stylesheet!.CellFormats.Cast<CellFormat>()
                    .Any(f => CellFormatsAreEqual(f, info, compareAlignment: true));

            if (foundOne) continue;

            var cellFormat = GetCellFormat(styleInfo);
            cellFormat.FormatId = 0;
            var alignment = new Alignment
            {
                Horizontal = styleInfo.Style.Alignment.Horizontal.ToOpenXml(),
                Vertical = styleInfo.Style.Alignment.Vertical.ToOpenXml(),
                Indent = (uint)styleInfo.Style.Alignment.Indent,
                ReadingOrder = (uint)styleInfo.Style.Alignment.ReadingOrder,
                WrapText = styleInfo.Style.Alignment.WrapText,
                TextRotation = (uint)GetOpenXmlTextRotation(styleInfo.Style.Alignment),
                ShrinkToFit = styleInfo.Style.Alignment.ShrinkToFit,
                RelativeIndent = styleInfo.Style.Alignment.RelativeIndent,
                JustifyLastLine = styleInfo.Style.Alignment.JustifyLastLine
            };
            cellFormat.AppendChild(alignment);

            if (cellFormat.ApplyProtection!.Value)
                cellFormat.AppendChild(GetProtection(styleInfo));

            workbookStylesPart.Stylesheet!.CellFormats!.AppendChild(cellFormat);
        }

        workbookStylesPart.Stylesheet!.CellFormats!.Count = (uint)workbookStylesPart.Stylesheet!.CellFormats!.Count();

        static int GetOpenXmlTextRotation(XLAlignmentValue alignment)
        {
            var textRotation = alignment.TextRotation;
            return textRotation >= 0
                ? textRotation
                : 90 - textRotation;
        }
    }

    private static void ResolveCellStyleFormats(WorkbookStylesPart workbookStylesPart,
        SaveContext context)
    {
        workbookStylesPart.Stylesheet!.CellStyleFormats ??= new CellStyleFormats();

        foreach (var styleInfo in context.SharedStyles.Values)
        {
            var info = styleInfo;
            var foundOne =
                workbookStylesPart.Stylesheet!.CellStyleFormats.Cast<CellFormat>()
                    .Any(f => CellFormatsAreEqual(f, info, compareAlignment: false));

            if (foundOne) continue;

            var cellStyleFormat = GetCellFormat(styleInfo);

            if (cellStyleFormat.ApplyProtection!.Value)
                cellStyleFormat.AppendChild(GetProtection(styleInfo));

            workbookStylesPart.Stylesheet!.CellStyleFormats!.AppendChild(cellStyleFormat);
        }

        workbookStylesPart.Stylesheet!.CellStyleFormats!.Count =
            (uint)workbookStylesPart.Stylesheet!.CellStyleFormats!.Count();
    }

    private static bool ApplyFill(StyleInfo styleInfo)
    {
        return styleInfo.Style.Fill.PatternType.ToOpenXml() == PatternValues.None;
    }

    private static bool ApplyBorder(StyleInfo styleInfo)
    {
        var opBorder = styleInfo.Style.Border;
        return (opBorder.BottomBorder.ToOpenXml() != BorderStyleValues.None
                || opBorder.DiagonalBorder.ToOpenXml() != BorderStyleValues.None
                || opBorder.RightBorder.ToOpenXml() != BorderStyleValues.None
                || opBorder.LeftBorder.ToOpenXml() != BorderStyleValues.None
                || opBorder.TopBorder.ToOpenXml() != BorderStyleValues.None);
    }

    private static bool ApplyProtection(StyleInfo styleInfo)
    {
        return styleInfo.Style.Protection != null;
    }

    private static CellFormat GetCellFormat(StyleInfo styleInfo)
    {
        var cellFormat = new CellFormat
        {
            NumberFormatId = (uint)styleInfo.NumberFormatId,
            FontId = styleInfo.FontId,
            FillId = styleInfo.FillId,
            BorderId = styleInfo.BorderId,
            QuotePrefix = OpenXmlHelper.GetBooleanValue(styleInfo.IncludeQuotePrefix, false),
            ApplyNumberFormat = true,
            ApplyAlignment = true,
            ApplyFill = ApplyFill(styleInfo),
            ApplyBorder = ApplyBorder(styleInfo),
            ApplyProtection = ApplyProtection(styleInfo)
        };
        return cellFormat;
    }

    private static Protection GetProtection(StyleInfo styleInfo)
    {
        return new Protection
        {
            Locked = styleInfo.Style.Protection.Locked,
            Hidden = styleInfo.Style.Protection.Hidden
        };
    }

    /// <summary>
    /// Check if two styles are equivalent.
    /// </summary>
    /// <param name="f">Style in the OpenXML format.</param>
    /// <param name="styleInfo">Style in the XLibur format.</param>
    /// <param name="compareAlignment">Flag specifying whether compare the alignments of two styles.
    /// Styles in the x:cellStyleXfs section do not include alignment, so we don't have to compare it in this case.
    /// Styles in the x:cellXfs section, on the opposite, do include alignments, and we must compare them.
    /// </param>
    /// <returns>True if two formats are equivalent, false otherwise.</returns>
    private static bool CellFormatsAreEqual(CellFormat f, StyleInfo styleInfo, bool compareAlignment)
    {
        return
            f.BorderId != null && styleInfo.BorderId == f.BorderId
                               && f.FillId != null && styleInfo.FillId == f.FillId
                               && f.FontId != null && styleInfo.FontId == f.FontId
                               && f.NumberFormatId != null && styleInfo.NumberFormatId == f.NumberFormatId
                               && QuotePrefixesAreEqual(f.QuotePrefix, styleInfo.IncludeQuotePrefix)
                               && (f.ApplyFill == null && styleInfo.Style.Fill == XLFillValue.Default ||
                                   f.ApplyFill != null && f.ApplyFill == ApplyFill(styleInfo))
                               && (f.ApplyBorder == null && styleInfo.Style.Border == XLBorderValue.Default ||
                                   f.ApplyBorder != null && f.ApplyBorder == ApplyBorder(styleInfo))
                               && (!compareAlignment || AlignmentsAreEqual(f.Alignment, styleInfo.Style.Alignment))
                               && ProtectionsAreEqual(f.Protection, styleInfo.Style.Protection)
            ;
    }

    private static bool ProtectionsAreEqual(Protection? protection, XLProtectionValue xlProtection)
    {
        var p = XLProtectionValue.Default.Key;
        if (protection is not null)
            p = OpenXmlHelper.ProtectionToXLibur(protection, p);

        return p.Equals(xlProtection.Key);
    }

    private static bool QuotePrefixesAreEqual(BooleanValue? quotePrefix, bool includeQuotePrefix)
    {
        return OpenXmlHelper.GetBooleanValueAsBool(quotePrefix, false) == includeQuotePrefix;
    }

    private static bool AlignmentsAreEqual(Alignment? alignment, XLAlignmentValue xlAlignment)
    {
        if (alignment is null) return XLStyle.Default.Value.Alignment.Equals(xlAlignment);
        var a = OpenXmlHelper.AlignmentToXLibur(alignment, XLAlignmentValue.Default.Key);
        return a.Equals(xlAlignment.Key);
    }

    private static Dictionary<XLBorderValue, BorderInfo> ResolveBorders(WorkbookStylesPart workbookStylesPart,
        Dictionary<XLBorderValue, BorderInfo> sharedBorders)
    {
        workbookStylesPart.Stylesheet!.Borders ??= new Borders();

        var allSharedBorders = new Dictionary<XLBorderValue, BorderInfo>();
        foreach (var borderInfo in sharedBorders.Values)
        {
            var borderId = 0;
            var foundOne = false;
            foreach (var openXmlElement in workbookStylesPart.Stylesheet!.Borders)
            {
                var f = (Border)openXmlElement;
                if (BordersAreEqual(f, borderInfo.Border))
                {
                    foundOne = true;
                    break;
                }

                borderId++;
            }

            if (!foundOne)
            {
                var border = GetNewBorder(borderInfo);
                workbookStylesPart.Stylesheet!.Borders.AppendChild(border);
            }

            allSharedBorders.Add(borderInfo.Border,
                borderInfo with { BorderId = (uint)borderId });
        }

        workbookStylesPart.Stylesheet!.Borders.Count = (uint)workbookStylesPart.Stylesheet!.Borders.Count();
        return allSharedBorders;
    }

    private static Border GetNewBorder(BorderInfo borderInfo, bool ignoreMod = true)
    {
        var border = new Border();
        if (borderInfo.Border.DiagonalUp != XLBorderValue.Default.DiagonalUp || ignoreMod)
            border.DiagonalUp = borderInfo.Border.DiagonalUp;

        if (borderInfo.Border.DiagonalDown != XLBorderValue.Default.DiagonalDown || ignoreMod)
            border.DiagonalDown = borderInfo.Border.DiagonalDown;

        if (borderInfo.Border.LeftBorder != XLBorderValue.Default.LeftBorder || ignoreMod)
        {
            var leftBorder = new LeftBorder { Style = borderInfo.Border.LeftBorder.ToOpenXml() };
            if (borderInfo.Border.LeftBorderColor != XLBorderValue.Default.LeftBorderColor || ignoreMod)
            {
                var leftBorderColor = new Color().FromXLiburColor<Color>(borderInfo.Border.LeftBorderColor);
                leftBorder.AppendChild(leftBorderColor);
            }

            border.AppendChild(leftBorder);
        }

        if (borderInfo.Border.RightBorder != XLBorderValue.Default.RightBorder || ignoreMod)
        {
            var rightBorder = new RightBorder { Style = borderInfo.Border.RightBorder.ToOpenXml() };
            if (borderInfo.Border.RightBorderColor != XLBorderValue.Default.RightBorderColor || ignoreMod)
            {
                var rightBorderColor = new Color().FromXLiburColor<Color>(borderInfo.Border.RightBorderColor);
                rightBorder.AppendChild(rightBorderColor);
            }

            border.AppendChild(rightBorder);
        }

        if (borderInfo.Border.TopBorder != XLBorderValue.Default.TopBorder || ignoreMod)
        {
            var topBorder = new TopBorder { Style = borderInfo.Border.TopBorder.ToOpenXml() };
            if (borderInfo.Border.TopBorderColor != XLBorderValue.Default.TopBorderColor || ignoreMod)
            {
                var topBorderColor = new Color().FromXLiburColor<Color>(borderInfo.Border.TopBorderColor);
                topBorder.AppendChild(topBorderColor);
            }

            border.AppendChild(topBorder);
        }

        if (borderInfo.Border.BottomBorder != XLBorderValue.Default.BottomBorder || ignoreMod)
        {
            var bottomBorder = new BottomBorder { Style = borderInfo.Border.BottomBorder.ToOpenXml() };
            if (borderInfo.Border.BottomBorderColor != XLBorderValue.Default.BottomBorderColor || ignoreMod)
            {
                var bottomBorderColor = new Color().FromXLiburColor<Color>(borderInfo.Border.BottomBorderColor);
                bottomBorder.AppendChild(bottomBorderColor);
            }

            border.AppendChild(bottomBorder);
        }

        if (borderInfo.Border.DiagonalBorder != XLBorderValue.Default.DiagonalBorder || ignoreMod)
        {
            var diagonalBorder = new DiagonalBorder { Style = borderInfo.Border.DiagonalBorder.ToOpenXml() };
            if (borderInfo.Border.DiagonalBorderColor != XLBorderValue.Default.DiagonalBorderColor || ignoreMod)
                if (borderInfo.Border.DiagonalBorderColor != null)
                {
                    var DiagonalBorderColor =
                        new Color().FromXLiburColor<Color>(borderInfo.Border.DiagonalBorderColor);
                    diagonalBorder.AppendChild(DiagonalBorderColor);
                }

            border.AppendChild(diagonalBorder);
        }

        return border;
    }

    private static bool BordersAreEqual(Border border, XLBorderValue xlBorder)
    {
        var convertedBorder = OpenXmlHelper.BorderToXLibur(
            border,
            XLBorderValue.Default.Key);
        return convertedBorder.Equals(xlBorder.Key);
    }

    private static Dictionary<XLFillValue, FillInfo> ResolveFills(WorkbookStylesPart workbookStylesPart,
        Dictionary<XLFillValue, FillInfo> sharedFills)
    {
        workbookStylesPart.Stylesheet!.Fills ??= new Fills();

        var fills = workbookStylesPart.Stylesheet!.Fills;

        // Pattern idx 0 and idx 1 are hardcoded to Excel with values None (0) and Gray125. Excel will ignore
        // values from the file. Every file has had these values inside to keep the first available idx at 2.
        ResolveFillWithPattern(fills, 0, PatternValues.None);
        ResolveFillWithPattern(fills, 1, PatternValues.Gray125);

        var allSharedFills = new Dictionary<XLFillValue, FillInfo>();
        foreach (var fillInfo in sharedFills.Values)
        {
            var fillId = 0;
            var foundOne = false;
            foreach (var openXmlElement in fills)
            {
                var f = (Fill)openXmlElement;
                if (FillsAreEqual(f, fillInfo.Fill, fromDifferentialFormat: false))
                {
                    foundOne = true;
                    break;
                }

                fillId++;
            }

            if (!foundOne)
            {
                var fill = GetNewFill(fillInfo, differentialFillFormat: false);
                fills.AppendChild(fill);
            }

            allSharedFills.Add(fillInfo.Fill, fillInfo with { FillId = (uint)fillId });
        }

        fills.Count = (uint)fills.Count();
        return allSharedFills;
    }

    private static void ResolveFillWithPattern(Fills fills, int index, PatternValues patternValues)
    {
        var fill = (Fill?)fills.ElementAtOrDefault(index);
        if (fill is null)
        {
            fills.InsertAt(new Fill { PatternFill = new PatternFill { PatternType = patternValues } }, index);
            return;
        }

        var fillHasExpectedValue =
            fill.PatternFill?.PatternType?.Value == patternValues &&
            fill.PatternFill.ForegroundColor is null &&
            fill.PatternFill.BackgroundColor is null;

        if (fillHasExpectedValue)
            return;

        fill.PatternFill = new PatternFill { PatternType = patternValues };
    }

    private static Fill GetNewFill(FillInfo fillInfo, bool differentialFillFormat)
    {
        var fill = new Fill();

        var patternFill = new PatternFill
        {
            PatternType = fillInfo.Fill.PatternType.ToOpenXml()
        };

        BackgroundColor backgroundColor;
        ForegroundColor foregroundColor;

        switch (fillInfo.Fill.PatternType)
        {
            case XLFillPatternValues.None:
                break;

            case XLFillPatternValues.Solid:

                if (differentialFillFormat)
                {
                    patternFill.AppendChild(new ForegroundColor { Auto = true });
                    backgroundColor =
                        new BackgroundColor().FromXLiburColor<BackgroundColor>(fillInfo.Fill.BackgroundColor, true);
                    if (backgroundColor.HasAttributes)
                        patternFill.AppendChild(backgroundColor);
                }
                else
                {
                    // XLibur Background color to be populated into OpenXML fgColor
                    foregroundColor =
                        new ForegroundColor().FromXLiburColor<ForegroundColor>(fillInfo.Fill.BackgroundColor);
                    if (foregroundColor.HasAttributes)
                        patternFill.AppendChild(foregroundColor);
                }

                break;

            case XLFillPatternValues.DarkDown:
            case XLFillPatternValues.DarkGray:
            case XLFillPatternValues.DarkGrid:
            case XLFillPatternValues.DarkHorizontal:
            case XLFillPatternValues.DarkTrellis:
            case XLFillPatternValues.DarkUp:
            case XLFillPatternValues.DarkVertical:
            case XLFillPatternValues.Gray0625:
            case XLFillPatternValues.Gray125:
            case XLFillPatternValues.LightDown:
            case XLFillPatternValues.LightGray:
            case XLFillPatternValues.LightGrid:
            case XLFillPatternValues.LightHorizontal:
            case XLFillPatternValues.LightTrellis:
            case XLFillPatternValues.LightUp:
            case XLFillPatternValues.LightVertical:
            case XLFillPatternValues.MediumGray:
            default:

                foregroundColor = new ForegroundColor().FromXLiburColor<ForegroundColor>(fillInfo.Fill.PatternColor);
                if (foregroundColor.HasAttributes)
                    patternFill.AppendChild(foregroundColor);

                backgroundColor =
                    new BackgroundColor().FromXLiburColor<BackgroundColor>(fillInfo.Fill.BackgroundColor);
                if (backgroundColor.HasAttributes)
                    patternFill.AppendChild(backgroundColor);

                break;
        }

        if (patternFill.HasChildren)
            fill.AppendChild(patternFill);

        return fill;
    }

    private static bool FillsAreEqual(Fill f, XLFillValue xlFill, bool fromDifferentialFormat)
    {
        var nF = new XLFill();

        OpenXmlHelper.LoadFill(f, nF, fromDifferentialFormat);

        return nF.Key.Equals(xlFill.Key);
    }

    private static void ResolveFonts(WorkbookStylesPart workbookStylesPart, SaveContext context)
    {
        workbookStylesPart.Stylesheet!.Fonts ??= new Fonts();

        var newFonts = new Dictionary<XLFontValue, FontInfo>();
        foreach (var fontInfo in context.SharedFonts.Values)
        {
            var fontId = 0;
            var foundOne = false;
            foreach (var openXmlElement in workbookStylesPart.Stylesheet!.Fonts)
            {
                var f = (Font)openXmlElement;
                if (FontsAreEqual(f, fontInfo.Font))
                {
                    foundOne = true;
                    break;
                }

                fontId++;
            }

            if (!foundOne)
            {
                var font = GetNewFont(fontInfo);
                workbookStylesPart.Stylesheet!.Fonts.AppendChild(font);
            }

            newFonts.Add(fontInfo.Font, new FontInfo { Font = fontInfo.Font, FontId = (uint)fontId });
        }

        context.SharedFonts.Clear();
        foreach (var kp in newFonts)
            context.SharedFonts.Add(kp.Key, kp.Value);

        workbookStylesPart.Stylesheet!.Fonts.Count = (uint)workbookStylesPart.Stylesheet!.Fonts.Count();
    }

    private static Font GetNewFont(FontInfo fontInfo, bool ignoreMod = true)
    {
        var font = new Font();
        var bold = (fontInfo.Font.Bold != XLFontValue.Default.Bold || ignoreMod) && fontInfo.Font.Bold
            ? new Bold()
            : null;
        var italic = (fontInfo.Font.Italic != XLFontValue.Default.Italic || ignoreMod) && fontInfo.Font.Italic
            ? new Italic()
            : null;
        var underline = (fontInfo.Font.Underline != XLFontValue.Default.Underline || ignoreMod) &&
                        fontInfo.Font.Underline != XLFontUnderlineValues.None
            ? new Underline { Val = fontInfo.Font.Underline.ToOpenXml() }
            : null;
        var strike = (fontInfo.Font.Strikethrough != XLFontValue.Default.Strikethrough || ignoreMod) &&
                     fontInfo.Font.Strikethrough
            ? new Strike()
            : null;
        var verticalAlignment = fontInfo.Font.VerticalAlignment != XLFontValue.Default.VerticalAlignment || ignoreMod
            ? new VerticalTextAlignment { Val = fontInfo.Font.VerticalAlignment.ToOpenXml() }
            : null;

        var shadow = (fontInfo.Font.Shadow != XLFontValue.Default.Shadow || ignoreMod) && fontInfo.Font.Shadow
            ? new Shadow()
            : null;
        var fontSize = fontInfo.Font.FontSize != XLFontValue.Default.FontSize || ignoreMod
            ? new FontSize { Val = fontInfo.Font.FontSize }
            : null;
        var color = fontInfo.Font.FontColor != XLFontValue.Default.FontColor || ignoreMod
            ? new Color().FromXLiburColor<Color>(fontInfo.Font.FontColor)
            : null;

        var fontName = fontInfo.Font.FontName != XLFontValue.Default.FontName || ignoreMod
            ? new FontName { Val = fontInfo.Font.FontName }
            : null;
        var fontFamilyNumbering =
            fontInfo.Font.FontFamilyNumbering != XLFontValue.Default.FontFamilyNumbering || ignoreMod
                ? new FontFamilyNumbering { Val = (int)fontInfo.Font.FontFamilyNumbering }
                : null;

        var fontCharSet = (fontInfo.Font.FontCharSet != XLFontValue.Default.FontCharSet || ignoreMod) &&
                          fontInfo.Font.FontCharSet != XLFontCharSet.Default
            ? new FontCharSet { Val = (int)fontInfo.Font.FontCharSet }
            : null;

        var fontScheme = (fontInfo.Font.FontScheme != XLFontValue.Default.FontScheme || ignoreMod) &&
                         fontInfo.Font.FontScheme != XLFontScheme.None
            ? new FontScheme { Val = fontInfo.Font.FontScheme.ToOpenXmlEnum() }
            : null;

        if (bold != null)
            font.AppendChild(bold);
        if (italic != null)
            font.AppendChild(italic);
        if (underline != null)
            font.AppendChild(underline);
        if (strike != null)
            font.AppendChild(strike);
        if (verticalAlignment != null)
            font.AppendChild(verticalAlignment);
        if (shadow != null)
            font.AppendChild(shadow);
        if (fontSize != null)
            font.AppendChild(fontSize);
        if (color != null)
            font.AppendChild(color);
        if (fontName != null)
            font.AppendChild(fontName);
        if (fontFamilyNumbering != null)
            font.AppendChild(fontFamilyNumbering);
        if (fontCharSet != null)
            font.AppendChild(fontCharSet);
        if (fontScheme != null)
            font.AppendChild(fontScheme);

        return font;
    }

    private static bool FontsAreEqual(Font font, XLFontValue xlFont)
    {
        var convertedFont = OpenXmlHelper.FontToXLibur(
            font,
            XLFontValue.Default.Key);
        return convertedFont.Equals(xlFont.Key);
    }

    private static Dictionary<XLNumberFormatValue, NumberFormatInfo> ResolveNumberFormats(
        WorkbookStylesPart workbookStylesPart,
        HashSet<XLNumberFormatValue> customNumberFormats,
        uint defaultFormatId)
    {
        if (workbookStylesPart.Stylesheet!.NumberingFormats == null)
        {
            workbookStylesPart.Stylesheet!.NumberingFormats = new NumberingFormats();
            workbookStylesPart.Stylesheet!.NumberingFormats.AppendChild(new NumberingFormat()
            {
                NumberFormatId = 0,
                FormatCode = ""
            });
        }

        var allSharedNumberFormats = new Dictionary<XLNumberFormatValue, NumberFormatInfo>();
        var partNumberingFormats = workbookStylesPart.Stylesheet!.NumberingFormats;

        // number format ids in the part can have holes in the sequence, and the first id can be greater than the last built-in style id.
        // In some cases, there are also existing number formats with id below the last built-in style id.
        var availableNumberFormatId = partNumberingFormats.Any()
            ? Math.Max(partNumberingFormats.Cast<NumberingFormat>().Max(nf => nf.NumberFormatId!.Value) + 1,
                XLConstants.NumberOfBuiltInStyles)
            : XLConstants.NumberOfBuiltInStyles; // 0-based

        // Merge custom formats used in the workbook that are not already present in the part to the part and assign ids
        foreach (var customNumberFormat in customNumberFormats.Where(nf => nf.NumberFormatId != defaultFormatId))
        {
            NumberingFormat? partNumberFormat = null;
            foreach (var nf in workbookStylesPart.Stylesheet!.NumberingFormats.Cast<NumberingFormat>())
            {
                if (!CustomNumberFormatsAreEqual(nf, customNumberFormat)) continue;
                partNumberFormat = nf;
                break;
            }

            if (partNumberFormat is null)
            {
                partNumberFormat = new NumberingFormat
                {
                    NumberFormatId = availableNumberFormatId++,
                    FormatCode = customNumberFormat.Format
                };
                workbookStylesPart.Stylesheet!.NumberingFormats.AppendChild(partNumberFormat);
            }

            allSharedNumberFormats.Add(customNumberFormat,
                new NumberFormatInfo
                {
                    NumberFormat = customNumberFormat,
                    NumberFormatId = (int)partNumberFormat.NumberFormatId!.Value
                });
        }

        workbookStylesPart.Stylesheet!.NumberingFormats.Count =
            (uint)workbookStylesPart.Stylesheet!.NumberingFormats.Count();
        return allSharedNumberFormats;
    }

    private static bool CustomNumberFormatsAreEqual(NumberingFormat nf, XLNumberFormatValue xlNumberFormat)
    {
        if (nf.FormatCode != null && !string.IsNullOrWhiteSpace(nf.FormatCode.Value))
            return string.Equals(xlNumberFormat.Format, nf.FormatCode.Value);

        return false;
    }
}
