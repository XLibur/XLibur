using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.AutoFilters;
using XLibur.Excel.Coordinates;
using XLibur.Excel.Tables;
using XLibur.Utils;
using static XLibur.Excel.XLWorkbook;
using EnumerableExtensions = XLibur.Extensions.EnumerableExtensions;

namespace XLibur.Excel.IO;

internal static class WorkbookStylesPartWriter
{
    internal static void GenerateContent(Stylesheet stylesheet, XLWorkbook workbook,
        SaveContext context)
    {
        var defaultStyle = DefaultStyleValue;

        if (!context.SharedFonts.ContainsKey(defaultStyle.Font))
            context.SharedFonts.Add(defaultStyle.Font, new FontInfo { FontId = 0, Font = defaultStyle.Font });

        stylesheet.CellStyles ??= new CellStyles();

        var defaultFormatId = ResolveDefaultFormatId(stylesheet);

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

        foreach (var font in xlStyles.Select(s => s.Font).Distinct()
                     .Where(f => !context.SharedFonts.ContainsKey(f)))
        {
            context.SharedFonts.Add(font, new FontInfo { FontId = fontCount++, Font = font });
        }

        var sharedFills = xlStyles.Select(s => s.Fill).Distinct().ToDictionary(
            f => f, f => new FillInfo { FillId = fillCount++, Fill = f });

        var sharedBorders = xlStyles.Select(s => s.Border).Distinct().ToDictionary(
            b => b, b => new BorderInfo { BorderId = borderCount++, Border = b });

        var customNumberFormats = CollectCustomNumberFormats(xlStyles, pivotCustomFormats);

        var allSharedNumberFormats = ResolveNumberFormats(stylesheet, customNumberFormats, defaultFormatId);
        foreach (var nf in allSharedNumberFormats)
            context.SharedNumberFormats.Add(nf.Key, nf.Value);

        ResolveFonts(stylesheet, context.SharedFonts);
        var allSharedFills = ResolveFills(stylesheet, sharedFills);
        var allSharedBorders = ResolveBorders(stylesheet, sharedBorders);

        BuildSharedStyleMappings(context, xlStyles, allSharedNumberFormats, allSharedFills, allSharedBorders);

        ResolveCellStyleFormats(stylesheet, context);
        ResolveRest(stylesheet, context);

        if (!stylesheet.CellStyles.Elements<CellStyle>().Any(c =>
                c.BuiltinId != null && c.BuiltinId.HasValue && c.BuiltinId.Value == 0U))
            stylesheet.CellStyles.AppendChild(new CellStyle
            { Name = "Normal", FormatId = defaultFormatId, BuiltinId = 0U });

        stylesheet.CellStyles.Count = (uint)stylesheet.CellStyles.ChildElements.Count;

        RemapStyleIds(stylesheet, context);

        AddDifferentialFormats(stylesheet, workbook, context);
    }

    /// <summary>
    /// Styles part for the forward-only streaming writer.
    /// </summary>
    /// <remarks>
    /// The normal path collects the workbook's styles into a set, writes deduplicated
    /// <c>cellXf</c>s and then reads the final ids back out. The streaming writer cannot work
    /// that way: it hands a style id to a cell the moment that cell is written, long before the
    /// styles part exists, and that sheet XML is already in the package by the time this runs.
    /// So the ids are the input, not the output - one <c>cellXf</c> per style in exactly the
    /// order the ids were handed out, no deduplication and no remap, which makes index i the
    /// style at <paramref name="orderedStyles"/>[i] by construction.
    /// <paramref name="orderedStyles"/>[0] must be the default style.
    /// <para>
    /// Takes no <c>SaveContext</c>: it needs three of that type's ten members and the caller
    /// discards all of them, so the bag was a signature to satisfy rather than state to carry.
    /// </para>
    /// </remarks>
    internal static void GenerateStreamingContent(Stylesheet stylesheet,
        IReadOnlyList<XLStyleValue> orderedStyles)
    {
        stylesheet.CellStyles ??= new CellStyles();

        var sharedFonts = new Dictionary<XLFontValue, FontInfo>();
        var sharedNumberFormats = new Dictionary<XLNumberFormatValue, NumberFormatInfo>();
        var sharedStyles = new Dictionary<XLStyleValue, StyleInfo>();

        var defaultStyle = DefaultStyleValue;
        if (!sharedFonts.ContainsKey(defaultStyle.Font))
            sharedFonts.Add(defaultStyle.Font, new FontInfo { FontId = 0, Font = defaultStyle.Font });

        uint fontCount = 1;
        uint fillCount = 3;
        uint borderCount = 1;

        foreach (var style in orderedStyles)
        {
            if (!sharedFonts.ContainsKey(style.Font))
                sharedFonts.Add(style.Font, new FontInfo { FontId = fontCount++, Font = style.Font });
        }

        var sharedFills = new Dictionary<XLFillValue, FillInfo>();
        var sharedBorders = new Dictionary<XLBorderValue, BorderInfo>();
        var customNumberFormats = new HashSet<XLNumberFormatValue>();
        foreach (var style in orderedStyles)
        {
            if (!sharedFills.ContainsKey(style.Fill))
                sharedFills.Add(style.Fill, new FillInfo { FillId = fillCount++, Fill = style.Fill });

            if (!sharedBorders.ContainsKey(style.Border))
                sharedBorders.Add(style.Border, new BorderInfo { BorderId = borderCount++, Border = style.Border });

            if (style.NumberFormat.NumberFormatId == -1)
                customNumberFormats.Add(style.NumberFormat);
        }

        var allSharedNumberFormats = ResolveNumberFormats(stylesheet, customNumberFormats, 0);
        foreach (var nf in allSharedNumberFormats)
            sharedNumberFormats.Add(nf.Key, nf.Value);

        ResolveFonts(stylesheet, sharedFonts);
        var allSharedFills = ResolveFills(stylesheet, sharedFills);
        var allSharedBorders = ResolveBorders(stylesheet, sharedBorders);

        // A single cellStyleXf holding the default format; every cellXf points at it via FormatId 0.
        stylesheet.CellStyleFormats ??= new CellStyleFormats();
        stylesheet.CellStyleFormats.AppendChild(new CellFormat
        {
            NumberFormatId = 0,
            FontId = 0,
            FillId = 0,
            BorderId = 0
        });
        stylesheet.CellStyleFormats.Count = (uint)stylesheet.CellStyleFormats.ChildElements.Count;

        stylesheet.CellFormats ??= new CellFormats();
        for (var styleId = 0; styleId < orderedStyles.Count; styleId++)
        {
            var style = orderedStyles[styleId];
            var numberFormatId = style.NumberFormat.NumberFormatId >= 0
                ? style.NumberFormat.NumberFormatId
                : allSharedNumberFormats[style.NumberFormat].NumberFormatId;

            var styleInfo = new StyleInfo
            {
                StyleId = (uint)styleId,
                Style = style,
                FontId = sharedFonts[style.Font].FontId,
                FillId = allSharedFills[style.Fill].FillId,
                BorderId = allSharedBorders[style.Border].BorderId,
                NumberFormatId = numberFormatId,
                IncludeQuotePrefix = style.IncludeQuotePrefix
            };

            sharedStyles[style] = styleInfo;
            stylesheet.CellFormats.AppendChild(BuildCellFormat(styleInfo));
        }

        stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.ChildElements.Count;

        stylesheet.CellStyles.AppendChild(new CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U });
        stylesheet.CellStyles.Count = (uint)stylesheet.CellStyles.ChildElements.Count;
    }

    /// <summary>
    /// Determine the default workbook style by looking for the style with builtInId = 0.
    /// </summary>
    private static uint ResolveDefaultFormatId(Stylesheet stylesheet)
    {
        var cellStyles = stylesheet.CellStyles!;

        if (cellStyles.Elements<CellStyle>()
            .Any(c => c.BuiltinId != null && c.BuiltinId.HasValue && c.BuiltinId.Value == 0))
        {
            // Possible to have duplicate default cell styles - occurs when file gets saved under different cultures.
            // We prefer the style named Normal
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

            // Read the effective style straight from the slices. Going through GetCells() would
            // materialise an XLCell wrapper for every used cell just to read one property, which
            // on a large sheet is the single biggest allocation of the whole save.
            var cellsCollection = worksheet.Internals.CellsCollection;
            var cells = new XLCellsCollection.SlicesEnumerator(Area.Full, cellsCollection);
            while (cells.MoveNext())
                styles.Add(worksheet.GetStyleValue(cells.Current));

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
    private static void RemapStyleIds(Stylesheet stylesheet, SaveContext context)
    {
        var newSharedStyles = new Dictionary<XLStyleValue, StyleInfo>();
        foreach (var ss in context.SharedStyles)
        {
            var styleId = -1;
            foreach (var openXmlElement in stylesheet.CellFormats!)
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
    /// Rebuilds <c>&lt;dxfs&gt;</c> from the live workbook, and fills
    /// <see cref="SaveContext.DifferentialFormats"/> with the style-to-index map that every
    /// <c>dxfId</c> reference is then written from.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The collection is rebuilt wholesale rather than appended to, so a dxf that nothing
    /// references any longer does not survive the save. That is safe because <b>every</b>
    /// <c>dxfId</c> XLibur emits is re-derived here from the object model in the same pass —
    /// conditional-format rules and pivot formats through
    /// <see cref="SaveContext.DifferentialFormats"/>, table fields through the same map, and
    /// auto-filter colour filters through <see cref="SaveContext.ColorFilterDxfIds"/>. No
    /// reference is carried over from the loaded file, so no index has to stay stable across a
    /// save.
    /// </para>
    /// <para>
    /// Until this was fixed the method opened by clearing the collection and then calling a
    /// <c>FillDifferentialFormatsCollection</c> helper that iterated it to build a reuse map from
    /// "the differential formats currently in the file". It iterated the collection emptied on the
    /// line above, so the map was always empty and the reuse never happened — the rebuild below is
    /// what has always actually run. The helper is gone rather than repaired: arming it would mean
    /// keeping existing dxfs and their indices, which is a different save contract, not a bug fix.
    /// </para>
    /// </remarks>
    private static void AddDifferentialFormats(Stylesheet stylesheet, XLWorkbook workbook,
        SaveContext context)
    {
        stylesheet.DifferentialFormats ??= new DifferentialFormats();

        var differentialFormats = stylesheet.DifferentialFormats;
        differentialFormats.RemoveAllChildren();

        foreach (var ws in workbook.WorksheetsInternal)
        {
            AddConditionalFormatDxfs(differentialFormats, ws, context);
            AddTableFieldDxfs(differentialFormats, ws, context);
            AddPivotTableDxfs(differentialFormats, ws, context);
            AddAutoFilterColorFilterDxfs(differentialFormats, ws, context);
        }

        differentialFormats.Count = (uint)differentialFormats.ChildElements.Count;
        if (differentialFormats.Count == 0)
            stylesheet.DifferentialFormats = null;
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
        foreach (var xlStyleValue in pt.Formats
                     .Select(f => f.DxfStyleValue)
                     .Where(s => !s.Equals(XLStyleValue.Default) &&
                                 !context.DifferentialFormats.ContainsKey(s)))
        {
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
        AddColorFilterDxfs(differentialFormats, ws.AutoFilter, context);

        foreach (var table in ws.Tables.Cast<XLTable>())
            AddColorFilterDxfs(differentialFormats, table.AutoFilter, context);
    }

    private static void AddColorFilterDxfs(DifferentialFormats differentialFormats, XLAutoFilter autoFilter,
        SaveContext context)
    {
        foreach (var (_, xlFilterColumn) in autoFilter.Columns)
        {
            if (xlFilterColumn.FilterType != XLFilterType.Color)
                continue;

            var key = (xlFilterColumn.FilterColor.Key, xlFilterColumn.FilterByCellColor);
            if (context.ColorFilterDxfIds.ContainsKey(key))
                continue;

            var differentialFormat = CreateColorDifferentialFormat(xlFilterColumn);
            differentialFormats.Append(differentialFormat);
            context.ColorFilterDxfIds.Add(key, differentialFormats.ChildElements.Count - 1);
        }
    }

    private static DifferentialFormat CreateColorDifferentialFormat(XLFilterColumn xlFilterColumn)
    {
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

        return differentialFormat;
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
                NumberFormatId = (uint)(XLConstants.NumberOfBuiltInStyles + differentialFormats.ChildElements.Count),
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

        context.DifferentialFormats.Add(styleValue, differentialFormats.ChildElements.Count - 1);
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

        var diffAlignment = GetNewDifferentialAlignment(style.Alignment);
        if (diffAlignment is not null)
            differentialFormat.Append(diffAlignment);

        differentialFormats.Append(differentialFormat);

        context.DifferentialFormats.Add(style, differentialFormats.ChildElements.Count - 1);
    }

    private static void ResolveRest(Stylesheet stylesheet, SaveContext context)
    {
        stylesheet.CellFormats ??= new CellFormats();

        foreach (var styleInfo in context.SharedStyles.Values)
        {
            var info = styleInfo;
            var foundOne =
                stylesheet.CellFormats.Cast<CellFormat>()
                    .Any(f => CellFormatsAreEqual(f, info, compareAlignment: true));

            if (foundOne) continue;

            stylesheet.CellFormats.AppendChild(BuildCellFormat(styleInfo));
        }

        stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.ChildElements.Count;
    }

    /// <summary>
    /// Build the <c>cellXf</c> for a style, including its alignment and (when applicable)
    /// protection children.
    /// </summary>
    private static CellFormat BuildCellFormat(StyleInfo styleInfo)
    {
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

        return cellFormat;
    }

    private static int GetOpenXmlTextRotation(XLAlignmentValue alignment)
    {
        var textRotation = alignment.TextRotation;
        return textRotation >= 0
            ? textRotation
            : 90 - textRotation;
    }

    private static void ResolveCellStyleFormats(Stylesheet stylesheet,
        SaveContext context)
    {
        stylesheet.CellStyleFormats ??= new CellStyleFormats();

        foreach (var styleInfo in context.SharedStyles.Values)
        {
            var info = styleInfo;
            var foundOne =
                stylesheet.CellStyleFormats.Cast<CellFormat>()
                    .Any(f => CellFormatsAreEqual(f, info, compareAlignment: false));

            if (foundOne) continue;

            var cellStyleFormat = GetCellFormat(styleInfo);

            if (cellStyleFormat.ApplyProtection!.Value)
                cellStyleFormat.AppendChild(GetProtection(styleInfo));

            stylesheet.CellStyleFormats.AppendChild(cellStyleFormat);
        }

        stylesheet.CellStyleFormats.Count =
            (uint)stylesheet.CellStyleFormats.ChildElements.Count;
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
            p = StyleDecoder.ProtectionKey(protection, p);

        return p.Equals(xlProtection.Key);
    }

    private static bool QuotePrefixesAreEqual(BooleanValue? quotePrefix, bool includeQuotePrefix)
    {
        return OpenXmlHelper.GetBooleanValueAsBool(quotePrefix, false) == includeQuotePrefix;
    }

    private static bool AlignmentsAreEqual(Alignment? alignment, XLAlignmentValue xlAlignment)
    {
        if (alignment is null) return XLStyle.Default.Value.Alignment.Equals(xlAlignment);
        var a = StyleDecoder.AlignmentKey(alignment, XLAlignmentValue.Default.Key);
        return a.Equals(xlAlignment.Key);
    }

    private static Dictionary<XLBorderValue, BorderInfo> ResolveBorders(Stylesheet stylesheet,
        Dictionary<XLBorderValue, BorderInfo> sharedBorders)
    {
        stylesheet.Borders ??= new Borders();

        var allSharedBorders = new Dictionary<XLBorderValue, BorderInfo>();
        foreach (var borderInfo in sharedBorders.Values)
        {
            var borderId = 0;
            var foundOne = false;
            foreach (var openXmlElement in stylesheet.Borders)
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
                stylesheet.Borders.AppendChild(border);
            }

            allSharedBorders.Add(borderInfo.Border,
                borderInfo with { BorderId = (uint)borderId });
        }

        stylesheet.Borders.Count = (uint)stylesheet.Borders.ChildElements.Count;
        return allSharedBorders;
    }

    private static Border GetNewBorder(BorderInfo borderInfo, bool ignoreMod = true)
    {
        var border = new Border();
        if (borderInfo.Border.DiagonalUp != XLBorderValue.Default.DiagonalUp || ignoreMod)
            border.DiagonalUp = borderInfo.Border.DiagonalUp;

        if (borderInfo.Border.DiagonalDown != XLBorderValue.Default.DiagonalDown || ignoreMod)
            border.DiagonalDown = borderInfo.Border.DiagonalDown;

        AppendBorderSideWithColor<LeftBorder>(border, borderInfo.Border.LeftBorder, XLBorderValue.Default.LeftBorder,
            borderInfo.Border.LeftBorderColor, XLBorderValue.Default.LeftBorderColor, ignoreMod);
        AppendBorderSideWithColor<RightBorder>(border, borderInfo.Border.RightBorder, XLBorderValue.Default.RightBorder,
            borderInfo.Border.RightBorderColor, XLBorderValue.Default.RightBorderColor, ignoreMod);
        AppendBorderSideWithColor<TopBorder>(border, borderInfo.Border.TopBorder, XLBorderValue.Default.TopBorder,
            borderInfo.Border.TopBorderColor, XLBorderValue.Default.TopBorderColor, ignoreMod);
        AppendBorderSideWithColor<BottomBorder>(border, borderInfo.Border.BottomBorder, XLBorderValue.Default.BottomBorder,
            borderInfo.Border.BottomBorderColor, XLBorderValue.Default.BottomBorderColor, ignoreMod);

        if (borderInfo.Border.DiagonalBorder != XLBorderValue.Default.DiagonalBorder || ignoreMod)
        {
            var diagonalBorder = new DiagonalBorder { Style = borderInfo.Border.DiagonalBorder.ToOpenXml() };
            if ((borderInfo.Border.DiagonalBorderColor != XLBorderValue.Default.DiagonalBorderColor || ignoreMod)
                && borderInfo.Border.DiagonalBorderColor != null)
            {
                var diagonalBorderColor = new Color().FromXLiburColor<Color>(borderInfo.Border.DiagonalBorderColor);
                diagonalBorder.AppendChild(diagonalBorderColor);
            }

            border.AppendChild(diagonalBorder);
        }

        return border;
    }

    private static Alignment? GetNewDifferentialAlignment(XLAlignmentValue alignment)
    {
        var d = XLAlignmentValue.Default;
        if (alignment.Horizontal == d.Horizontal &&
            alignment.Vertical == d.Vertical &&
            alignment.Indent == d.Indent &&
            alignment.ReadingOrder == d.ReadingOrder &&
            alignment.WrapText == d.WrapText &&
            alignment.TextRotation == d.TextRotation &&
            alignment.ShrinkToFit == d.ShrinkToFit &&
            alignment.RelativeIndent == d.RelativeIndent &&
            alignment.JustifyLastLine == d.JustifyLastLine)
        {
            return null;
        }

        var result = new Alignment();
        if (alignment.Horizontal != d.Horizontal)
            result.Horizontal = alignment.Horizontal.ToOpenXml();
        if (alignment.Vertical != d.Vertical)
            result.Vertical = alignment.Vertical.ToOpenXml();
        if (alignment.Indent != d.Indent)
            result.Indent = (uint)alignment.Indent;
        if (alignment.ReadingOrder != d.ReadingOrder)
            result.ReadingOrder = alignment.ReadingOrder.ToOpenXml();
        if (alignment.WrapText != d.WrapText)
            result.WrapText = alignment.WrapText;
        if (alignment.TextRotation != d.TextRotation)
        {
            var textRotation = alignment.TextRotation;
            result.TextRotation = (uint)(textRotation >= 0 ? textRotation : 90 - textRotation);
        }
        if (alignment.ShrinkToFit != d.ShrinkToFit)
            result.ShrinkToFit = alignment.ShrinkToFit;
        if (alignment.RelativeIndent != d.RelativeIndent)
            result.RelativeIndent = alignment.RelativeIndent;
        if (alignment.JustifyLastLine != d.JustifyLastLine)
            result.JustifyLastLine = alignment.JustifyLastLine;

        return result;
    }

    private static void AppendBorderSideWithColor<TSide>(Border border,
        XLBorderStyleValues sideStyle, XLBorderStyleValues defaultStyle,
        XLColor sideColor, XLColor defaultColor, bool ignoreMod)
        where TSide : BorderPropertiesType, new()
    {
        if (sideStyle == defaultStyle && !ignoreMod)
            return;

        var side = new TSide { Style = sideStyle.ToOpenXml() };
        if (sideColor != defaultColor || ignoreMod)
        {
            var color = new Color().FromXLiburColor<Color>(sideColor);
            side.AppendChild(color);
        }

        border.AppendChild(side);
    }

    private static bool BordersAreEqual(Border border, XLBorderValue xlBorder)
    {
        var convertedBorder = StyleDecoder.BorderKey(
            border,
            XLBorderValue.Default.Key);
        return convertedBorder.Equals(xlBorder.Key);
    }

    private static Dictionary<XLFillValue, FillInfo> ResolveFills(Stylesheet stylesheet,
        Dictionary<XLFillValue, FillInfo> sharedFills)
    {
        stylesheet.Fills ??= new Fills();

        var fills = stylesheet.Fills;

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

        fills.Count = (uint)fills.ChildElements.Count;
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

            // Every remaining pattern (the Dark*, Light*, Gray* and MediumGray family) writes both
            // colors the same way, so they share the default arm rather than being listed as empty
            // fall-through labels.
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
        // Was a throwaway XLFill mutated through the load-side decoder purely to reach its key -
        // the same shape spec 28 removes everywhere else. One decode, straight to the key.
        return StyleDecoder.FillKey(f, fromDifferentialFormat, XLFillValue.Default.Key)
            .Equals(xlFill.Key);
    }

    private static void ResolveFonts(Stylesheet stylesheet, Dictionary<XLFontValue, FontInfo> sharedFonts)
    {
        stylesheet.Fonts ??= new Fonts();

        var newFonts = new Dictionary<XLFontValue, FontInfo>();
        foreach (var fontInfo in sharedFonts.Values)
        {
            var fontId = 0;
            var foundOne = false;
            foreach (var openXmlElement in stylesheet.Fonts)
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
                stylesheet.Fonts.AppendChild(font);
            }

            newFonts.Add(fontInfo.Font, new FontInfo { Font = fontInfo.Font, FontId = (uint)fontId });
        }

        sharedFonts.Clear();
        foreach (var kp in newFonts)
            sharedFonts.Add(kp.Key, kp.Value);

        stylesheet.Fonts.Count = (uint)stylesheet.Fonts.ChildElements.Count;
    }

    private static Font GetNewFont(FontInfo fontInfo, bool ignoreMod = true)
    {
        var font = new Font();
        var f = fontInfo.Font;
        var d = XLFontValue.Default;

        AppendFontFlagElements(font, f, d, ignoreMod);
        AppendFontScalarElements(font, f, d, ignoreMod);

        return font;
    }

#pragma warning disable S3776 // Each property check is independent and flat
    private static void AppendFontFlagElements(Font font, XLFontValue f, XLFontValue d, bool ignoreMod)
    {
        if ((f.Bold != d.Bold || ignoreMod) && f.Bold)
            font.AppendChild(new Bold());

        if ((f.Italic != d.Italic || ignoreMod) && f.Italic)
            font.AppendChild(new Italic());

        if ((f.Underline != d.Underline || ignoreMod) && f.Underline != XLFontUnderlineValues.None)
            font.AppendChild(new Underline { Val = f.Underline.ToOpenXml() });

        if ((f.Strikethrough != d.Strikethrough || ignoreMod) && f.Strikethrough)
            font.AppendChild(new Strike());

        if (f.VerticalAlignment != d.VerticalAlignment || ignoreMod)
            font.AppendChild(new VerticalTextAlignment { Val = f.VerticalAlignment.ToOpenXml() });

        if ((f.Shadow != d.Shadow || ignoreMod) && f.Shadow)
            font.AppendChild(new Shadow());
    }
#pragma warning restore S3776

#pragma warning disable S3776 // Each property check is independent and flat
    private static void AppendFontScalarElements(Font font, XLFontValue f, XLFontValue d, bool ignoreMod)
    {
        if (!XLHelper.AreEqual(f.FontSize, d.FontSize) || ignoreMod)
            font.AppendChild(new FontSize { Val = f.FontSize });

        if (f.FontColor != d.FontColor || ignoreMod)
            font.AppendChild(new Color().FromXLiburColor<Color>(f.FontColor));

        if (f.FontName != d.FontName || ignoreMod)
            font.AppendChild(new FontName { Val = f.FontName });

        if (f.FontFamilyNumbering != d.FontFamilyNumbering || ignoreMod)
            font.AppendChild(new FontFamilyNumbering { Val = (int)f.FontFamilyNumbering });

        if ((f.FontCharSet != d.FontCharSet || ignoreMod) && f.FontCharSet != XLFontCharSet.Default)
            font.AppendChild(new FontCharSet { Val = (int)f.FontCharSet });

        if ((f.FontScheme != d.FontScheme || ignoreMod) && f.FontScheme != XLFontScheme.None)
            font.AppendChild(new FontScheme { Val = f.FontScheme.ToOpenXmlEnum() });
    }
#pragma warning restore S3776    

    private static bool FontsAreEqual(Font font, XLFontValue xlFont)
    {
        var convertedFont = StyleDecoder.FontKey(
            font,
            XLFontValue.Default.Key);
        return convertedFont.Equals(xlFont.Key);
    }

    private static Dictionary<XLNumberFormatValue, NumberFormatInfo> ResolveNumberFormats(
        Stylesheet stylesheet,
        HashSet<XLNumberFormatValue> customNumberFormats,
        uint defaultFormatId)
    {
        if (stylesheet.NumberingFormats == null)
        {
            stylesheet.NumberingFormats = new NumberingFormats();
            stylesheet.NumberingFormats.AppendChild(new NumberingFormat
            {
                NumberFormatId = 0,
                FormatCode = ""
            });
        }

        var allSharedNumberFormats = new Dictionary<XLNumberFormatValue, NumberFormatInfo>();
        var partNumberingFormats = stylesheet.NumberingFormats;

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
            foreach (var nf in stylesheet.NumberingFormats.Cast<NumberingFormat>())
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
                stylesheet.NumberingFormats.AppendChild(partNumberFormat);
            }

            allSharedNumberFormats.Add(customNumberFormat,
                new NumberFormatInfo
                {
                    NumberFormat = customNumberFormat,
                    NumberFormatId = (int)partNumberFormat.NumberFormatId!.Value
                });
        }

        stylesheet.NumberingFormats.Count =
            (uint)stylesheet.NumberingFormats.ChildElements.Count;
        return allSharedNumberFormats;
    }

    private static bool CustomNumberFormatsAreEqual(NumberingFormat nf, XLNumberFormatValue xlNumberFormat)
    {
        if (nf.FormatCode != null && !string.IsNullOrWhiteSpace(nf.FormatCode.Value))
            return string.Equals(xlNumberFormat.Format, nf.FormatCode.Value);

        return false;
    }
}
