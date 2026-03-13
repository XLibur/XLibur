using XLibur.Extensions;
using XLibur.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using XLibur.Excel.IO;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Op = DocumentFormat.OpenXml.CustomProperties;

namespace XLibur.Excel;

using Ap;
using Drawings;
using Op;

// ReSharper disable once InconsistentNaming
public partial class XLWorkbook
{
    private void Load(string file)
    {
        LoadSheets(file);
    }

    private void Load(Stream stream)
    {
        LoadSheets(stream);
    }

    private void LoadSheets(string fileName)
    {
        using var dSpreadsheet = SpreadsheetDocument.Open(fileName, false);
        LoadSpreadsheetDocument(dSpreadsheet);
    }

    private void LoadSheets(Stream stream)
    {
        using var dSpreadsheet = SpreadsheetDocument.Open(stream, false);
        LoadSpreadsheetDocument(dSpreadsheet);
    }

    private void LoadSheetsFromTemplate(string fileName)
    {
        using (var dSpreadsheet = SpreadsheetDocument.CreateFromTemplate(fileName))
            LoadSpreadsheetDocument(dSpreadsheet);

        // If we load a workbook as a template, we have to treat it as a "new" workbook.
        // The original file will NOT be copied into place before changes are applied
        // Hence all loaded RelIds have to be cleared
        ResetAllRelIds();
    }

    private void ResetAllRelIds()
    {
        foreach (var pc in PivotCachesInternal)
            pc.WorkbookCacheRelId = null;

        var sheetId = 1u;
        foreach (var ws in WorksheetsInternal)
        {
            // Ensure unique sheetId for each sheet.
            ws.SheetId = sheetId++;
            ws.RelId = null;

            foreach (var pt in ws.PivotTables.Cast<XLPivotTable>())
            {
                pt.CacheDefinitionRelId = null;
                pt.RelId = null;
            }

            foreach (var picture in ws.Pictures.Cast<XLPicture>())
                picture.RelId = null;

            foreach (var table in ws.Tables.Cast<XLTable>())
                table.RelId = null;
        }
    }

    private void LoadSpreadsheetDocument(SpreadsheetDocument dSpreadsheet)
    {
        var context = new LoadContext();
        ShapeIdManager = new XLIdManager();
        SetProperties(dSpreadsheet);

        SharedStringItem[]? sharedStrings = null;
        var workbookPart = dSpreadsheet.WorkbookPart!;
        if (workbookPart.GetPartsOfType<SharedStringTablePart>().Any())
        {
            var shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            sharedStrings = shareStringPart.SharedStringTable!.Elements<SharedStringItem>().ToArray();
        }

        LoadWorkbookTheme(workbookPart.ThemePart, this);

        if (dSpreadsheet.CustomFilePropertiesPart != null)
        {
            foreach (var m in dSpreadsheet.CustomFilePropertiesPart.Properties!.Elements<CustomDocumentProperty>())
            {
                var name = m.Name?.Value;

                if (string.IsNullOrWhiteSpace(name))
                    continue;

                if (m.VTLPWSTR != null)
                    CustomProperties.Add(name, m.VTLPWSTR.Text);
                else if (m.VTFileTime != null)
                {
                    CustomProperties.Add(name,
                        DateTime.ParseExact(m.VTFileTime.Text, "yyyy'-'MM'-'dd'T'HH':'mm':'ssK",
                            CultureInfo.InvariantCulture));
                }
                else if (m.VTDouble != null)
                    CustomProperties.Add(name, double.Parse(m.VTDouble.Text, CultureInfo.InvariantCulture));
                else if (m.VTBool != null)
                    CustomProperties.Add(name, m.VTBool.Text == "true");
            }
        }

        var wbProps = workbookPart.Workbook!.WorkbookProperties;
        if (wbProps != null)
            Use1904DateSystem = OpenXmlHelper.GetBooleanValueAsBool(wbProps.Date1904, false);

        var wbFilesharing = workbookPart.Workbook!.FileSharing;
        if (wbFilesharing != null)
        {
            FileSharing.ReadOnlyRecommended =
                OpenXmlHelper.GetBooleanValueAsBool(wbFilesharing.ReadOnlyRecommended, false);
            FileSharing.UserName = wbFilesharing.UserName?.Value;
        }

        LoadWorkbookProtection(workbookPart.Workbook!.WorkbookProtection, this);

        var calculationProperties = workbookPart.Workbook!.CalculationProperties;
        if (calculationProperties != null)
        {
            var calculateMode = calculationProperties.CalculationMode;
            if (calculateMode != null)
                CalculateMode = calculateMode.Value.ToXLibur();

            var calculationOnSave = calculationProperties.CalculationOnSave;
            if (calculationOnSave != null)
                CalculationOnSave = calculationOnSave.Value;

            var forceFullCalculation = calculationProperties.ForceFullCalculation;
            if (forceFullCalculation != null)
                ForceFullCalculation = forceFullCalculation.Value;

            var fullCalculationOnLoad = calculationProperties.FullCalculationOnLoad;
            if (fullCalculationOnLoad != null)
                FullCalculationOnLoad = fullCalculationOnLoad.Value;

            var fullPrecision = calculationProperties.FullPrecision;
            if (fullPrecision != null)
                FullPrecision = fullPrecision.Value;

            var referenceMode = calculationProperties.ReferenceMode;
            if (referenceMode != null)
                ReferenceStyle = referenceMode.Value.ToXLibur();
        }

        var efp = dSpreadsheet.ExtendedFilePropertiesPart;
        if (efp is { Properties: not null })
        {
            if (efp.Properties.Elements<Company>().Any())
                Properties.Company = efp.Properties.GetFirstChild<Company>()!.Text;

            if (efp.Properties.Elements<Manager>().Any())
                Properties.Manager = efp.Properties.GetFirstChild<Manager>()!.Text;
        }

        var s = workbookPart.WorkbookStylesPart?.Stylesheet;
        var numberingFormats = s?.NumberingFormats;
        context.LoadNumberFormats(numberingFormats);
        var fills = s?.Fills;
        var borders = s?.Borders;
        var fonts = s?.Fonts;
        var dfCount = 0;
        var differentialFormats = s is { DifferentialFormats: not null }
            ? s.DifferentialFormats.Elements<DifferentialFormat>().ToDictionary(_ => dfCount++)
            : new Dictionary<int, DifferentialFormat>();

        // If the loaded workbook has a changed "Normal" style, it might affect the default width of a column.
        var normalStyle = s?.CellStyles?.Elements<CellStyle>()
            .FirstOrDefault(x => x.BuiltinId is not null && x.BuiltinId.Value == 0);
        if (normalStyle != null)
        {
            var normalStyleKey = ((XLStyle)Style).Key;
            WorksheetSheetDataReader.LoadStyle(ref normalStyleKey, (int)normalStyle.FormatId!.Value, s!, fills!, borders!, fonts!, numberingFormats);
            Style = new XLStyle(null!, normalStyleKey);
            ColumnWidth = CalculateColumnWidth(8, Style.Font, this);
        }

        // We loop through the sheets in 2 passes: first just to add the sheets and second to add all the data for the sheets.
        // We do this mainly because it skips a very costly calculation invalidation step, but it also make things more consistent,
        // e.g. when reading calculations that reference other sheets, we know that those sheets always already exist.
        // That consistency point isn't required yet but could be taken advantage of in the future.
        var sheets = workbookPart.Workbook!.Sheets;
        var position = 0;
        foreach (var dSheet in sheets!.OfType<Sheet>())
        {
            position++;
            string sheetName = dSheet.Name!.Value!;
            var sheetIdValue = dSheet.SheetId!.Value;

            if (string.IsNullOrEmpty(dSheet.Id))
            {
                // Some non-Excel producers create sheets with empty relId.
                var emptySheet = WorksheetsInternal.Add(sheetName, position, sheetIdValue);
                if (dSheet.State != null)
                    emptySheet.Visibility = dSheet.State.Value.ToXLibur();

                continue;
            }

            // Although the relationship to worksheet is most common, there can be other types
            // than worksheet, e.g., chartSheet. Since we can't load them, add them to the list
            // of unsupported sheets and copy them when saving. See Codeplex #6932.
            if (workbookPart.GetPartById(dSheet.Id!.Value!) is not WorksheetPart)
            {
                UnsupportedSheets.Add(new UnsupportedSheet { SheetId = sheetIdValue, Position = position });
                continue;
            }

            var ws = WorksheetsInternal.Add(sheetName, position, sheetIdValue);
            ws.RelId = dSheet.Id;

            if (dSheet.State != null)
                ws.Visibility = dSheet.State.Value.ToXLibur();
        }

        position = 0;
        foreach (var dSheet in sheets!.OfType<Sheet>())
        {
            position++;
            string sheetName = dSheet.Name!.Value!;

            if (string.IsNullOrEmpty(dSheet.Id))
            {
                // Some non-Excel producers create sheets with empty relId.
                continue;
            }

            // Although relationship to worksheet is most common, there can be other types
            // than worksheet, e.g. chartSheet. Since we can't load them, add them to list
            // of unsupported sheets and copy them when saving. See Codeplex #6932.
            var worksheetPart = workbookPart.GetPartById(dSheet.Id!.Value!) as WorksheetPart;
            if (worksheetPart == null)
            {
                continue;
            }

            var sharedFormulasR1C1 = new Dictionary<uint, string>();
            if (!WorksheetsInternal.TryGetWorksheet(sheetName, out var ws))
            {
                // This shouldn't be possible, as all worksheets should have already been added in the loop before this loop
                continue;
            }

            WorksheetSheetDataReader.ApplyStyle(ws, 0, s!, fills!, borders!, fonts!, numberingFormats);

            var styleList = new Dictionary<int, IXLStyle>(); // {{0, ws.Style}};
            PageSetupProperties? pageSetupProperties = null;

            var lastRow = 0;
            var lastColumnNumber = 0;

            using (var reader = new OpenXmlPartReader(worksheetPart))
            {
                Type[] ignoredElements =
                [
                    typeof(CustomSheetViews) // Custom sheet views contain their own auto filter data, and more, which should be ignored for now
                ];

                while (reader.Read())
                {
                    while (ignoredElements.Contains(reader.ElementType))
                        reader.ReadNextSibling();

                    if (reader.ElementType == typeof(SheetFormatProperties))
                    {
                        var sheetFormatProperties = (SheetFormatProperties?)reader.LoadCurrentElement();
                        if (sheetFormatProperties != null)
                        {
                            if (sheetFormatProperties.DefaultRowHeight != null)
                                ws.RowHeight = sheetFormatProperties.DefaultRowHeight;

                            ws.RowHeightChanged = (sheetFormatProperties.CustomHeight != null &&
                                                   sheetFormatProperties.CustomHeight.Value);

                            if (sheetFormatProperties.DefaultColumnWidth != null)
                                ws.ColumnWidth =
                                    XLHelper.ConvertWidthToNoC(sheetFormatProperties.DefaultColumnWidth.Value,
                                        ws.Style.Font, this);
                            else if (sheetFormatProperties.BaseColumnWidth != null)
                                ws.ColumnWidth = CalculateColumnWidth(sheetFormatProperties.BaseColumnWidth.Value,
                                    ws.Style.Font, this);
                        }
                    }
                    else if (reader.ElementType == typeof(SheetViews))
                        WorksheetElementReader.LoadSheetViews((SheetViews)reader.LoadCurrentElement()!, ws);
                    else if (reader.ElementType == typeof(MergeCells))
                    {
                        var mergedCells = (MergeCells?)reader.LoadCurrentElement();
                        if (mergedCells != null)
                        {
                            foreach (var mergeCell in mergedCells.Elements<MergeCell>())
                                ws.Range(mergeCell.Reference!)!.Merge(false);
                        }
                    }
                    else if (reader.ElementType == typeof(Columns))
                        WorksheetSheetDataReader.LoadColumns(s!, numberingFormats, fills!, borders!, fonts!, ws,
                            (Columns)reader.LoadCurrentElement()!);
                    else if (reader.ElementType == typeof(Row))
                    {
                        WorksheetSheetDataReader.LoadRow(s!, numberingFormats, fills!, borders!, fonts!, ws, sharedStrings, sharedFormulasR1C1,
                            styleList, reader, ref lastRow, ref lastColumnNumber, Use1904DateSystem);
                    }
                    else if (reader.ElementType == typeof(AutoFilter))
                        WorksheetElementReader.LoadAutoFilter((AutoFilter)reader.LoadCurrentElement()!, ws, differentialFormats);
                    else if (reader.ElementType == typeof(SheetProtection))
                        WorksheetElementReader.LoadSheetProtection((SheetProtection)reader.LoadCurrentElement()!, ws);
                    else if (reader.ElementType == typeof(DataValidations))
                        WorksheetElementReader.LoadDataValidations((DataValidations)reader.LoadCurrentElement()!, ws);
                    else if (reader.ElementType == typeof(ConditionalFormatting))
                        ConditionalFormatReader.LoadConditionalFormatting((ConditionalFormatting)reader.LoadCurrentElement()!, ws,
                            differentialFormats, context);
                    else if (reader.ElementType == typeof(Hyperlinks))
                        WorksheetElementReader.LoadHyperlinks((Hyperlinks)reader.LoadCurrentElement()!, worksheetPart, ws);
                    else if (reader.ElementType == typeof(PrintOptions))
                        WorksheetElementReader.LoadPrintOptions((PrintOptions)reader.LoadCurrentElement()!, ws);
                    else if (reader.ElementType == typeof(PageMargins))
                        WorksheetElementReader.LoadPageMargins((PageMargins)reader.LoadCurrentElement()!, ws);
                    else if (reader.ElementType == typeof(PageSetup))
                        WorksheetElementReader.LoadPageSetup((PageSetup)reader.LoadCurrentElement()!, ws, pageSetupProperties);
                    else if (reader.ElementType == typeof(HeaderFooter))
                        WorksheetElementReader.LoadHeaderFooter((HeaderFooter)reader.LoadCurrentElement()!, ws);
                    else if (reader.ElementType == typeof(SheetProperties))
                        WorksheetElementReader.LoadSheetProperties((SheetProperties)reader.LoadCurrentElement()!, ws, out pageSetupProperties);
                    else if (reader.ElementType == typeof(RowBreaks))
                        WorksheetElementReader.LoadRowBreaks((RowBreaks)reader.LoadCurrentElement()!, ws);
                    else if (reader.ElementType == typeof(ColumnBreaks))
                        WorksheetElementReader.LoadColumnBreaks((ColumnBreaks)reader.LoadCurrentElement()!, ws);
                    else if (reader.ElementType == typeof(WorksheetExtensionList))
                        ConditionalFormatReader.LoadExtensions((WorksheetExtensionList)reader.LoadCurrentElement()!, ws, this);
                    else if (reader.ElementType == typeof(LegacyDrawing))
                        ws.LegacyDrawingId = ((LegacyDrawing)reader.LoadCurrentElement()!).Id?.Value;
                }

                reader.Close();
            }

            ws.ConditionalFormats.ReorderAccordingToOriginalPriority();

            #region LoadTables

            foreach (var tableDefinitionPart in worksheetPart.TableDefinitionParts)
            {
                var relId = worksheetPart.GetIdOfPart(tableDefinitionPart);
                var dTable = tableDefinitionPart.Table!;

                var reference = dTable.Reference!.Value!;
                var tableName = dTable.Name?.Value ?? dTable.DisplayName?.Value ?? string.Empty;
                if (string.IsNullOrWhiteSpace(tableName))
                    throw new InvalidDataException("The table name is missing.");

                var xlTable = ws.Range(reference)!.CreateTable(tableName, false) as XLTable;
                xlTable!.RelId = relId;

                if (dTable.HeaderRowCount != null && dTable.HeaderRowCount == 0)
                {
                    xlTable._showHeaderRow = false;
                    xlTable.AddFields(dTable.TableColumns!.Cast<TableColumn>()
                        .Select(t => DrawingPartReader.GetTableColumnName(t.Name!.Value!)));
                }
                else
                {
                    xlTable.InitializeAutoFilter();
                }

                if (dTable.TotalsRowCount != null && dTable.TotalsRowCount.Value > 0)
                    xlTable._showTotalsRow = true;

                if (dTable.TableStyleInfo != null)
                {
                    if (dTable.TableStyleInfo.ShowFirstColumn != null)
                        xlTable.EmphasizeFirstColumn = dTable.TableStyleInfo.ShowFirstColumn.Value;
                    if (dTable.TableStyleInfo.ShowLastColumn != null)
                        xlTable.EmphasizeLastColumn = dTable.TableStyleInfo.ShowLastColumn.Value;
                    if (dTable.TableStyleInfo.ShowRowStripes != null)
                        xlTable.ShowRowStripes = dTable.TableStyleInfo.ShowRowStripes.Value;
                    if (dTable.TableStyleInfo.ShowColumnStripes != null)
                        xlTable.ShowColumnStripes = dTable.TableStyleInfo.ShowColumnStripes.Value;
                    if (dTable.TableStyleInfo.Name != null)
                    {
                        var theme = XLTableTheme.FromName(dTable.TableStyleInfo.Name.Value!);
                        xlTable.Theme = theme ?? new XLTableTheme(dTable.TableStyleInfo.Name.Value!);
                    }
                    else
                        xlTable.Theme = XLTableTheme.None;
                }

                if (dTable.AutoFilter != null)
                {
                    xlTable.ShowAutoFilter = true;
                    WorksheetElementReader.LoadAutoFilterColumns(dTable.AutoFilter, xlTable.AutoFilter);
                }
                else
                    xlTable.ShowAutoFilter = false;

                if (xlTable.ShowTotalsRow)
                {
                    foreach (var tableColumn in dTable.TableColumns!.Cast<TableColumn>())
                    {
                        var tableColumnName = DrawingPartReader.GetTableColumnName(tableColumn.Name!.Value!);
                        if (tableColumn.TotalsRowFunction != null)
                            xlTable.Field(tableColumnName).TotalsRowFunction =
                                tableColumn.TotalsRowFunction.Value.ToXLibur();

                        if (tableColumn.TotalsRowFormula != null)
                            xlTable.Field(tableColumnName).TotalsRowFormulaA1 =
                                tableColumn.TotalsRowFormula.Text;

                        if (tableColumn.TotalsRowLabel != null)
                            xlTable.Field(tableColumnName).TotalsRowLabel = tableColumn.TotalsRowLabel.Value;
                    }

                    if (xlTable.AutoFilter != null)
                        xlTable.AutoFilter.Range = xlTable.Worksheet.Range(
                            xlTable.RangeAddress.FirstAddress.RowNumber, xlTable.RangeAddress.FirstAddress.ColumnNumber,
                            xlTable.RangeAddress.LastAddress.RowNumber - 1,
                            xlTable.RangeAddress.LastAddress.ColumnNumber);
                }
                else if (xlTable.AutoFilter != null)
                    xlTable.AutoFilter.Range = xlTable.Worksheet.Range(xlTable.RangeAddress);
            }

            #endregion LoadTables

            DrawingPartReader.LoadDrawings(worksheetPart, ws);

            #region LoadComments

            if (worksheetPart.WorksheetCommentsPart != null)
            {
                var root = worksheetPart.WorksheetCommentsPart.Comments!;
                var authors = root.GetFirstChild<Authors>()!.ChildElements.OfType<Author>().ToList();
                var comments = root.GetFirstChild<CommentList>()!.ChildElements.OfType<Comment>().ToList();

                // **** MAYBE FUTURE SHAPE SIZE SUPPORT
                var shapes = DrawingPartReader.GetCommentShapes(worksheetPart);

                for (var i = 0; i < comments.Count; i++)
                {
                    var c = comments[i];

                    XElement? shape = null;
                    if (i < shapes.Count)
                        shape = shapes[i];

                    // find cell by reference
                    var cell = ws.Cell(c.Reference!);

                    var shapeIdString = shape?.Attribute("id")?.Value;
                    if (shapeIdString?.StartsWith("_x0000_s") ?? false)
                        shapeIdString = shapeIdString[8..];

                    int? shapeId = int.TryParse(shapeIdString, out var sid) ? sid : null;
                    var xlComment = cell!.CreateComment(shapeId);

                    xlComment.Author = authors[(int)c.AuthorId!.Value].InnerText;
                    ShapeIdManager.Add(xlComment.ShapeId);

                    var commentTextNode = c.GetFirstChild<CommentText>()!;
                    var runs = commentTextNode.Elements<Run>();
                    foreach (var run in runs)
                    {
                        var runProperties = run.RunProperties;
                        var text = run.Text!.InnerText.FixNewLines();
                        var rt = xlComment.AddText(text);
                        OpenXmlHelper.LoadFont(runProperties, rt);
                    }

                    // Comments can have text not wrapped in a Run element (e.g. Google Sheets exports)
                    if (commentTextNode.Text != null)
                    {
                        var plainText = commentTextNode.Text.Text.FixNewLines();
                        xlComment.AddText(plainText);
                    }

                    if (shape != null)
                    {
                        DrawingPartReader.LoadShapeProperties(xlComment, shape);

                        var clientData = shape.Elements().First(e => e.Name.LocalName == "ClientData");
                        DrawingPartReader.LoadClientData(xlComment, clientData);

                        var textBox = shape.Elements().FirstOrDefault(e => e.Name.LocalName == "textbox");
                        if (textBox is not null)
                            DrawingPartReader.LoadTextBox(xlComment, textBox, DpiX, DpiY);

                        var alt = shape.Attribute("alt");
                        if (alt != null) xlComment.Style.Web.SetAlternateText(alt.Value);

                        DrawingPartReader.LoadColorsAndLines(xlComment, shape);
                    }
                }
            }

            #endregion LoadComments
        }

        var workbook = workbookPart.Workbook;

        var bookViews = workbook.BookViews;
        if (bookViews?.FirstOrDefault() is WorkbookView workbookView)
        {
            if (workbookView.ActiveTab == null || !workbookView.ActiveTab.HasValue)
            {
                Worksheets.First().SetTabActive().Unhide();
            }
            else
            {
                var unsupportedSheet =
                    UnsupportedSheets.FirstOrDefault(us => us.Position == (int)(workbookView.ActiveTab.Value + 1));
                if (unsupportedSheet != null)
                    unsupportedSheet.IsActive = true;
                else
                {
                    Worksheet((int)(workbookView.ActiveTab.Value + 1)).SetTabActive();
                }
            }
        }

        DefinedNameReader.LoadDefinedNames(workbook, this);

        PivotTableCacheDefinitionPartReader.Load(workbookPart, this);

        // Delay loading of pivot tables until all sheets have been loaded
        foreach (var dSheet in sheets!.OfType<Sheet>())
        {
            if (string.IsNullOrEmpty(dSheet.Id))
            {
                // Some non-Excel producers create sheets with empty relId.
                continue;
            }

            // The referenced sheet can also be ChartsheetPart. Only look for pivot tables in normal sheet parts.
            if (workbookPart.GetPartById(dSheet.Id!.Value!) is WorksheetPart worksheetPart)
            {
                var ws = (XLWorksheet)WorksheetsInternal.Worksheet(dSheet.Name!.Value!);

                foreach (var pivotTablePart in worksheetPart.PivotTableParts)
                {
                    PivotTableDefinitionPartReader.Load(workbookPart, differentialFormats, pivotTablePart,
                        worksheetPart, ws, context);
                }
            }
        }
    }

    /// <summary>
    /// Calculate expected column width as a number displayed in the column in Excel from
    /// the number of characters that should fit into the width and a font.
    /// </summary>
    internal static double CalculateColumnWidth(double charWidth, IXLFont font, XLWorkbook workbook)
    {
        // Convert width as a number of characters and translate it into a given number of pixels.
        var mdw = workbook.GraphicEngine.GetMaxDigitWidth(font, workbook.DpiX).RoundToInt();
        var defaultColWidthPx = XLHelper.NoCToPixels(charWidth, mdw).RoundToInt();

        // Excel then rounds this number up to the nearest multiple of 8 pixels so that
        // scrolling across columns and rows is faster.
        var roundUpToMultiple = defaultColWidthPx + (8 - defaultColWidthPx % 8);

        // and last, convert the width in pixels to width displayed in Excel. Shouldn't round the number, because
        // it causes inconsistency with conversion to other units, but other places in XLibur do = keep for now.
        var defaultColumnWidth = XLHelper.PixelToNoC(roundUpToMultiple, mdw).Round(2);
        return defaultColumnWidth;
    }

    private static void LoadWorkbookTheme(ThemePart? tp, XLWorkbook wb)
    {
        var colorScheme = tp?.Theme?.ThemeElements?.ColorScheme;
        if (colorScheme == null) return;
        var background1 = colorScheme.Light1Color?.RgbColorModelHex?.Val?.Value;
        if (!string.IsNullOrEmpty(background1))
        {
            wb.Theme.Background1 = XLColor.FromHexRgb(background1);
        }

        var text1 = colorScheme.Dark1Color?.RgbColorModelHex?.Val?.Value;
        if (!string.IsNullOrEmpty(text1))
        {
            wb.Theme.Text1 = XLColor.FromHexRgb(text1);
        }

        var background2 = colorScheme.Light2Color?.RgbColorModelHex?.Val?.Value;
        if (!string.IsNullOrEmpty(background2))
        {
            wb.Theme.Background2 = XLColor.FromHexRgb(background2);
        }

        var text2 = colorScheme.Dark2Color?.RgbColorModelHex?.Val?.Value;
        if (!string.IsNullOrEmpty(text2))
        {
            wb.Theme.Text2 = XLColor.FromHexRgb(text2);
        }

        var accent1 = colorScheme.Accent1Color?.RgbColorModelHex?.Val?.Value;
        if (!string.IsNullOrEmpty(accent1))
        {
            wb.Theme.Accent1 = XLColor.FromHexRgb(accent1);
        }

        var accent2 = colorScheme.Accent2Color?.RgbColorModelHex?.Val?.Value;
        if (!string.IsNullOrEmpty(accent2))
        {
            wb.Theme.Accent2 = XLColor.FromHexRgb(accent2);
        }

        var accent3 = colorScheme.Accent3Color?.RgbColorModelHex?.Val?.Value;
        if (!string.IsNullOrEmpty(accent3))
        {
            wb.Theme.Accent3 = XLColor.FromHexRgb(accent3);
        }

        var accent4 = colorScheme.Accent4Color?.RgbColorModelHex?.Val?.Value;
        if (!string.IsNullOrEmpty(accent4))
        {
            wb.Theme.Accent4 = XLColor.FromHexRgb(accent4);
        }

        var accent5 = colorScheme.Accent5Color?.RgbColorModelHex?.Val?.Value;
        if (!string.IsNullOrEmpty(accent5))
        {
            wb.Theme.Accent5 = XLColor.FromHexRgb(accent5);
        }

        var accent6 = colorScheme.Accent6Color?.RgbColorModelHex?.Val?.Value;
        if (!string.IsNullOrEmpty(accent6))
        {
            wb.Theme.Accent6 = XLColor.FromHexRgb(accent6);
        }

        var hyperlink = colorScheme.Hyperlink?.RgbColorModelHex?.Val?.Value;
        if (!string.IsNullOrEmpty(hyperlink))
        {
            wb.Theme.Hyperlink = XLColor.FromHexRgb(hyperlink);
        }

        var followedHyperlink = colorScheme.FollowedHyperlinkColor?.RgbColorModelHex?.Val?.Value;
        if (!string.IsNullOrEmpty(followedHyperlink))
        {
            wb.Theme.FollowedHyperlink = XLColor.FromHexRgb(followedHyperlink);
        }
    }

    private static void LoadWorkbookProtection(WorkbookProtection? wp, XLWorkbook wb)
    {
        if (wp == null) return;

        wb.Protection.IsProtected = true;

        var algorithmName = wp.WorkbookAlgorithmName?.Value ?? string.Empty;
        if (string.IsNullOrEmpty(algorithmName))
        {
            wb.Protection.PasswordHash = wp.WorkbookPassword?.Value ?? string.Empty;
            wb.Protection.Base64EncodedSalt = string.Empty;
        }
        else if (DescribedEnumParser<XLProtectionAlgorithm.Algorithm>.IsValidDescription(algorithmName))
        {
            wb.Protection.Algorithm =
                DescribedEnumParser<XLProtectionAlgorithm.Algorithm>.FromDescription(algorithmName);
            wb.Protection.PasswordHash = wp.WorkbookHashValue?.Value ?? string.Empty;
            wb.Protection.SpinCount = wp.WorkbookSpinCount?.Value ?? 0;
            wb.Protection.Base64EncodedSalt = wp.WorkbookSaltValue?.Value ?? string.Empty;
        }

        wb.Protection.AllowElement(XLWorkbookProtectionElements.Structure,
            !OpenXmlHelper.GetBooleanValueAsBool(wp.LockStructure, false));
        wb.Protection.AllowElement(XLWorkbookProtectionElements.Windows,
            !OpenXmlHelper.GetBooleanValueAsBool(wp.LockWindows, false));
    }

    private void SetProperties(SpreadsheetDocument dSpreadsheet)
    {
        var p = dSpreadsheet.PackageProperties;
        Properties.Author = p.Creator;
        Properties.Category = p.Category;
        Properties.Comments = p.Description;
        if (p.Created != null)
            Properties.Created = p.Created.Value;
        if (p.Modified != null)
            Properties.Modified = p.Modified.Value;
        Properties.Keywords = p.Keywords;
        Properties.LastModifiedBy = p.LastModifiedBy;
        Properties.Status = p.ContentStatus;
        Properties.Subject = p.Subject;
        Properties.Title = p.Title;
    }
}
