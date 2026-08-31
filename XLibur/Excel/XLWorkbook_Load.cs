using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.Drawings;
using XLibur.Excel.IO;
using XLibur.Excel.Tables;
using XLibur.Extensions;
using XLibur.Utils;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Op = DocumentFormat.OpenXml.CustomProperties;
using TC = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;

namespace XLibur.Excel;

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

            foreach (var chart in ws.Charts.Cast<XLChart>())
            {
                chart.RelId = null;
                chart.IsNew = true;
            }
        }
    }

    private void LoadSpreadsheetDocument(SpreadsheetDocument dSpreadsheet)
    {
        var context = new LoadContext();
        ShapeIdManager = new XLIdManager();
        SetProperties(dSpreadsheet);

        SharedStringEntry[]? sharedStrings = null;
        var workbookPart = dSpreadsheet.WorkbookPart!;
        var shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        if (shareStringPart is not null)
        {
            sharedStrings = SharedStringReader.Read(shareStringPart);

            // Pre-size the workbook's internal SST to avoid repeated resizing
            // as cell values reference these strings during sheet loading.
            if (sharedStrings.Length > 0)
                SharedStringTable.EnsureCapacity(sharedStrings.Length);
        }

        LoadWorkbookTheme(workbookPart.ThemePart, this);

        RichDataReader.LoadRichData(workbookPart, this, context);

        context.LoadDynamicArrayMetadata(workbookPart.CellMetadataPart?.Metadata);

        LoadCustomFileProperties(dSpreadsheet);

        if (workbookPart.Workbook!.WorkbookProperties is { } wbProps)
            Use1904DateSystem = OpenXmlHelper.GetBooleanValueAsBool(wbProps.Date1904, false);

        if (workbookPart.Workbook.FileSharing is { } wbFilesharing)
        {
            FileSharing.ReadOnlyRecommended =
                OpenXmlHelper.GetBooleanValueAsBool(wbFilesharing.ReadOnlyRecommended, false);
            FileSharing.UserName = wbFilesharing.UserName?.Value;
        }

        LoadWorkbookProtection(workbookPart.Workbook.WorkbookProtection, this);

        LoadCalculationProperties(workbookPart.Workbook.CalculationProperties);

        LoadExtendedFileProperties(dSpreadsheet);

        var s = workbookPart.WorkbookStylesPart?.Stylesheet;
        var numberingFormats = s?.NumberingFormats;
        var differentialFormats = s?.DifferentialFormats
            ?.Elements<DifferentialFormat>()
            .Select((df, i) => (df, i))
            .ToDictionary(x => x.i, x => x.df)
            ?? new Dictionary<int, DifferentialFormat>();

        context.Styles = new StylesheetData(s, numberingFormats, s?.Fills, s?.Borders, s?.Fonts, differentialFormats);

        // If the loaded workbook has a changed "Normal" style, it might affect the default width of a column.
        var normalStyle = s?.CellStyles?.Elements<CellStyle>()
            .FirstOrDefault(x => x.BuiltinId is not null && x.BuiltinId.Value == 0);
        if (normalStyle != null)
        {
            var normalStyleKey = StyleDecoder.Decode((int)normalStyle.FormatId!.Value,
                context.Styles, ((XLStyle)Style).Key);
            Style = new XLStyle(null!, normalStyleKey);
            ColumnWidth = CalculateColumnWidth(8, Style.Font, this);
        }

        // We loop through the sheets in 2 passes: first just to add the sheets and second to add all the data for the sheets.
        // We do this mainly because it skips a very costly calculation invalidation step, but it also make things more consistent,
        // e.g. when reading calculations that reference other sheets, we know that those sheets always already exist.
        // That consistency point isn't required yet but could be taken advantage of in the future.
        // Persons are workbook level and referenced by the threaded comments of every sheet, so they
        // have to be in place before the sheets are read.
        LoadPersons(workbookPart);

        var sheets = workbookPart.Workbook.Sheets;
        LoadSheetsPass1(workbookPart, sheets!);

        LoadSheetsPass2(workbookPart, sheets!, sharedStrings, context);

        LoadActiveTab(workbookPart.Workbook);

        DefinedNameReader.LoadDefinedNames(workbookPart.Workbook, this);

        PivotTableCacheDefinitionPartReader.Load(workbookPart, this);

        LoadPivotTables(workbookPart, sheets!, context);

        // Last, because a slicer binds to the pivot tables and tables it filters, and both have to
        // exist before it can find them.
        SlicerReader.LoadSlicers(workbookPart, sheets!, WorksheetsInternal);

        // Same ordering constraint as slicers: a timeline binds to the pivot tables it filters, and
        // they have to exist before it can find them.
        TimelineReader.LoadTimelines(workbookPart, sheets!, WorksheetsInternal);
    }

    private void LoadCustomFileProperties(SpreadsheetDocument dSpreadsheet)
    {
        if (dSpreadsheet.CustomFilePropertiesPart != null)
        {
            foreach (var m in dSpreadsheet.CustomFilePropertiesPart.Properties!.Elements<Op.CustomDocumentProperty>())
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
    }

    private void LoadCalculationProperties(CalculationProperties? calculationProperties)
    {
        if (calculationProperties is null)
            return;

        if (calculationProperties.CalculationMode is { } calculateMode)
            CalculateMode = calculateMode.Value.ToXLibur();

        if (calculationProperties.CalculationOnSave is { } calculationOnSave)
            CalculationOnSave = calculationOnSave.Value;

        if (calculationProperties.ForceFullCalculation is { } forceFullCalculation)
            ForceFullCalculation = forceFullCalculation.Value;

        if (calculationProperties.FullCalculationOnLoad is { } fullCalculationOnLoad)
            FullCalculationOnLoad = fullCalculationOnLoad.Value;

        if (calculationProperties.FullPrecision is { } fullPrecision)
            FullPrecision = fullPrecision.Value;

        if (calculationProperties.ReferenceMode is { } referenceMode)
            ReferenceStyle = referenceMode.Value.ToXLibur();
    }

    private void LoadExtendedFileProperties(SpreadsheetDocument dSpreadsheet)
    {
        var efp = dSpreadsheet.ExtendedFilePropertiesPart;
        if (efp is { Properties: not null })
        {
            if (efp.Properties.Elements<Ap.Company>().Any())
                Properties.Company = efp.Properties.GetFirstChild<Ap.Company>()!.Text;

            if (efp.Properties.Elements<Ap.Manager>().Any())
                Properties.Manager = efp.Properties.GetFirstChild<Ap.Manager>()!.Text;
        }
    }

    private void LoadSheetsPass1(WorkbookPart workbookPart, Sheets sheets)
    {
        var position = 0;
        foreach (var dSheet in sheets.OfType<Sheet>())
        {
            position++;
            var sheetName = dSheet.Name!.Value!;
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
            if (workbookPart.GetPartById(dSheet.Id.Value!) is not WorksheetPart)
            {
                UnsupportedSheets.Add(new UnsupportedSheet { SheetId = sheetIdValue, Position = position });
                continue;
            }

            var ws = WorksheetsInternal.Add(sheetName, position, sheetIdValue);
            ws.RelId = dSheet.Id;

            if (dSheet.State != null)
                ws.Visibility = dSheet.State.Value.ToXLibur();
        }
    }

    private void LoadSheetsPass2(
        WorkbookPart workbookPart,
        Sheets sheets,
        SharedStringEntry[]? sharedStrings,
        LoadContext context)
    {
        var styles = context.Styles;

        foreach (var dSheet in sheets.OfType<Sheet>())
        {
            if (string.IsNullOrEmpty(dSheet.Id))
            {
                // Some non-Excel producers create sheets with empty relId.
                continue;
            }

            // Although the relationship to worksheet is most common, there can be other types
            // than worksheet, e.g., chartSheet. Since we can't load them, add them to a list
            // of unsupported sheets and copy them when saving. See Codeplex #6932.
            if (workbookPart.GetPartById(dSheet.Id.Value!) is not WorksheetPart worksheetPart)
                continue;

            var sheetName = dSheet.Name!.Value!;
            if (!WorksheetsInternal.TryGetWorksheet(sheetName, out var ws))
            {
                // This shouldn't be possible, as all worksheets should have already been added in the loop before this loop
                continue;
            }

            StyleDecoder.ApplyStyle(ws, 0, styles);

            LoadWorksheetElements(worksheetPart, ws, sharedStrings, context);

            // Hydrate in-cell images from rich data metadata
            LoadRichValueImages(context, ws);

            ws.ConditionalFormats.ReorderAccordingToOriginalPriority();

            LoadTables(worksheetPart, ws);

            DrawingPartReader.LoadDrawings(worksheetPart, ws);

            ChartReader.LoadCharts(worksheetPart, ws);

            LoadComments(worksheetPart, ws);

            LoadThreadedComments(worksheetPart, ws);
        }
    }

    private void LoadWorksheetElements(
        WorksheetPart worksheetPart,
        XLWorksheet ws,
        SharedStringEntry[]? sharedStrings,
        LoadContext context)
    {
        var styles = context.Styles;
        var sharedFormulasR1C1 = new Dictionary<uint, string>();
        var numberDataTypeCache = new Dictionary<XLNumberFormatValue, XLDataType>();
        var sheetDataContext = new WorksheetSheetDataReader.SheetDataReadContext(
            styles, ws, sharedStrings, sharedFormulasR1C1, context.StyleCache, numberDataTypeCache,
            Use1904DateSystem, context.DynamicArrayCmIndexes);
        var sheetDataState = new WorksheetSheetDataReader.SheetDataReadState();
        var elementContext = new WorksheetElementContext
        {
            Part = worksheetPart,
            Worksheet = ws,
            Styles = styles,
            Load = context,
            Workbook = this,
        };
        var elementState = default(WorksheetElementState);

        // Pass 1: structural elements via the OpenXML SDK reader (the proven DOM path). The
        // <sheetData> hot path is skipped here — it is read in pass 2 with a raw XmlReader, which
        // is ~4x faster and allocates ~5x less than materializing every cell through the SDK
        // reader's object model. Structural elements such as <cols> are parsed here (before pass 2
        // runs), so column styles are already available when cells resolve their inherited style.
        using (var reader = new OpenXmlPartReader(worksheetPart))
        {
            while (reader.Read())
            {
                // Skipped wholesale, without descending:
                //  - CustomSheetViews carries its own auto filter data and more, ignored for now.
                //  - SheetData is read in pass 2 by the raw reader.
                // ReadNextSibling leaves the reader *on* the next sibling rather than needing
                // another Read, which is why this is a leading loop rather than a `continue`.
                while (reader.ElementType == typeof(CustomSheetViews) || reader.ElementType == typeof(SheetData))
                    reader.ReadNextSibling();

                WorksheetElementReader.TryLoad(reader, in elementContext, ref elementState);
            }
        }

        // Pass 2: read <sheetData> rows/cells directly from a raw XmlReader.
        LoadSheetDataRaw(worksheetPart, in sheetDataContext, ref sheetDataState);
    }

    /// <summary>
    /// Reads the <c>&lt;sheetData&gt;</c> rows and cells from a raw <see cref="XmlReader"/> opened
    /// over the worksheet part stream. See <see cref="WorksheetSheetDataReader.LoadSheetDataRows"/>.
    /// </summary>
    /// <remarks>
    /// Opening a second stream over the part and rescanning to <c>&lt;sheetData&gt;</c> is close to
    /// free — measured at 0.05–0.65 ms per load across sheet shapes, because everything ahead of
    /// <c>&lt;sheetData&gt;</c> is small. Collapsing the two passes into one is therefore not worth
    /// the loader rewrite it would take.
    /// </remarks>
    private static void LoadSheetDataRaw(WorksheetPart worksheetPart,
        in WorksheetSheetDataReader.SheetDataReadContext context,
        ref WorksheetSheetDataReader.SheetDataReadState state)
    {
        using var stream = worksheetPart.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = PartXmlReader.Create(stream);

        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "sheetData"
                || reader.NamespaceURI != OpenXmlConst.Main2006SsNs)
                continue;

            if (reader.IsEmptyElement)
                return;

            reader.Read(); // Move into <sheetData> (first <row> or </sheetData>).
            WorksheetSheetDataReader.LoadSheetDataRows(reader, in context, ref state);
            return;
        }
    }

    private static void LoadRichValueImages(LoadContext context, XLWorksheet ws)
    {
        // Hydrate in-cell images from rich data metadata
        if (context.RichValueImages is not null)
        {
            foreach (var cell in ws.Internals.CellsCollection.GetCells(c => c.ValueMetaIndex is not null))
            {
                if (context.RichValueImages.TryGetValue(cell.ValueMetaIndex!.Value, out var cellImage))
                {
                    cell.CellImage = cellImage;
                }
            }
        }
    }

    private static void LoadTables(WorksheetPart worksheetPart, XLWorksheet ws)
    {
        foreach (var tableDefinitionPart in worksheetPart.TableDefinitionParts)
        {
            var relId = worksheetPart.GetIdOfPart(tableDefinitionPart);
            LoadSingleTable(tableDefinitionPart.Table!, relId, ws);
        }
    }

    private static void LoadSingleTable(Table dTable, string relId, XLWorksheet ws)
    {
        var reference = dTable.Reference!.Value!;
        var tableName = dTable.Name?.Value ?? dTable.DisplayName?.Value ?? string.Empty;
        if (string.IsNullOrWhiteSpace(tableName))
            throw new InvalidDataException("The table name is missing.");

        var xlTable = (XLTable)ws.Table(ws.Range(reference)!, tableName, addToTables: true, setAutofilter: false, validateOverlap: false);
        xlTable.RelId = relId;

        if (dTable.HeaderRowCount is not null && dTable.HeaderRowCount == 0)
        {
            xlTable.HydrateShowHeaderRow(false);
            xlTable.AddFields(dTable.TableColumns!.Cast<TableColumn>()
                .Select(t => DrawingPartReader.GetTableColumnName(t.Name!.Value!)));
        }
        else
        {
            xlTable.InitializeAutoFilter();
        }

        if (dTable.TotalsRowCount is not null && dTable.TotalsRowCount.Value > 0)
            xlTable.HydrateShowTotalsRow(true);

        LoadTableStyleInfo(dTable, xlTable);
        LoadTableAutoFilter(dTable, xlTable);
        LoadTableTotalsRow(dTable, xlTable);
    }

    private static void LoadTableAutoFilter(Table dTable, XLTable xlTable)
    {
        if (dTable.AutoFilter is not null)
        {
            xlTable.ShowAutoFilter = true;
            WorksheetElementReader.LoadAutoFilterColumns(dTable.AutoFilter, xlTable.AutoFilter);
        }
        else
        {
            xlTable.ShowAutoFilter = false;
        }
    }

    private static void LoadTableTotalsRow(Table dTable, XLTable xlTable)
    {
        if (!xlTable.ShowTotalsRow)
        {
            if (xlTable.AutoFilter is not null)
                xlTable.AutoFilter.Range = xlTable.Worksheet.Range(xlTable.RangeAddress);

            return;
        }

        foreach (var tableColumn in dTable.TableColumns!.Cast<TableColumn>())
        {
            var tableColumnName = DrawingPartReader.GetTableColumnName(tableColumn.Name!.Value!);
            var field = xlTable.Field(tableColumnName);

            if (tableColumn.TotalsRowFunction is not null)
                field.TotalsRowFunction = tableColumn.TotalsRowFunction.Value.ToXLibur();

            if (tableColumn.TotalsRowFormula is not null)
                field.TotalsRowFormulaA1 = tableColumn.TotalsRowFormula.Text;

            if (tableColumn.TotalsRowLabel is not null)
                field.TotalsRowLabel = tableColumn.TotalsRowLabel.Value;
        }

        if (xlTable.AutoFilter is not null)
            xlTable.AutoFilter.Range = xlTable.Worksheet.Range(
                xlTable.RangeAddress.FirstAddress.RowNumber, xlTable.RangeAddress.FirstAddress.ColumnNumber,
                xlTable.RangeAddress.LastAddress.RowNumber - 1,
                xlTable.RangeAddress.LastAddress.ColumnNumber);
    }

    private static void LoadTableStyleInfo(Table dTable, XLTable xlTable)
    {
        if (dTable.TableStyleInfo is not { } info)
        {
            xlTable.Theme = XLTableTheme.None;
            xlTable.ShowRowStripes = false;
            xlTable.ShowColumnStripes = false;
            xlTable.EmphasizeFirstColumn = false;
            xlTable.EmphasizeLastColumn = false;
            return;
        }

        if (info.ShowFirstColumn != null)
            xlTable.EmphasizeFirstColumn = info.ShowFirstColumn.Value;
        if (info.ShowLastColumn != null)
            xlTable.EmphasizeLastColumn = info.ShowLastColumn.Value;
        if (info.ShowRowStripes != null)
            xlTable.ShowRowStripes = info.ShowRowStripes.Value;
        if (info.ShowColumnStripes != null)
            xlTable.ShowColumnStripes = info.ShowColumnStripes.Value;

        if (info.Name != null)
        {
            var theme = XLTableTheme.FromName(info.Name.Value!);
            xlTable.Theme = theme ?? new XLTableTheme(info.Name.Value!);
        }
        else
            xlTable.Theme = XLTableTheme.None;
    }

    private void LoadComments(WorksheetPart worksheetPart, XLWorksheet ws)
    {
        if (worksheetPart.WorksheetCommentsPart != null)
        {
            var root = worksheetPart.WorksheetCommentsPart.Comments!;
            var authors = root.GetFirstChild<Authors>()!.ChildElements.OfType<Author>().ToList();
            var comments = root.GetFirstChild<CommentList>()!.ChildElements.OfType<Comment>().ToList();

            // **** MAYBE FUTURE SHAPE SIZE SUPPORT
            var shapes = DrawingPartReader.GetCommentShapes(worksheetPart);

            for (var i = 0; i < comments.Count; i++)
            {
                var shape = i < shapes.Count ? shapes[i] : null;
                LoadSingleComment(comments[i], shape, authors, ws);
            }
        }
    }

    private void LoadSingleComment(Comment c, XElement? shape, List<Author> authors, XLWorksheet ws)
    {
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
            StyleDecoder.ApplyRunFont(runProperties, rt);
        }

        // Comments can have text not wrapped in a Run element (e.g., Google Sheets exports)
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

    /// <summary>
    /// Reads <c>xl/persons/person.xml</c>, the workbook-level list of identities that threaded
    /// comments are attributed to.
    /// </summary>
    private void LoadPersons(WorkbookPart workbookPart)
    {
        foreach (var personPart in workbookPart.GetPartsOfType<WorkbookPersonPart>())
        {
            var personList = personPart.PersonList;
            if (personList is null)
                continue;

            foreach (var person in personList.Elements<TC.Person>())
            {
                // A person without a parsable id cannot be referenced by a threaded comment, so
                // there is nothing meaningful to attach it to.
                if (!TryParseGuid(person.Id?.Value, out var id))
                    continue;

                PersonsInternal.AddOrGet(
                    id,
                    person.DisplayName?.Value ?? string.Empty,
                    person.UserId?.Value,
                    person.ProviderId?.Value);
            }
        }
    }

    /// <summary>
    /// Reads the threaded comments of a sheet into the cell model. Excel pairs every thread with a
    /// legacy note carrying "[Threaded comment]" boilerplate so that older versions show something;
    /// that note is taken off the cell here and kept only for its shape, because the thread is what
    /// the user model exposes.
    /// </summary>
    private void LoadThreadedComments(WorksheetPart worksheetPart, XLWorksheet ws)
    {
        foreach (var threadedComments in worksheetPart.WorksheetThreadedCommentsParts
            .Select(p => p.ThreadedComments)
            .Where(tc => tc is not null))
        {
            // Roots have no parentId; replies point at their root through it. A file may list them
            // in any order, so roots are created first and replies attached afterwards.
            var all = threadedComments!.Elements<TC.ThreadedComment>().ToList();
            var rootsById = new Dictionary<Guid, XLThreadedComment>();

            foreach (var tc in all)
            {
                if (tc.ParentId is not null)
                    continue;

                if (LoadThreadRoot(tc, ws) is { } root)
                    rootsById[root.Id] = root;
            }

            LoadThreadReplies(all, rootsById);
        }
    }

    private XLThreadedComment? LoadThreadRoot(TC.ThreadedComment tc, XLWorksheet ws)
    {
        var reference = tc.Ref?.Value;
        if (string.IsNullOrEmpty(reference) || !TryParseGuid(tc.Id?.Value, out var id))
            return null;

        var cell = ws.Cell(reference);
        if (cell is null)
            return null;

        var author = ResolvePerson(tc.PersonId?.Value);

        // The paired fallback note was loaded moments ago by LoadComments. Move it onto the thread
        // so that its position and size survive a round trip, and take it off the cell so that
        // HasComment reports what the user actually sees.
        var fallbackNote = cell.SliceComment;
        cell.SliceComment = null;

        var root = new XLThreadedComment(ws, id, author, tc.ThreadedCommentText?.InnerText ?? string.Empty,
            ParseThreadedCommentDate(tc.DT?.Value))
        {
            LegacyNote = fallbackNote,
            MentionsXml = tc.ThreadedCommentMentions?.OuterXml
        };

        if (tc.Done?.Value == true)
            root.Resolved = true;

        cell.SliceThreadedComment = root;
        return root;
    }

    private void LoadThreadReplies(List<TC.ThreadedComment> all, Dictionary<Guid, XLThreadedComment> rootsById)
    {
        // Excel writes replies in thread order, but the schema does not guarantee it, so sort by
        // timestamp to get a deterministic order regardless of producer.
        foreach (var tc in all
            .Where(tc => tc.ParentId is not null)
            .OrderBy(tc => ParseThreadedCommentDate(tc.DT?.Value)))
        {
            if (!TryParseGuid(tc.ParentId?.Value, out var parentId) ||
                !rootsById.TryGetValue(parentId, out var root))
            {
                // A reply whose root is missing has nothing to attach to. Excel does not produce
                // this, but a damaged or hand-edited file can.
                continue;
            }

            if (!TryParseGuid(tc.Id?.Value, out var id))
                continue;

            var reply = root.AddLoadedReply(
                id,
                ResolvePerson(tc.PersonId?.Value),
                tc.ThreadedCommentText?.InnerText ?? string.Empty,
                ParseThreadedCommentDate(tc.DT?.Value));

            reply.MentionsXml = tc.ThreadedCommentMentions?.OuterXml;
        }
    }

    /// <summary>
    /// Returns the person a threaded comment is attributed to, inventing a placeholder when the
    /// file references an id that <c>person.xml</c> does not define. Losing the comment over a
    /// dangling reference would be worse than showing it without a real author.
    /// </summary>
    private XLPerson ResolvePerson(string? personId)
    {
        if (TryParseGuid(personId, out var id))
        {
            if (PersonsInternal.Get(id) is XLPerson known)
                return known;

            return PersonsInternal.AddOrGet(id, string.Empty, userId: null, providerId: null);
        }

        return PersonsInternal.AddOrGet(Guid.NewGuid(), string.Empty, userId: null, providerId: null);
    }

    /// <summary>
    /// Parses the <c>{XXXXXXXX-XXXX-...}</c> braced GUIDs Excel writes for person and comment ids.
    /// </summary>
    private static bool TryParseGuid(string? value, out Guid guid)
    {
        if (!string.IsNullOrEmpty(value))
            return Guid.TryParse(value, out guid);

        guid = Guid.Empty;
        return false;
    }

    /// <summary>
    /// Excel writes the timestamp without a time zone designator and means UTC by it, so a parsed
    /// value has to be pinned to UTC rather than left as Unspecified or shifted by the local zone.
    /// </summary>
    private static DateTime ParseThreadedCommentDate(DateTime? value)
    {
        // Kind matters even with no value to convert: CreatedUtc promises UTC, and a bare default
        // would hand back an Unspecified DateTime that silently shifts if anyone converts it.
        if (value is not { } dt)
            return DateTime.SpecifyKind(default, DateTimeKind.Utc);

        return dt.Kind switch
        {
            DateTimeKind.Utc => dt,
            DateTimeKind.Local => dt.ToUniversalTime(),
            _ => DateTime.SpecifyKind(dt, DateTimeKind.Utc)
        };
    }

    private void LoadActiveTab(Workbook workbook)
    {
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
    }

    private void LoadPivotTables(
        WorkbookPart workbookPart,
        Sheets sheets,
        LoadContext context)
    {
        // Delay loading of pivot tables until all sheets have been loaded
        foreach (var dSheet in sheets.OfType<Sheet>())
        {
            if (string.IsNullOrEmpty(dSheet.Id))
            {
                // Some non-Excel producers create sheets with empty relId.
                continue;
            }

            // The referenced sheet can also be ChartsheetPart. Only look for pivot tables in normal sheet parts.
            if (workbookPart.GetPartById(dSheet.Id.Value!) is WorksheetPart worksheetPart)
            {
                var ws = (XLWorksheet)WorksheetsInternal.Worksheet(dSheet.Name!.Value!);

                foreach (var pivotTablePart in worksheetPart.PivotTableParts)
                {
                    PivotTableDefinitionPartReader.Load(workbookPart, context.Styles.DifferentialFormats, pivotTablePart,
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
        var mdw = workbook.FontEngine.GetMaxDigitWidth(font, workbook.DpiX).RoundToInt();
        var defaultColWidthPx = XLHelper.NoCToPixels(charWidth, mdw).RoundToInt();

        // Excel then rounds this number up to the nearest multiple of 8 pixels so that
        // scrolling across columns and rows is faster.
        var roundUpToMultiple = defaultColWidthPx + (8 - defaultColWidthPx % 8);

        // and last, convert the width in pixels to width displayed in Excel. Shouldn't round the number, because
        // it causes inconsistency with conversion to other units, but other places in XLibur do = keep for now.
        var defaultColumnWidth = XLHelper.PixelToNoC(roundUpToMultiple, mdw).Round(2);
        return defaultColumnWidth;
    }

    /// <summary>
    /// Reads the twelve theme colours out of <c>&lt;a:clrScheme&gt;</c>.
    /// </summary>
    /// <remarks>
    /// Deliberately a raw <see cref="XmlReader"/> rather than <c>tp.Theme</c>. Touching the DOM
    /// property materialises the whole theme part — font scheme, format scheme, gradient fills,
    /// effect styles, and often an <c>extraClrSchemeLst</c> — which for a stock Excel theme is
    /// ~7 KB of dense XML built into an object graph, all to read twelve hex strings. That is a
    /// fixed cost paid on every workbook load regardless of its size, and it measured at ~8% of
    /// total load time for a small workbook.
    /// <para>
    /// Only <c>&lt;a:srgbClr&gt;</c> is honoured, matching the DOM version this replaced: a slot
    /// carrying <c>&lt;a:sysClr&gt;</c> leaves the corresponding <see cref="XLWorkbook.Theme"/>
    /// property at its default.
    /// </para>
    /// </remarks>
    private static void LoadWorkbookTheme(ThemePart? tp, XLWorkbook wb)
    {
        if (tp is null) return;

        using var stream = tp.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = PartXmlReader.Create(stream);

        // <a:theme> (0) / <a:themeElements> (1) / <a:clrScheme> (2). Matching on depth as well as
        // name is what keeps the scan off the <a:clrScheme> nested inside <a:extraClrSchemeLst>,
        // which lives a level deeper and would otherwise overwrite the real scheme.
        if (!MoveToThemeElement(reader, "theme", depth: 0) ||
            !MoveToThemeElement(reader, "themeElements", depth: 1) ||
            !MoveToThemeElement(reader, "clrScheme", depth: 2) ||
            reader.IsEmptyElement)
            return;

        var schemeDepth = reader.Depth;
        var theme = wb.Theme;

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == schemeDepth)
                break;

            if (reader.NodeType != XmlNodeType.Element ||
                reader.NamespaceURI != OpenXmlConst.DrawingMain2006Ns)
                continue;

            // ReadSlotColor consumes the slot's subtree, so the next Read lands on the next slot.
            var slot = reader.LocalName;
            var hex = ReadSlotColor(reader);
            if (string.IsNullOrEmpty(hex))
                continue;

            // A theme is decoration, so a malformed slot should cost that one colour rather than
            // the whole workbook. FromHexRgb throws FormatException both for a value that is not
            // six characters and for one containing a non-hex character.
            XLColor color;
            try
            {
                color = XLColor.FromHexRgb(hex);
            }
            catch (FormatException)
            {
                continue;
            }

            ApplyThemeSlot(theme, slot, color);
        }
    }

    /// <summary>
    /// Assigns a parsed colour to the theme property its <c>clrScheme</c> slot names. A slot this
    /// table does not list is ignored, leaving that theme property at its default.
    /// </summary>
    private static void ApplyThemeSlot(IXLTheme theme, string slot, XLColor color)
    {
        switch (slot)
        {
            case "lt1": theme.Background1 = color; break;
            case "dk1": theme.Text1 = color; break;
            case "lt2": theme.Background2 = color; break;
            case "dk2": theme.Text2 = color; break;
            case "accent1": theme.Accent1 = color; break;
            case "accent2": theme.Accent2 = color; break;
            case "accent3": theme.Accent3 = color; break;
            case "accent4": theme.Accent4 = color; break;
            case "accent5": theme.Accent5 = color; break;
            case "accent6": theme.Accent6 = color; break;
            case "hlink": theme.Hyperlink = color; break;
            case "folHlink": theme.FollowedHyperlink = color; break;
        }
    }

    /// <summary>
    /// Advances to the next DrawingML element with the given name at the given depth, giving up
    /// once the reader leaves the current parent.
    /// </summary>
    private static bool MoveToThemeElement(XmlReader reader, string localName, int depth)
    {
        while (reader.Read())
        {
            if (reader.Depth < depth)
                return false;

            if (reader.NodeType == XmlNodeType.Element &&
                reader.Depth == depth &&
                reader.LocalName == localName &&
                reader.NamespaceURI == OpenXmlConst.DrawingMain2006Ns)
                return true;
        }

        return false;
    }

    /// <summary>
    /// Reads the <c>val</c> of the first <c>&lt;a:srgbClr&gt;</c> inside the colour slot the
    /// reader is positioned on, consuming the slot's subtree.
    /// </summary>
    private static string? ReadSlotColor(XmlReader reader)
    {
        if (reader.IsEmptyElement)
            return null;

        var slotDepth = reader.Depth;
        string? hex = null;

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == slotDepth)
                break;

            if (hex is null &&
                reader.NodeType == XmlNodeType.Element &&
                reader.LocalName == "srgbClr" &&
                reader.NamespaceURI == OpenXmlConst.DrawingMain2006Ns)
                hex = reader.GetAttribute("val");
        }

        return hex;
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
