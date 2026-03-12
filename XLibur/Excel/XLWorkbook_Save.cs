using XLibur.Extensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Path = System.IO.Path;
using XLibur.Excel.IO;
using System.Diagnostics;

namespace XLibur.Excel;

public partial class XLWorkbook
{
    private static void Validate(SpreadsheetDocument package)
    {
        var backupCulture = Thread.CurrentThread.CurrentCulture;

        IList<ValidationErrorInfo> errors;
        try
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            var validator = new OpenXmlValidator();
            errors = validator.Validate(package).ToArray();
        }
        finally
        {
            Thread.CurrentThread.CurrentCulture = backupCulture;
        }

        if (!errors.Any()) return;
        var message = string.Join("\r\n", errors.Select(e => $"Part {e.Part?.Uri}, Path {e.Path?.XPath}: {e.Description}").ToArray());
        throw new ApplicationException(message);
    }

    private void CreatePackage(string filePath, SpreadsheetDocumentType spreadsheetDocumentType, SaveOptions options)
    {
        var directoryName = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrWhiteSpace(directoryName)) Directory.CreateDirectory(directoryName);

        var package = File.Exists(filePath)
            ? SpreadsheetDocument.Open(filePath, true)
            : SpreadsheetDocument.Create(filePath, spreadsheetDocumentType);

        using (package)
        {
            if (package.DocumentType != spreadsheetDocumentType)
            {
                package.ChangeDocumentType(spreadsheetDocumentType);
            }

            CreateParts(package, options);
            if (options.ValidatePackage) Validate(package);
        }
    }

    private void CreatePackage(Stream stream, bool newStream, SpreadsheetDocumentType spreadsheetDocumentType, SaveOptions options)
    {
        var package = newStream
            ? SpreadsheetDocument.Create(stream, spreadsheetDocumentType)
            : SpreadsheetDocument.Open(stream, true);

        using (package)
        {
            CreateParts(package, options);
            if (options.ValidatePackage) Validate(package);
        }
    }

    // http://blogs.msdn.com/b/vsod/archive/2010/02/05/how-to-delete-a-worksheet-from-excel-using-open-xml-sdk-2-0.aspx
    private static void DeleteSheetAndDependencies(WorkbookPart wbPart, string sheetId)
    {
        var sheet = wbPart.Workbook!.Descendants<Sheet>().FirstOrDefault(s => s.Id == sheetId);
        if (sheet == null)
            return;

        string sheetName = sheet.Name!;

        DeleteLinkedPivotTableCaches(wbPart, sheetName);
        DeleteWorksheetPart(wbPart, sheet, sheetId);
        DeleteDefinedNamesForSheet(wbPart, sheetName);
        DeleteCalculationChainEntries(wbPart, sheetId);
    }

    private static void DeleteLinkedPivotTableCaches(WorkbookPart wbPart, string sheetName)
    {
        var partsToDelete = new List<PivotTableCacheDefinitionPart>();
        foreach (var part in wbPart.PivotTableCacheDefinitionParts)
        {
            var cacheSource = part.PivotCacheDefinition?.Descendants<CacheSource>()
                .Any(cs => cs.WorksheetSource?.Sheet == sheetName);
            if (cacheSource == true)
                partsToDelete.Add(part);
        }

        foreach (var part in partsToDelete)
            wbPart.DeletePart(part);
    }

    private static void DeleteWorksheetPart(WorkbookPart wbPart, Sheet sheet, string sheetId)
    {
        var worksheetPart = (WorksheetPart)wbPart.GetPartById(sheetId);
        sheet.Remove();
        wbPart.DeletePart(worksheetPart);
    }

    private static void DeleteDefinedNamesForSheet(WorkbookPart wbPart, string sheetName)
    {
        var definedNames = wbPart.Workbook!.Descendants<DefinedNames>().FirstOrDefault();
        if (definedNames == null)
            return;

        var toDelete = definedNames.OfType<DefinedName>()
            .Where(dn => dn.Text.Contains(sheetName + "!"))
            .ToList();

        foreach (var item in toDelete)
            item.Remove();
    }

    private static void DeleteCalculationChainEntries(WorkbookPart wbPart, string sheetId)
    {
        var calChainPart = wbPart.CalculationChainPart;
        if (calChainPart == null)
            return;

        var toDelete = calChainPart.CalculationChain!
            .Descendants<CalculationCell>()
            .Where(c => c.SheetId == sheetId)
            .ToList();

        foreach (var item in toDelete)
            item.Remove();

        if (!calChainPart.CalculationChain!.Any())
            wbPart.DeletePart(calChainPart);
    }

    // Adds child parts and generates content of the specified part.
    private void CreateParts(SpreadsheetDocument document, SaveOptions options)
    {
        var context = new SaveContext();
        var workbookPart = document.WorkbookPart ?? document.AddWorkbookPart();
        var worksheets = WorksheetsInternal;

        DeleteRemovedWorksheets(workbookPart, worksheets);

        context.RelIdGenerator.AddExistingValues(workbookPart, this);

        GenerateWorkbookLevelParts(document, workbookPart, options, context);
        PreparePivotCaches(workbookPart, context);

        foreach (var worksheet in WorksheetsInternal.Cast<XLWorksheet>().OrderBy(w => w.Position))
        {
            var (worksheetPart, partIsEmpty) = GetOrCreateWorksheetPart(workbookPart, worksheet);
            GenerateCommentsAndVmlParts(worksheetPart, worksheet, context);
            GenerateTableParts(worksheetPart, worksheet, context);
            WorksheetPartWriter.GenerateWorksheetPartContent(partIsEmpty, worksheetPart, worksheet, options, context);

            if (worksheet.PivotTables.Any<XLPivotTable>())
                GeneratePivotTables(workbookPart, worksheetPart, worksheet, context);
        }

        GenerateSupplementaryParts(document, workbookPart, options, context);

        // Clear list of deleted worksheets to prevent errors on multiple saves
        worksheets.Deleted.Clear();
    }

    private static void DeleteRemovedWorksheets(WorkbookPart workbookPart, XLWorksheets worksheets)
    {
        var partsToRemove = workbookPart.Parts.Where(s => worksheets.Deleted.Contains(s.RelationshipId)).ToList();

        var pivotCacheDefinitionsToRemove = partsToRemove
            .SelectMany(s => ((WorksheetPart)s.OpenXmlPart).PivotTableParts.Select(pt => pt.PivotTableCacheDefinitionPart))
            .Where(c => c is not null)
            .Select(c => c!)
            .Distinct()
            .ToList();
        pivotCacheDefinitionsToRemove.ForEach(c => workbookPart.DeletePart(c));

        if (workbookPart.Workbook is { PivotCaches: not null })
        {
            var idList = pivotCacheDefinitionsToRemove.Select(workbookPart.GetIdOfPart).ToList();
            var pivotCachesToRemove = workbookPart.Workbook.PivotCaches
                .Where(pc => ((PivotCache)pc).Id?.Value is { } idVal && idList.Contains(idVal))
                .Distinct()
                .ToList();
            pivotCachesToRemove.ForEach(c => workbookPart.Workbook.PivotCaches.RemoveChild(c));
        }

        worksheets.Deleted.ToList().ForEach(ws => DeleteSheetAndDependencies(workbookPart, ws));
    }

    private void GenerateWorkbookLevelParts(SpreadsheetDocument document, WorkbookPart workbookPart, SaveOptions options, SaveContext context)
    {
        var extendedFilePropertiesPart = document.ExtendedFilePropertiesPart ??
                                         document.AddNewPart<ExtendedFilePropertiesPart>(
                                             context.RelIdGenerator.GetNext(RelType.Workbook));
        ExtendedFilePropertiesPartWriter.GenerateContent(extendedFilePropertiesPart, this);

        WorkbookPartWriter.GenerateContent(workbookPart, this, options, context);

        var sharedStringTablePart = workbookPart.SharedStringTablePart ??
                                    workbookPart.AddNewPart<SharedStringTablePart>(
                                        context.RelIdGenerator.GetNext(RelType.Workbook));
        SharedStringTableWriter.GenerateSharedStringTablePartContent(this, sharedStringTablePart, context);

        var workbookStylesPart = workbookPart.WorkbookStylesPart ??
                                 workbookPart.AddNewPart<WorkbookStylesPart>(
                                     context.RelIdGenerator.GetNext(RelType.Workbook));
        WorkbookStylesPartWriter.GenerateContent(workbookStylesPart, this, context);
    }

    private void PreparePivotCaches(WorkbookPart workbookPart, SaveContext context)
    {
        var cacheRelIds = PivotCachesInternal
            .Select<XLPivotCache, string?>(ps => ps.WorkbookCacheRelId)
            .Where(relId => !string.IsNullOrWhiteSpace(relId))
            .Select(relId => relId!)
            .Distinct();

        foreach (var relId in cacheRelIds)
        {
            if (workbookPart.GetPartById(relId) is PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart)
                pivotTableCacheDefinitionPart.PivotCacheDefinition!.CacheFields!.RemoveAllChildren();
        }

        var allPivotTables = WorksheetsInternal.SelectMany<XLWorksheet, IXLPivotTable>(ws => ws.PivotTables).ToList();

        SynchronizePivotTableParts(workbookPart, allPivotTables, context);

        if (allPivotTables.Count != 0)
            GeneratePivotCaches(workbookPart, context);
    }

    private static (WorksheetPart worksheetPart, bool partIsEmpty) GetOrCreateWorksheetPart(WorkbookPart workbookPart, XLWorksheet worksheet)
    {
        var wsRelId = worksheet.RelId;
        if (workbookPart.Parts.Any(p => p.RelationshipId == wsRelId))
            return ((WorksheetPart)workbookPart.GetPartById(wsRelId!), false);

        return (workbookPart.AddNewPart<WorksheetPart>(wsRelId!), true);
    }

    private static void GenerateCommentsAndVmlParts(WorksheetPart worksheetPart, XLWorksheet worksheet, SaveContext context)
    {
        var worksheetHasComments = worksheet.Internals.CellsCollection.GetCells(c => c.HasComment).Any();

        // VML part is the source of truth for shapes of comments, form controls and likely others.
        // Excel won't display any shape without VML. The drawing part is always present but is likely
        // only a different rendering of VML (more precisely the shapes behind VML).
        var vmlDrawingPart = worksheetPart.VmlDrawingParts.FirstOrDefault();
        var hasAnyVmlElements = DeleteExistingCommentsShapes(vmlDrawingPart);

        if (worksheetHasComments)
            hasAnyVmlElements = EnsureCommentAndVmlParts(worksheetPart, worksheet, context, ref vmlDrawingPart);
        else
            RemoveCommentsPartIfPresent(worksheetPart);

        RemoveEmptyVmlPart(worksheetPart, worksheet, vmlDrawingPart, hasAnyVmlElements);
    }

    private static bool EnsureCommentAndVmlParts(WorksheetPart worksheetPart, XLWorksheet worksheet, SaveContext context, ref VmlDrawingPart? vmlDrawingPart)
    {
        var commentsPart = worksheetPart.WorksheetCommentsPart
                           ?? worksheetPart.AddNewPart<WorksheetCommentsPart>(context.RelIdGenerator.GetNext(RelType.Workbook));

        if (vmlDrawingPart == null)
        {
            if (string.IsNullOrWhiteSpace(worksheet.LegacyDrawingId))
                worksheet.LegacyDrawingId = context.RelIdGenerator.GetNext(RelType.Workbook);

            vmlDrawingPart = worksheetPart.AddNewPart<VmlDrawingPart>(worksheet.LegacyDrawingId);
        }

        CommentPartWriter.GenerateWorksheetCommentsPartContent(commentsPart, worksheet);
        return VmlDrawingPartWriter.GenerateContent(vmlDrawingPart, worksheet);
    }

    private static void RemoveCommentsPartIfPresent(WorksheetPart worksheetPart)
    {
        if (worksheetPart.WorksheetCommentsPart is { } commentsPart)
            worksheetPart.DeletePart(commentsPart);
    }

    private static void RemoveEmptyVmlPart(WorksheetPart worksheetPart, XLWorksheet worksheet, VmlDrawingPart? vmlDrawingPart, bool hasAnyVmlElements)
    {
        if (!hasAnyVmlElements && vmlDrawingPart is not null)
        {
            worksheet.LegacyDrawingId = null;
            worksheetPart.DeletePart(vmlDrawingPart);
        }
    }

    private static void GenerateTableParts(WorksheetPart worksheetPart, XLWorksheet worksheet, SaveContext context)
    {
        var xlTables = worksheet.Tables;

        // Phase 1 - synchronize part existence with tables, so each
        // table has a corresponding part and parts that don't are deleted.
        TablePartWriter.SynchronizeTableParts(xlTables, worksheetPart, context);

        // Phase 2 - all pieces have corresponding parts, fill in content.
        TablePartWriter.GenerateTableParts(xlTables, worksheetPart, context);
    }

    private void GenerateSupplementaryParts(SpreadsheetDocument document, WorkbookPart workbookPart, SaveOptions options, SaveContext context)
    {
        if (options.GenerateCalculationChain)
        {
            CalculationChainPartWriter.GenerateContent(workbookPart, this, context);
        }
        else
        {
            if (workbookPart.CalculationChainPart is not null)
                workbookPart.DeletePart(workbookPart.CalculationChainPart);
        }

        if (workbookPart.ThemePart == null)
        {
            var themePart = workbookPart.AddNewPart<ThemePart>(context.RelIdGenerator.GetNext(RelType.Workbook));
            ThemePartWriter.GenerateContent(themePart, (XLTheme)Theme);
        }

        if (CustomProperties.Any())
        {
            var customFilePropertiesPart =
                document.CustomFilePropertiesPart ?? document.AddNewPart<CustomFilePropertiesPart>(context.RelIdGenerator.GetNext(RelType.Workbook));
            CustomFilePropertiesPartWriter.GenerateContent(customFilePropertiesPart, this);
        }
        else
        {
            if (document.CustomFilePropertiesPart != null)
                document.DeletePart(document.CustomFilePropertiesPart);
        }

        SetPackageProperties(document);
    }

    private static bool DeleteExistingCommentsShapes(VmlDrawingPart? vmlDrawingPart)
    {
        if (vmlDrawingPart == null)
            return false;

        // Nuke the VmlDrawingPart elements for comments.
        using var vmlStream = vmlDrawingPart.GetStream(FileMode.Open);
        var xdoc = XDocumentExtensions.Load(vmlStream);
        if (xdoc == null)
            return false;

        // Remove existing shapes for comments
        xdoc.Root!
            .Elements()
            .Where(e => e.Name.LocalName == "shapetype" && e.Attribute("id")?.Value == XLConstants.Comment.ShapeTypeId)
            .Remove();

        xdoc.Root!
            .Elements()
            .Where(e => e.Name.LocalName == "shape" && e.Attribute("type")?.Value == "#" + XLConstants.Comment.ShapeTypeId)
            .Remove();

        vmlStream.Position = 0;

        using (var writer = new XmlTextWriter(vmlStream, Encoding.UTF8))
        {
            var contents = xdoc.ToString();
            writer.WriteRaw(contents);
            vmlStream.SetLength(contents.Length);
        }

        return xdoc.Root.HasElements;
    }

    private void SetPackageProperties(OpenXmlPackage document)
    {
        var created = Properties.Created == DateTime.MinValue ? DateTime.Now : Properties.Created;
        var modified = Properties.Modified == DateTime.MinValue ? DateTime.Now : Properties.Modified;
        document.PackageProperties.Created = created;
        document.PackageProperties.Modified = modified;

        if (Properties.LastModifiedBy == null) document.PackageProperties.LastModifiedBy = "";
        if (Properties.Author == null) document.PackageProperties.Creator = "";
        if (Properties.Title == null) document.PackageProperties.Title = "";
        if (Properties.Subject == null) document.PackageProperties.Subject = "";
        if (Properties.Category == null) document.PackageProperties.Category = "";
        if (Properties.Keywords == null) document.PackageProperties.Keywords = "";
        if (Properties.Comments == null) document.PackageProperties.Description = "";
        if (Properties.Status == null) document.PackageProperties.ContentStatus = "";

        document.PackageProperties.LastModifiedBy = Properties.LastModifiedBy;
        document.PackageProperties.Creator = Properties.Author;
        document.PackageProperties.Title = Properties.Title;
        document.PackageProperties.Subject = Properties.Subject;
        document.PackageProperties.Category = Properties.Category;
        document.PackageProperties.Keywords = Properties.Keywords;
        document.PackageProperties.Description = Properties.Comments;
        document.PackageProperties.ContentStatus = Properties.Status;
    }

    private static void SynchronizePivotTableParts(WorkbookPart workbookPart, IReadOnlyList<IXLPivotTable> allPivotTables, SaveContext context)
    {
        RemoveUnusedPivotCacheDefinitionParts(workbookPart, allPivotTables);
        AddUsedPivotCacheDefinitionParts(workbookPart, allPivotTables, context);
        SynchronizeWorkbookPivotCacheReferences(workbookPart, allPivotTables, context);
    }

    /// <summary>
    /// Remove pivot cache parts that are in the loaded document but aren't used by any pivot table.
    /// </summary>
    private static void RemoveUnusedPivotCacheDefinitionParts(WorkbookPart workbookPart, IReadOnlyList<IXLPivotTable> allPivotTables)
    {
        var workbookCacheRelIds = allPivotTables
            .Select(pt => pt.PivotCache.CastTo<XLPivotCache>().WorkbookCacheRelId)
            .Distinct()
            .ToList();

        var orphanedParts = workbookPart
            .GetPartsOfType<PivotTableCacheDefinitionPart>()
            .Where(pcdp => !workbookCacheRelIds.Contains(workbookPart.GetIdOfPart(pcdp)))
            .ToList();

        foreach (var orphanPart in orphanedParts)
        {
            orphanPart.DeletePart(orphanPart.PivotTableCacheRecordsPart!);
            workbookPart.DeletePart(orphanPart);
        }

        if (workbookPart.Workbook!.PivotCaches is not null)
        {
            workbookPart.Workbook.PivotCaches.Elements<PivotCache>()
                .Where(pc => pc.Id is null || !workbookPart.HasPartWithId(pc.Id!.Value!))
                .ToList()
                .ForEach(pc => pc.Remove());
        }
    }

    /// <summary>
    /// Add cache definition parts for pivot caches that don't yet have a corresponding part in the workbook.
    /// </summary>
    private static void AddUsedPivotCacheDefinitionParts(WorkbookPart workbookPart, IReadOnlyList<IXLPivotTable> allPivotTables, SaveContext context)
    {
        var newPivotSources = allPivotTables
            .Select(pt => pt.PivotCache.CastTo<XLPivotCache>())
            .Where(ps => string.IsNullOrEmpty(ps.WorkbookCacheRelId) || !workbookPart.HasPartWithId(ps.WorkbookCacheRelId))
            .Distinct()
            .ToList();

        foreach (var pivotSource in newPivotSources)
        {
            var cacheRelId = context.RelIdGenerator.GetNext(RelType.Workbook);
            pivotSource.WorkbookCacheRelId = cacheRelId;
            workbookPart.AddNewPart<PivotTableCacheDefinitionPart>(pivotSource.WorkbookCacheRelId!);
        }
    }

    /// <summary>
    /// Rebuild the <c>&lt;pivotCaches&gt;</c> element in workbook.xml to match the pivot tables being saved.
    /// </summary>
    private static void SynchronizeWorkbookPivotCacheReferences(WorkbookPart workbookPart, IReadOnlyList<IXLPivotTable> allPivotTables, SaveContext context)
    {
        context.PivotSourceCacheId = 0;
        var xlUsedCaches = allPivotTables.Select(pt => pt.PivotCache).Distinct().Cast<XLPivotCache>().ToList();

        if (xlUsedCaches.Count != 0)
        {
            var pivotCaches = new PivotCaches();
            workbookPart.Workbook!.PivotCaches = pivotCaches;

            foreach (var source in xlUsedCaches)
            {
                var cacheId = context.PivotSourceCacheId++;
                source.CacheId = cacheId;
                pivotCaches.AppendChild(new PivotCache { CacheId = cacheId, Id = source.WorkbookCacheRelId });
            }
        }
        else if (workbookPart.Workbook!.PivotCaches is not null)
        {
            workbookPart.Workbook.RemoveChild(workbookPart.Workbook.PivotCaches);
        }
    }

    private void GeneratePivotCaches(WorkbookPart workbookPart, SaveContext context)
    {
        var pivotTables = WorksheetsInternal.SelectMany<XLWorksheet, XLPivotTable>(ws => ws.PivotTables);

        var xlPivotCaches = pivotTables.Select(pt => pt.PivotCache).Distinct();
        foreach (var xlPivotCache in xlPivotCaches)
        {
            Debug.Assert(workbookPart.Workbook!.PivotCaches is not null);
            Debug.Assert(!string.IsNullOrEmpty(xlPivotCache.WorkbookCacheRelId));

            var pivotTableCacheDefinitionPart = (PivotTableCacheDefinitionPart)workbookPart.GetPartById(xlPivotCache.WorkbookCacheRelId!);

            PivotTableCacheDefinitionPartWriter.GenerateContent(pivotTableCacheDefinitionPart, xlPivotCache, context);

            var pivotTableCacheRecordsPart = pivotTableCacheDefinitionPart.GetPartsOfType<PivotTableCacheRecordsPart>().Any()
                ? pivotTableCacheDefinitionPart.GetPartsOfType<PivotTableCacheRecordsPart>().Single()
                : pivotTableCacheDefinitionPart.AddNewPart<PivotTableCacheRecordsPart>("rId1");

            PivotTableCacheRecordsPartWriter.WriteContent(pivotTableCacheRecordsPart, xlPivotCache);
        }
    }

    private static void GeneratePivotTables(
        WorkbookPart workbookPart,
        WorksheetPart worksheetPart,
        XLWorksheet xlWorksheet,
        SaveContext context)
    {
        foreach (var pt in xlWorksheet.PivotTables)
        {
            PivotTablePart pivotTablePart;
            var createNewPivotTablePart = string.IsNullOrWhiteSpace(pt.RelId);
            if (createNewPivotTablePart)
            {
                var relId = context.RelIdGenerator.GetNext(RelType.Workbook);
                pt.RelId = relId;
                pivotTablePart = worksheetPart.AddNewPart<PivotTablePart>(relId);
            }
            else
                pivotTablePart = (PivotTablePart)worksheetPart.GetPartById(pt.RelId!);

            var pivotSource = pt.PivotCache;
            var pivotTableCacheDefinitionPart = pivotTablePart.PivotTableCacheDefinitionPart;
            if (!workbookPart.GetPartById(pivotSource.WorkbookCacheRelId!).Equals(pivotTableCacheDefinitionPart))
            {
                pivotTablePart.DeletePart(pivotTableCacheDefinitionPart!);
                pivotTablePart.CreateRelationshipToPart(workbookPart.GetPartById(pivotSource.WorkbookCacheRelId!), context.RelIdGenerator.GetNext(RelType.Workbook));
            }

            PivotTableDefinitionPartWriter2.WriteContent(pivotTablePart, pt, context);
        }
    }
}
