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
        EnsureDynamicArrayMetadata(workbookPart, context);
        EnsureRichValueImageParts(workbookPart, context);

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
        // Collect relationship IDs before deleting parts, because GetIdOfPart
        // throws after the part has been removed.
        var pivotCacheRelIds = workbookPart.Workbook is { PivotCaches: not null }
            ? pivotCacheDefinitionsToRemove.Select(workbookPart.GetIdOfPart).ToHashSet()
            : null;

        pivotCacheDefinitionsToRemove.ForEach(c => workbookPart.DeletePart(c));

        if (pivotCacheRelIds is not null)
        {
            var idList = pivotCacheRelIds;
            var pivotCachesToRemove = workbookPart.Workbook!.PivotCaches!
                .Where(pc => ((PivotCache)pc).Id?.Value is { } idVal && idList.Contains(idVal))
                .Distinct()
                .ToList();
            pivotCachesToRemove.ForEach(c => workbookPart.Workbook.PivotCaches!.RemoveChild(c));
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

    /// <summary>
    /// If any cell in the workbook has a dynamic array formula, ensure that the
    /// <c>CellMetadataPart</c> (metadata.xml) contains the <c>XLDAPR</c> future
    /// metadata entry required by Excel 365+ to suppress the implicit intersection
    /// <c>@</c> operator. Sets <see cref="SaveContext.DynamicArrayMetaIndex"/> so
    /// the sheet writer can emit the <c>cm</c> attribute on those cells.
    /// </summary>
    private void EnsureDynamicArrayMetadata(WorkbookPart workbookPart, SaveContext context)
    {
        var hasDynamicArray = WorksheetsInternal
            .Cast<XLWorksheet>()
            .Any(ws => ws.Internals.CellsCollection
                .GetCells(c => c.HasFormula && c.Formula!.IsDynamicArray)
                .Any());

        if (!hasDynamicArray)
            return;

        // The XLDAPR metadata structure in metadata.xml consists of three parts:
        // 1. metadataTypes - declares the "XLDAPR" type with its capabilities
        // 2. futureMetadata - contains the dynamic array properties extension
        // 3. cellMetadata - one record referencing the type, used by cm attribute on cells

        var cellMetadataPart = workbookPart.CellMetadataPart;
        if (cellMetadataPart?.Metadata is not null)
        {
            SetDynamicArrayMetaFromExisting(cellMetadataPart.Metadata, context);
            return;
        }

        // No metadata part at all - create from scratch
        cellMetadataPart ??= workbookPart.AddNewPart<CellMetadataPart>(
            context.RelIdGenerator.GetNext(RelType.Workbook));

        var newMetadata = CreateXldaprMetadata();
        cellMetadataPart.Metadata = newMetadata;
        context.DynamicArrayMetaIndex = 1;
    }

    private static void SetDynamicArrayMetaFromExisting(Metadata metadata, SaveContext context)
    {
        var metadataTypes = metadata.MetadataTypes;
        if (metadataTypes is not null)
        {
            var xldaprIndex = FindXldaprTypeIndex(metadataTypes);
            if (xldaprIndex is not null)
            {
                EnsureXldaprCellMetadata(metadata, xldaprIndex.Value, context);
                return;
            }
        }

        // XLDAPR type doesn't exist - add everything
        AppendXldaprToMetadata(metadata);
        context.DynamicArrayMetaIndex = metadata.GetFirstChild<CellMetadata>()!.Count!.Value;
    }

    private static uint? FindXldaprTypeIndex(MetadataTypes metadataTypes)
    {
        uint typeIndex = 1;
        foreach (var mt in metadataTypes.Elements<MetadataType>())
        {
            if (mt.Name?.Value == "XLDAPR")
                return typeIndex;
            typeIndex++;
        }
        return null;
    }

    private static void EnsureXldaprCellMetadata(Metadata metadata, uint typeIndex, SaveContext context)
    {
        var cellMeta = metadata.GetFirstChild<CellMetadata>();
        if (cellMeta is not null)
        {
            uint cmIndex = 1;
            foreach (var bk in cellMeta.Elements<MetadataBlock>())
            {
                var rc = bk.GetFirstChild<MetadataRecord>();
                if (rc?.TypeIndex?.Value == typeIndex)
                {
                    context.DynamicArrayMetaIndex = cmIndex;
                    return;
                }
                cmIndex++;
            }
        }

        // Type exists but no cellMetadata record for it - add one
        if (cellMeta is null)
        {
            cellMeta = new CellMetadata { Count = 0 };
            metadata.Append(cellMeta);
        }

        var newBlock = new MetadataBlock();
        newBlock.Append(new MetadataRecord { TypeIndex = typeIndex, Val = 0 });
        cellMeta.Append(newBlock);
        cellMeta.Count = (uint)(cellMeta.Count ?? 0) + 1;
        context.DynamicArrayMetaIndex = cellMeta.Count.Value;
    }

    /// <summary>
    /// Create a new <see cref="Metadata"/> element with the XLDAPR dynamic array support.
    /// </summary>
    private static Metadata CreateXldaprMetadata()
    {
        var metadata = new Metadata();
        AppendXldaprToMetadata(metadata);
        return metadata;
    }

    /// <summary>
    /// Append the XLDAPR metadata type, future metadata, and cell metadata record
    /// to an existing <see cref="Metadata"/> element.
    /// </summary>
    private static void AppendXldaprToMetadata(Metadata metadata)
    {
        // 1. Add MetadataType
        var metadataTypes = metadata.MetadataTypes;
        if (metadataTypes is null)
        {
            metadataTypes = new MetadataTypes { Count = 0 };
            metadata.Append(metadataTypes);
        }

        metadataTypes.Append(new MetadataType
        {
            Name = "XLDAPR",
            MinSupportedVersion = 120000,
            Copy = true,
            PasteAll = true,
            PasteValues = true,
            Merge = true,
            SplitFirst = true,
            RowColumnShift = true,
            ClearFormats = true,
            ClearComments = true,
            Assign = true,
            Coerce = true,
            CellMeta = true
        });
        metadataTypes.Count = (uint)(metadataTypes.Count ?? 0) + 1;

        var typeIndex = metadataTypes.Count.Value; // 1-based index of the just-added type

        // 2. Add FutureMetadata with dynamic array properties extension
        var futureMetadata = new FutureMetadata { Name = "XLDAPR", Count = 1 };
        var fmBlock = new FutureMetadataBlock();
        var extList = new ExtensionList();
        var ext = new Extension { Uri = "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}" };

        // The xda:dynamicArrayProperties element must be in the correct namespace
        const string xdaNs = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray";
        var dynArrayProps = new OpenXmlUnknownElement("xda", "dynamicArrayProperties", xdaNs);
        dynArrayProps.SetAttribute(new OpenXmlAttribute("", "fDynamic", "", "1"));
        dynArrayProps.SetAttribute(new OpenXmlAttribute("", "fCollapsed", "", "0"));
        ext.Append(dynArrayProps);
        extList.Append(ext);
        fmBlock.Append(extList);
        futureMetadata.Append(fmBlock);
        metadata.Append(futureMetadata);

        // 3. Add CellMetadata record referencing the XLDAPR type
        var cellMeta = metadata.GetFirstChild<CellMetadata>();
        if (cellMeta is null)
        {
            cellMeta = new CellMetadata { Count = 0 };
            metadata.Append(cellMeta);
        }

        var block = new MetadataBlock();
        block.Append(new MetadataRecord { TypeIndex = typeIndex, Val = 0 });
        cellMeta.Append(block);
        cellMeta.Count = (uint)(cellMeta.Count ?? 0) + 1;
    }

    /// <summary>
    /// If any cell in the workbook has an in-cell image (CellImage), write the
    /// four rich data XML parts and update metadata.xml with XLRICHVALUE entries.
    /// Sets each cell's ValueMetaIndex and SliceCellValue for the sheet writer.
    /// </summary>
    private void EnsureRichValueImageParts(WorkbookPart workbookPart, SaveContext context)
    {
        // Collect all cells with CellImage across all worksheets
        var cellsWithImages = new List<(XLCell Cell, XLCellImage Image)>();
        foreach (var ws in WorksheetsInternal.Cast<XLWorksheet>())
        {
            foreach (var cell in ws.Internals.CellsCollection.GetCells(c => c.CellImage is not null))
            {
                cellsWithImages.Add((cell, cell.CellImage!));
            }
        }

        if (cellsWithImages.Count == 0)
            return;

        // Build rich value entries (one per cell)
        var entries = new List<RichDataWriter.RichValueEntry>(cellsWithImages.Count);
        foreach (var (_, image) in cellsWithImages)
        {
            entries.Add(new RichDataWriter.RichValueEntry(image.WorkbookImageIndex, image.AltText));
        }

        // Write the four rich data XML parts
        RichDataWriter.WriteRichDataParts(workbookPart, InCellImages, entries, context.RelIdGenerator);

        var metadata = EnsureMetadataPart(workbookPart, context);
        var richValueTypeIndex = EnsureRichValueMetadataType(metadata);

        AppendRichValueFutureMetadata(metadata, entries);
        var valueMeta = PrepareValueMetadata(metadata, richValueTypeIndex);

        // Add one valueMetadata record per cell, set each cell's ValueMetaIndex
        for (var i = 0; i < cellsWithImages.Count; i++)
        {
            var block = new MetadataBlock();
            block.Append(new MetadataRecord { TypeIndex = richValueTypeIndex, Val = (uint)i });
            valueMeta.Append(block);
            valueMeta.Count = (uint)(valueMeta.Count ?? 0) + 1;

            var cell = cellsWithImages[i].Cell;
            cell.ValueMetaIndex = valueMeta.Count.Value; // 1-based
            cell.SliceCellValue = (double)i; // rv index as number
        }
    }

    private static Metadata EnsureMetadataPart(WorkbookPart workbookPart, SaveContext context)
    {
        var cellMetadataPart = workbookPart.CellMetadataPart;
        if (cellMetadataPart is not null)
            return cellMetadataPart.Metadata ?? new Metadata();

        cellMetadataPart = workbookPart.AddNewPart<CellMetadataPart>(
            context.RelIdGenerator.GetNext(RelType.Workbook));
        var metadata = new Metadata();
        cellMetadataPart.Metadata = metadata;
        return metadata;
    }

    private static uint EnsureRichValueMetadataType(Metadata metadata)
    {
        var metadataTypes = metadata.MetadataTypes;
        if (metadataTypes is null)
        {
            metadataTypes = new MetadataTypes { Count = 0 };
            metadata.Append(metadataTypes);
        }

        uint typeIdx = 1;
        foreach (var mt in metadataTypes.Elements<MetadataType>())
        {
            if (mt.Name?.Value == "XLRICHVALUE")
                return typeIdx;
            typeIdx++;
        }

        metadataTypes.Append(new MetadataType
        {
            Name = "XLRICHVALUE",
            MinSupportedVersion = 120000,
            Copy = true,
            PasteAll = true,
            PasteValues = true,
            Merge = true,
            SplitFirst = true,
            RowColumnShift = true,
            ClearFormats = true,
            ClearComments = true,
            Assign = true,
            Coerce = true
        });
        metadataTypes.Count = (uint)(metadataTypes.Count ?? 0) + 1;
        return metadataTypes.Count.Value;
    }

    private static void AppendRichValueFutureMetadata(Metadata metadata, List<RichDataWriter.RichValueEntry> entries)
    {
        var existingFm = metadata.Elements<FutureMetadata>()
            .FirstOrDefault(fm => fm.Name?.Value == "XLRICHVALUE");
        existingFm?.Remove();

        const string xlrvNs = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";
        var futureMetadata = new FutureMetadata { Name = "XLRICHVALUE", Count = (uint)entries.Count };
        for (var i = 0; i < entries.Count; i++)
        {
            var fmBlock = new FutureMetadataBlock();
            var extList = new ExtensionList();
            var ext = new Extension { Uri = "{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}" };

            var rvb = new OpenXmlUnknownElement("xlrv", "rvb", xlrvNs);
            rvb.SetAttribute(new OpenXmlAttribute("", "i", "", i.ToString()));
            ext.Append(rvb);
            extList.Append(ext);
            fmBlock.Append(extList);
            futureMetadata.Append(fmBlock);
        }

        metadata.Append(futureMetadata);
    }

    private static ValueMetadata PrepareValueMetadata(Metadata metadata, uint richValueTypeIndex)
    {
        var valueMeta = metadata.GetFirstChild<ValueMetadata>();
        if (valueMeta is not null)
        {
            var blocksToRemove = new List<MetadataBlock>();
            foreach (var bk in valueMeta.Elements<MetadataBlock>())
            {
                var rc = bk.GetFirstChild<MetadataRecord>();
                if (rc?.TypeIndex?.Value == richValueTypeIndex)
                    blocksToRemove.Add(bk);
            }

            foreach (var bk in blocksToRemove)
            {
                bk.Remove();
                valueMeta.Count = (uint)(valueMeta.Count ?? 1) - 1;
            }
        }
        else
        {
            valueMeta = new ValueMetadata { Count = 0 };
            metadata.Append(valueMeta);
        }

        return valueMeta;
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
            // The part might have been removed when a worksheet with pivot tables was deleted.
            if (workbookPart.TryGetPartById(relId, out var part) && part is PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart)
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
