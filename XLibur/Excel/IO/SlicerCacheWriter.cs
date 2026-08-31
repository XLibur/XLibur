using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.Tables;
using static XLibur.Excel.IO.OpenXmlConst;
using static XLibur.Excel.XLWorkbook;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;

namespace XLibur.Excel.IO;

/// <summary>
/// Writes the workbook half of a slicer: its cache part, the workbook's registration of that cache
/// and the defined name Excel writes alongside it.
/// </summary>
/// <remarks>
/// <para>
/// A created slicer needs six things, not the two obvious ones, and Excel offers to repair the file
/// if any is missing: the slicer definition and its worksheet reference (both
/// <see cref="SlicerWriter"/>), the cache part, the workbook <c>extLst</c> registration, the
/// <c>#N/A</c> defined name, and a drawing anchor. Everything but the anchor is written here or
/// next door.
/// </para>
/// <para>
/// The registration is split across two extensions and the split is not cosmetic:
/// <c>x14:slicerCaches</c> holds pivot slicer caches and <c>x15:slicerCaches</c> holds table slicer
/// caches. Putting a cache in the wrong one orphans it as surely as leaving it out.
/// </para>
/// <para>
/// Loaded caches are not rewritten. As with <see cref="SlicerPatcher"/>, the part a slicer was read
/// from is left exactly as it arrived, which is what carries its <c>xr10:uid</c>, its OLAP
/// selections and its extension list through a save.
/// </para>
/// </remarks>
internal static class SlicerCacheWriter
{
    /// <summary>The workbook extension registering pivot slicer caches.</summary>
    private const string PivotSlicerCachesExtensionUri = "{BBE1A952-AA13-448e-AADC-164F8A28A991}";

    /// <summary>The workbook extension registering table slicer caches.</summary>
    private const string TableSlicerCachesExtensionUri = "{46BE6895-7355-4a93-B00E-2C351335B9C9}";

    /// <summary>The slicer cache extension carrying a table slicer's binding.</summary>
    private const string TableSlicerCacheExtensionUri = "{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}";

    private const string X15Main2010SsNs = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";

    /// <summary>
    /// Allocates cache parts and defined names for slicers created since the last save, and takes
    /// them away again for slicers removed since.
    /// </summary>
    /// <remarks>
    /// Runs before the workbook part is generated, because the <c>#N/A</c> defined names have to be
    /// in the model by the time <see cref="WorkbookPartWriter"/> rebuilds the whole defined-name
    /// block from it. The cache parts only get their content later, once the ids they quote —
    /// a pivot cache's and a table's — have been assigned by the writers that own them.
    /// </remarks>
    internal static void PrepareSlicerCaches(WorkbookPart workbookPart, XLWorkbook workbook, SaveContext context)
    {
        foreach (var worksheet in workbook.WorksheetsInternal)
        {
            foreach (var removed in worksheet.SlicersInternal.Removed)
                RemoveCache(workbookPart, workbook, removed.Cache);

            foreach (var slicer in worksheet.SlicersInternal.Items)
            {
                if (slicer.IsNew)
                    AddCache(workbookPart, workbook, slicer.Cache, context);
            }
        }
    }

    /// <summary>
    /// Writes the content of every slicer cache part created this save, and registers it in the
    /// workbook's extension list.
    /// </summary>
    /// <remarks>
    /// Runs after the worksheets, because a table slicer cache quotes the <c>table/@id</c> the
    /// table part was written under and a pivot slicer cache quotes the pivot cache identifier —
    /// neither of which exists until those parts have been generated.
    /// </remarks>
    internal static void WriteSlicerCaches(WorkbookPart workbookPart, XLWorkbook workbook, SaveContext context)
    {
        foreach (var worksheet in workbook.WorksheetsInternal)
        {
            foreach (var slicer in worksheet.SlicersInternal.Items)
            {
                var cache = slicer.Cache;
                if (!cache.IsNew || cache.WorkbookRelId is not { } relId)
                    continue;

                var part = (SlicerCachePart)workbookPart.GetPartById(relId);
                part.SlicerCacheDefinition = BuildDefinition(cache, context);

                RegisterCache(workbookPart, cache.SourceKind, relId);
                cache.IsNew = false;
            }

            worksheet.SlicersInternal.Removed.Clear();
        }
    }

    // ── Cache parts ─────────────────────────────────────────────────────

    private static void AddCache(
        WorkbookPart workbookPart, XLWorkbook workbook, XLSlicerCache cache, SaveContext context)
    {
        if (cache.WorkbookRelId is not null)
            return;

        var relId = context.RelIdGenerator.GetNext(RelType.Workbook);
        cache.WorkbookRelId = relId;
        workbookPart.AddNewPart<SlicerCachePart>(relId);

        // Excel writes a #N/A defined name per slicer cache, named after the cache. Adding it the
        // way the reader does keeps it out of formula validation, which would reject #N/A.
        if (workbook.DefinedNamesInternal.All<XLDefinedName>(n => !XLHelper.NameComparer.Equals(n.Name, cache.Name)))
        {
            workbook.DefinedNamesInternal.Add(
                cache.Name, "#N/A", comment: null, validateName: false, validateRangeAddress: false);
        }

        // A pivot slicer cache names its pivot cache by an identifier that lives in an extension of
        // the pivot cache definition. A cache read from a file already has one; a cache XLibur
        // created has none until now, and the pivot cache writer emits it once this is set.
        if (cache.SourceKind == XLSlicerSourceKind.PivotTable && cache.PivotCache is { } pivotCache)
            pivotCache.PivotCacheId ??= NextPivotCacheId(workbook);
    }

    private static void RemoveCache(WorkbookPart workbookPart, XLWorkbook workbook, XLSlicerCache cache)
    {
        if (cache.WorkbookRelId is not { } relId)
            return;

        if (workbookPart.Parts.Any(p => p.RelationshipId == relId)
            && workbookPart.GetPartById(relId) is SlicerCachePart part)
        {
            workbookPart.DeletePart(part);
        }

        UnregisterCache(workbookPart, cache.SourceKind, relId);

        var definedName = workbook.DefinedNamesInternal
            .FirstOrDefault<XLDefinedName>(n => XLHelper.NameComparer.Equals(n.Name, cache.Name));
        if (definedName is not null)
            workbook.DefinedNamesInternal.Delete(definedName.Name);

        cache.WorkbookRelId = null;
    }

    /// <summary>
    /// An identifier no pivot cache in the workbook is already using.
    /// </summary>
    /// <remarks>
    /// Excel writes a large arbitrary number here; the value carries no meaning beyond matching the
    /// <c>pivotCacheId</c> a slicer cache quotes. Counting up from the highest in use keeps it
    /// deterministic, which matters because a save has to be reproducible.
    /// </remarks>
    private static uint NextPivotCacheId(XLWorkbook workbook)
    {
        uint highest = 0;
        foreach (var cache in workbook.PivotCachesInternal)
        {
            if (cache.PivotCacheId is { } id && id > highest)
                highest = id;
        }

        return highest + 1;
    }

    private static X14.SlicerCacheDefinition BuildDefinition(XLSlicerCache cache, SaveContext context)
    {
        var definition = new X14.SlicerCacheDefinition
        {
            Name = cache.Name,
            SourceName = cache.SourceName,
        };
        definition.AddNamespaceDeclaration("x", Main2006SsNs);

        if (cache.SourceKind == XLSlicerSourceKind.Table)
            AppendTableBinding(definition, cache, context);
        else
            AppendPivotBinding(definition, cache);

        return definition;
    }

    private static void AppendPivotBinding(X14.SlicerCacheDefinition definition, XLSlicerCache cache)
    {
        var pivotTables = new X14.SlicerCachePivotTables();
        foreach (var pivotTable in cache.PivotTables)
        {
            // tabId is the sheet the pivot table lives on, which need not be the sheet the slicer
            // is drawn on.
            var sheetId = ((XLWorksheet)pivotTable.Worksheet).SheetId;
            pivotTables.AppendChild(new X14.SlicerCachePivotTable
            {
                TabId = sheetId,
                Name = pivotTable.Name,
            });
        }

        definition.AppendChild(pivotTables);

        var tabular = new X14.TabularSlicerCache { PivotCacheId = cache.PivotCache?.PivotCacheId ?? 0 };
        var items = new X14.TabularSlicerCacheItems { Count = (uint)cache.Items.Count };

        // s is absent on an unselected item, so a slicer filtering nothing is every item selected
        // rather than an empty list — see XLSlicerCacheItem.
        foreach (var item in cache.Items)
        {
            var element = new X14.TabularSlicerCacheItem { Atom = item.Index };
            if (item.Selected)
                element.IsSelected = true;

            items.AppendChild(element);
        }

        tabular.AppendChild(items);
        definition.AppendChild(new X14.SlicerCacheData(tabular));
    }

    private static void AppendTableBinding(
        X14.SlicerCacheDefinition definition, XLSlicerCache cache, SaveContext context)
    {
        // Both ids come from what the table part was actually written as, not from the model: the
        // table id is a counter over the tables in write order and the column id is the column's
        // position. A table slicer created against a table that was somehow not written has
        // nothing to bind to, and is left unbound rather than pointed at the wrong table.
        if (cache.Table is XLTable table && context.TableIds.TryGetValue(table, out var tableId))
        {
            cache.TableId = tableId;
            cache.TableColumnId = (uint?)cache.TableColumnPosition;
        }

        var tableSlicerCache = new X15.TableSlicerCache
        {
            TableId = cache.TableId ?? 0,
            Column = cache.TableColumnId ?? 1,
        };

        var extension = new SlicerCacheDefinitionExtension { Uri = TableSlicerCacheExtensionUri };
        extension.AddNamespaceDeclaration("x15", X15Main2010SsNs);
        extension.AppendChild(tableSlicerCache);

        definition.AppendChild(new X14.SlicerCacheDefinitionExtensionList(extension));
    }

    // ── Workbook registration ───────────────────────────────────────────

    private static void RegisterCache(WorkbookPart workbookPart, XLSlicerSourceKind kind, string relId)
    {
        var workbook = workbookPart.Workbook!;
        var extensionList = workbook.GetFirstChild<WorkbookExtensionList>();
        if (extensionList is null)
        {
            extensionList = new WorkbookExtensionList();
            workbook.AppendChild(extensionList);
        }

        var uri = RegistrationUri(kind);
        var extension = FindExtension(extensionList, uri);
        if (extension is null)
        {
            extension = new WorkbookExtension { Uri = uri };
            extension.AddNamespaceDeclaration(
                kind == XLSlicerSourceKind.Table ? "x15" : "x14",
                kind == XLSlicerSourceKind.Table ? X15Main2010SsNs : X14Main2009SsNs);

            if (kind == XLSlicerSourceKind.Table)
            {
                var caches = new X15.SlicerCaches();
                caches.AddNamespaceDeclaration("x14", X14Main2009SsNs);
                extension.AppendChild(caches);
            }
            else
            {
                extension.AppendChild(new X14.SlicerCaches());
            }

            extensionList.AppendChild(extension);
        }

        var container = CacheContainer(extension, kind);
        if (container is null)
            return;

        if (!container.Elements<X14.SlicerCache>().Any(c => c.Id?.Value == relId))
            container.AppendChild(new X14.SlicerCache { Id = relId });
    }

    private static void UnregisterCache(WorkbookPart workbookPart, XLSlicerSourceKind kind, string relId)
    {
        var extensionList = workbookPart.Workbook?.GetFirstChild<WorkbookExtensionList>();
        var extension = extensionList is null ? null : FindExtension(extensionList, RegistrationUri(kind));
        var container = extension is null ? null : CacheContainer(extension, kind);
        if (container is null)
            return;

        foreach (var registration in container.Elements<X14.SlicerCache>().Where(c => c.Id?.Value == relId).ToList())
            registration.Remove();

        // An empty registry is a schema violation rather than merely untidy, so the extension goes
        // once its last cache does.
        if (!container.Elements<X14.SlicerCache>().Any())
            extension!.Remove();

        if (extensionList is { HasChildren: false })
            extensionList.Remove();
    }

    /// <summary>
    /// The element holding the registrations, which is <c>x14:slicerCaches</c> for a pivot cache
    /// and <c>x15:slicerCaches</c> for a table cache. Both hold <c>x14:slicerCache</c> children.
    /// </summary>
    private static OpenXmlCompositeElement? CacheContainer(WorkbookExtension extension, XLSlicerSourceKind kind) =>
        kind == XLSlicerSourceKind.Table
            ? extension.GetFirstChild<X15.SlicerCaches>()
            : extension.GetFirstChild<X14.SlicerCaches>();

    private static WorkbookExtension? FindExtension(WorkbookExtensionList extensionList, string uri) =>
        extensionList.Elements<WorkbookExtension>()
            .FirstOrDefault(e => string.Equals(e.Uri?.Value, uri, System.StringComparison.OrdinalIgnoreCase));

    private static string RegistrationUri(XLSlicerSourceKind kind) =>
        kind == XLSlicerSourceKind.Table ? TableSlicerCachesExtensionUri : PivotSlicerCachesExtensionUri;
}
