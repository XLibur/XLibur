using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using static XLibur.Excel.XLWorkbook;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;

namespace XLibur.Excel.IO;

/// <summary>
/// Writes the workbook half of a timeline: its cache part, the workbook's registration of that cache
/// and the defined name Excel writes alongside it.
/// </summary>
/// <remarks>
/// <para>
/// A created timeline needs six things and Excel offers to repair the file if any is missing: the
/// timeline definition and its worksheet reference (both <see cref="TimelineWriter"/>), the cache
/// part, the workbook <c>extLst</c> registration, the <c>#N/A</c> defined name, and a drawing
/// anchor. Everything but the anchor is written here or next door.
/// </para>
/// <para>
/// Loaded caches are not rewritten. As with <see cref="TimelinePatcher"/>, the part a timeline was
/// read from is left exactly as it arrived.
/// </para>
/// </remarks>
internal static class TimelineCacheWriter
{
    /// <summary>The workbook extension registering timeline caches.</summary>
    private const string TimelineCachesExtensionUri = "{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}";

    private const string X15Main2010SsNs = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";

    /// <summary>
    /// Allocates cache parts and defined names for timelines created since the last save, and takes
    /// them away again for timelines removed since.
    /// </summary>
    /// <remarks>
    /// Runs before the workbook part is generated, because the <c>#N/A</c> defined names have to be
    /// in the model by the time <see cref="WorkbookPartWriter"/> rebuilds the whole defined-name
    /// block from it. The cache parts only get their content later, once the pivot cache identifier
    /// they quote has been assigned.
    /// </remarks>
    internal static void PrepareTimelineCaches(
        WorkbookPart workbookPart, XLWorkbook workbook, SaveContext context)
    {
        foreach (var worksheet in workbook.WorksheetsInternal)
        {
            foreach (var removed in worksheet.TimelinesInternal.Removed)
                RemoveCache(workbookPart, workbook, removed.Cache);

            foreach (var timeline in worksheet.TimelinesInternal.Items)
            {
                if (timeline.IsNew)
                    AddCache(workbookPart, workbook, timeline.Cache, context);
            }
        }
    }

    /// <summary>
    /// Writes the content of every timeline cache part created this save, and registers it in the
    /// workbook's extension list.
    /// </summary>
    /// <remarks>
    /// Runs after the worksheets, because the cache quotes the pivot cache identifier, which does
    /// not exist until the pivot cache part has been generated.
    /// </remarks>
    internal static void WriteTimelineCaches(
        WorkbookPart workbookPart, XLWorkbook workbook, SaveContext context)
    {
        foreach (var worksheet in workbook.WorksheetsInternal)
        {
            foreach (var timeline in worksheet.TimelinesInternal.Items)
            {
                var cache = timeline.Cache;
                if (!cache.IsNew || cache.WorkbookRelId is not { } relId)
                    continue;

                var part = (TimeLineCachePart)workbookPart.GetPartById(relId);
                part.TimelineCacheDefinition = BuildDefinition(cache);

                RegisterCache(workbookPart, relId);
                cache.IsNew = false;
            }

            worksheet.TimelinesInternal.Removed.Clear();
        }
    }

    // ── Cache parts ─────────────────────────────────────────────────────

    private static void AddCache(
        WorkbookPart workbookPart, XLWorkbook workbook, XLTimelineCache cache, SaveContext context)
    {
        if (cache.WorkbookRelId is not null)
            return;

        var relId = context.RelIdGenerator.GetNext(RelType.Workbook);
        cache.WorkbookRelId = relId;
        workbookPart.AddNewPart<TimeLineCachePart>(relId);

        // Excel writes a #N/A defined name per timeline cache, named after the cache. Adding it the
        // way the reader does keeps it out of formula validation, which would reject #N/A.
        if (workbook.DefinedNamesInternal.All<XLDefinedName>(n => !XLHelper.NameComparer.Equals(n.Name, cache.Name)))
        {
            workbook.DefinedNamesInternal.Add(
                cache.Name, "#N/A", comment: null, validateName: false, validateRangeAddress: false);
        }

        // A timeline cache names its pivot cache by an identifier that lives in an extension of the
        // pivot cache definition — the same one a slicer cache quotes. A cache read from a file
        // already has one; a cache XLibur created has none until now, and the pivot cache writer
        // emits it once this is set.
        if (cache.PivotCache is { } pivotCache)
            pivotCache.PivotCacheId ??= NextPivotCacheId(workbook);
    }

    private static void RemoveCache(WorkbookPart workbookPart, XLWorkbook workbook, XLTimelineCache cache)
    {
        if (cache.WorkbookRelId is not { } relId)
            return;

        if (workbookPart.Parts.Any(p => p.RelationshipId == relId)
            && workbookPart.GetPartById(relId) is TimeLineCachePart part)
        {
            workbookPart.DeletePart(part);
        }

        UnregisterCache(workbookPart, relId);

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
    /// Counting up from the highest in use keeps it deterministic, which matters because a save has
    /// to be reproducible. Shared with the slicer path by convention rather than by code: both read
    /// and write the same <c>XLPivotCache.PivotCacheId</c>, so a workbook holding both never
    /// allocates a collision.
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

    private static X15.TimelineCacheDefinition BuildDefinition(XLTimelineCache cache)
    {
        var definition = new X15.TimelineCacheDefinition
        {
            Name = cache.Name,
            SourceName = cache.SourceName,
        };
        definition.AddNamespaceDeclaration("x15", X15Main2010SsNs);

        var pivotTables = new X15.TimelineCachePivotTables();
        foreach (var pivotTable in cache.PivotTables)
        {
            // tabId is the sheet the pivot table lives on, which need not be the sheet the timeline
            // is drawn on.
            var sheetId = ((XLWorksheet)pivotTable.Worksheet).SheetId;
            pivotTables.AppendChild(new X15.TimelineCachePivotTable
            {
                TabId = (uint)sheetId,
                Name = pivotTable.Name,
            });
        }

        definition.AppendChild(pivotTables);

        var state = new X15.TimelineState
        {
            MinimalRefreshVersion = cache.MinimalRefreshVersion,
            LastRefreshVersion = cache.LastRefreshVersion,
            PivotCacheId = cache.PivotCache?.PivotCacheId ?? 0,

            // The SDK types this as an enumeration, but InnerText is what actually serialises, so a
            // token this build does not know still round-trips.
            FilterType = new EnumValue<PivotFilterValues> { InnerText = cache.FilterType },
        };

        if (cache.BoundsStart is { } boundsStart && cache.BoundsEnd is { } boundsEnd)
        {
            state.AppendChild(new X15.BoundsTimelineRange
            {
                StartDate = boundsStart,
                EndDate = boundsEnd,
            });
        }

        definition.AppendChild(state);
        return definition;
    }

    // ── Workbook registration ───────────────────────────────────────────

    private static void RegisterCache(WorkbookPart workbookPart, string relId)
    {
        var workbook = workbookPart.Workbook!;
        var extensionList = workbook.GetFirstChild<WorkbookExtensionList>();
        if (extensionList is null)
        {
            extensionList = new WorkbookExtensionList();
            workbook.AppendChild(extensionList);
        }

        var extension = FindExtension(extensionList);
        if (extension is null)
        {
            extension = new WorkbookExtension { Uri = TimelineCachesExtensionUri };
            extension.AddNamespaceDeclaration("x15", X15Main2010SsNs);
            extension.AppendChild(new X15.TimelineCacheReferences());
            extensionList.AppendChild(extension);
        }

        var container = extension.GetFirstChild<X15.TimelineCacheReferences>();
        if (container is null)
            return;

        if (!container.Elements<X15.TimelineCacheReference>().Any(c => c.Id?.Value == relId))
            container.AppendChild(new X15.TimelineCacheReference { Id = relId });
    }

    private static void UnregisterCache(WorkbookPart workbookPart, string relId)
    {
        var extensionList = workbookPart.Workbook?.GetFirstChild<WorkbookExtensionList>();
        var extension = extensionList is null ? null : FindExtension(extensionList);
        var container = extension?.GetFirstChild<X15.TimelineCacheReferences>();
        if (container is null)
            return;

        foreach (var registration in container
                     .Elements<X15.TimelineCacheReference>()
                     .Where(c => c.Id?.Value == relId)
                     .ToList())
        {
            registration.Remove();
        }

        // An empty registry is a schema violation rather than merely untidy, so the extension goes
        // once its last cache does.
        if (!container.Elements<X15.TimelineCacheReference>().Any())
            extension!.Remove();

        if (extensionList is { HasChildren: false })
            extensionList.Remove();
    }

    private static WorkbookExtension? FindExtension(WorkbookExtensionList extensionList) =>
        extensionList.Elements<WorkbookExtension>()
            .FirstOrDefault(e => string.Equals(
                e.Uri?.Value, TimelineCachesExtensionUri, System.StringComparison.OrdinalIgnoreCase));
}
