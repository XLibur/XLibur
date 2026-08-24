using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.Tables;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;

namespace XLibur.Excel.IO;

/// <summary>
/// Reads slicers and their caches out of a package and binds them to the pivot tables and tables
/// they filter.
/// </summary>
/// <remarks>
/// <para>
/// <b>Nothing here attaches a DOM to a part.</b> Slicer parts are one of the four independent
/// mechanisms by which slicers survive a round trip today (see <c>docs/round-trip-fidelity.md</c>),
/// and the reason the first of them works is that <c>xl/slicers/*.xml</c> and
/// <c>xl/slicerCaches/*.xml</c> are never opened, so they are copied through byte for byte with
/// every attribute XLibur has no model for. Reaching a part through <c>SlicersPart.Slicers</c> or
/// <c>SlicerCachePart.SlicerCacheDefinition</c> would create an attached DOM that the SDK tracks
/// and writes back on save, replacing those bytes with its own serialisation — the same trap
/// <see cref="WorksheetPartWriter"/> documents for <c>worksheetPart.Worksheet</c>. Every read below
/// therefore goes through an <see cref="OpenXmlPartReader"/>, which streams the part and hands back
/// a detached tree the part knows nothing about.
/// </para>
/// <para>
/// Reading runs after pivot tables and tables have been loaded, because binding needs both.
/// </para>
/// </remarks>
internal static class SlicerReader
{
    /// <remarks>
    /// Declared here rather than shared with <c>ChartSeriesFormatXml</c>, which holds the same
    /// constant: that file belongs to spec 16's DrawingML extraction and is not ours to touch.
    /// </remarks>
    private const double EmuPerPoint = 12700;

    /// <summary>
    /// Loads every slicer in the workbook into the worksheet that draws it.
    /// </summary>
    internal static void LoadSlicers(WorkbookPart workbookPart, Sheets sheets, XLWorksheets worksheets)
    {
        var caches = ReadCaches(workbookPart);
        if (caches.Count == 0)
            return;

        BindCaches(caches.Values, workbookPart, sheets, worksheets);
        ReadSlicers(workbookPart, sheets, worksheets, caches);
    }

    // ── Caches ──────────────────────────────────────────────────────────

    /// <summary>
    /// Reads the workbook's slicer cache parts, keyed by cache name. A slicer refers to its cache by
    /// that name and by nothing else.
    /// </summary>
    private static Dictionary<string, XLSlicerCache> ReadCaches(WorkbookPart workbookPart)
    {
        var caches = new Dictionary<string, XLSlicerCache>(XLHelper.NameComparer);

        foreach (var cachePart in workbookPart.SlicerCacheParts)
        {
            var definition = ReadDetached<X14.SlicerCacheDefinition>(cachePart);
            var name = definition?.Name?.Value;
            if (name is null)
                continue;

            var cache = ReadCache(definition!, name, workbookPart.GetIdOfPart(cachePart));
            caches[name] = cache;
        }

        return caches;
    }

    private static XLSlicerCache ReadCache(X14.SlicerCacheDefinition definition, string name, string relId)
    {
        // The source name is the pivot cache field or the table column the slicer draws its buttons
        // from. Excel always writes it; a file that does not is treated as binding to nothing rather
        // than as a reason to refuse the whole workbook.
        var sourceName = definition.SourceName?.Value ?? string.Empty;

        // The two binding paths are told apart by which one the part carries. A table slicer's cache
        // holds an x15:tableSlicerCache extension and no tabular element; a pivot slicer's cache is
        // the other way round.
        var tableSlicerCache = definition.Descendants<X15.TableSlicerCache>().FirstOrDefault();
        if (tableSlicerCache is not null)
        {
            return new XLSlicerCache(name, sourceName, XLSlicerSourceKind.Table)
            {
                WorkbookRelId = relId,
                TableId = tableSlicerCache.TableId?.Value,
                TableColumnId = tableSlicerCache.Column?.Value,
            };
        }

        var tabular = definition.SlicerCacheData?.GetFirstChild<X14.TabularSlicerCache>();
        var cache = new XLSlicerCache(name, sourceName, XLSlicerSourceKind.PivotTable)
        {
            WorkbookRelId = relId,
            PivotCacheId = tabular?.PivotCacheId?.Value,
        };

        foreach (var pivotTable in definition.SlicerCachePivotTables?.Elements<X14.SlicerCachePivotTable>()
                     ?? [])
        {
            if (pivotTable.Name?.Value is { } pivotTableName)
                cache.PivotTableNames.Add(pivotTableName);
        }

        foreach (var item in tabular?.TabularSlicerCacheItems?.Elements<X14.TabularSlicerCacheItem>() ?? [])
        {
            if (item.Atom?.Value is { } index)
                cache.Items.Add(new XLSlicerCacheItem(index, item.IsSelected?.Value ?? false));
        }

        return cache;
    }

    // ── Binding ─────────────────────────────────────────────────────────

    private static void BindCaches(
        IEnumerable<XLSlicerCache> caches, WorkbookPart workbookPart, Sheets sheets, XLWorksheets worksheets)
    {
        var pivotTables = PivotTablesByName(worksheets);
        Dictionary<uint, (XLTable Table, Dictionary<uint, int> ColumnPositions)>? tables = null;

        foreach (var cache in caches)
        {
            if (cache.SourceKind == XLSlicerSourceKind.PivotTable)
            {
                BindPivotTables(cache, pivotTables);
                continue;
            }

            // Table ids are assigned per package, not modelled, so the map is built from the table
            // parts themselves — and only when a table slicer actually needs it.
            tables ??= TablesById(workbookPart, sheets, worksheets);
            BindTable(cache, tables);
        }
    }

    private static void BindPivotTables(XLSlicerCache cache, Dictionary<string, XLPivotTable> pivotTables)
    {
        // A cache may name several pivot tables — that is how one set of buttons drives a whole
        // dashboard — and may name one that is no longer in the workbook, which is left out rather
        // than reported as a hole in the list.
        foreach (var name in cache.PivotTableNames)
        {
            if (pivotTables.TryGetValue(name, out var pivotTable))
                cache.PivotTables.Add(pivotTable);
        }

        // The item indices are indices into the shared items of the pivot cache behind those pivot
        // tables. Pivot tables sharing a slicer cache share a pivot cache, so the first one answers
        // for all of them.
        cache.PivotCache = cache.PivotTables.Count > 0 ? cache.PivotTables[0].PivotCache : null;
    }

    private static void BindTable(
        XLSlicerCache cache, Dictionary<uint, (XLTable Table, Dictionary<uint, int> ColumnPositions)> tables)
    {
        if (cache.TableId is not { } tableId || !tables.TryGetValue(tableId, out var entry))
            return;

        cache.Table = entry.Table;

        if (cache.TableColumnId is { } columnId && entry.ColumnPositions.TryGetValue(columnId, out var position))
            cache.TableColumnPosition = position;
    }

    private static Dictionary<string, XLPivotTable> PivotTablesByName(XLWorksheets worksheets)
    {
        var pivotTables = new Dictionary<string, XLPivotTable>(XLHelper.NameComparer);
        foreach (var worksheet in worksheets)
        {
            foreach (var pivotTable in worksheet.PivotTables.Cast<XLPivotTable>())
                pivotTables[pivotTable.Name] = pivotTable;
        }

        return pivotTables;
    }

    /// <summary>
    /// Maps each table's <c>table/@id</c> to the table and to the position of each of its columns by
    /// <c>tableColumn/@id</c>.
    /// </summary>
    /// <remarks>
    /// Neither id is modelled: XLibur renumbers tables and their columns on save. Both have to be
    /// read back off the table parts, which the load path has already materialised, so this costs a
    /// walk and no extra parsing.
    /// </remarks>
    private static Dictionary<uint, (XLTable Table, Dictionary<uint, int> ColumnPositions)> TablesById(
        WorkbookPart workbookPart, Sheets sheets, XLWorksheets worksheets)
    {
        var tables = new Dictionary<uint, (XLTable, Dictionary<uint, int>)>();

        foreach (var (worksheetPart, worksheet) in WorksheetParts(workbookPart, sheets, worksheets))
        {
            foreach (var tablePart in worksheetPart.TableDefinitionParts)
            {
                var dTable = tablePart.Table;
                if (dTable?.Id?.Value is not { } tableId)
                    continue;

                var relId = worksheetPart.GetIdOfPart(tablePart);
                var xlTable = worksheet.Tables.Cast<XLTable>().FirstOrDefault(t => t.RelId == relId);
                if (xlTable is null)
                    continue;

                var positions = new Dictionary<uint, int>();
                var position = 0;
                foreach (var column in dTable.TableColumns?.Elements<TableColumn>() ?? [])
                {
                    position++;
                    if (column.Id?.Value is { } columnId)
                        positions[columnId] = position;
                }

                tables[tableId] = (xlTable, positions);
            }
        }

        return tables;
    }

    // ── Slicers ─────────────────────────────────────────────────────────

    private static void ReadSlicers(
        WorkbookPart workbookPart,
        Sheets sheets,
        XLWorksheets worksheets,
        Dictionary<string, XLSlicerCache> caches)
    {
        foreach (var (worksheetPart, worksheet) in WorksheetParts(workbookPart, sheets, worksheets))
        {
            foreach (var slicersPart in worksheetPart.SlicersParts)
            {
                var slicers = ReadDetached<X14.Slicers>(slicersPart);
                if (slicers is null)
                    continue;

                var relId = worksheetPart.GetIdOfPart(slicersPart);
                foreach (var slicer in slicers.Elements<X14.Slicer>())
                    AddSlicer(slicer, relId, worksheet, caches);
            }
        }
    }

    private static void AddSlicer(
        X14.Slicer slicer, string relId, XLWorksheet worksheet, Dictionary<string, XLSlicerCache> caches)
    {
        var name = slicer.Name?.Value;
        var cacheName = slicer.Cache?.Value;
        if (name is null || cacheName is null || !caches.TryGetValue(cacheName, out var cache))
            return;

        worksheet.SlicersInternal.Add(new XLSlicer(worksheet, cache, name)
        {
            PartRelId = relId,

            // Excel omits the caption when it matches the name, and shows the name in that case.
            Caption = slicer.Caption?.Value ?? name,
            ShowCaption = slicer.ShowCaption?.Value ?? true,
            Style = slicer.Style?.Value,
            ColumnCount = slicer.ColumnCount?.Value ?? 1,
            RowHeightPt = slicer.RowHeight?.Value is { } rowHeight
                ? rowHeight / EmuPerPoint
                : null,
        });
    }

    // ── Plumbing ────────────────────────────────────────────────────────

    /// <summary>
    /// Pairs each worksheet part with the loaded worksheet it belongs to, in sheet order.
    /// </summary>
    private static IEnumerable<(WorksheetPart Part, XLWorksheet Worksheet)> WorksheetParts(
        WorkbookPart workbookPart, Sheets sheets, XLWorksheets worksheets)
    {
        foreach (var sheet in sheets.OfType<Sheet>())
        {
            // A sheet with an empty relationship id comes from a non-Excel producer, and the
            // relationship may point at a chartsheet rather than a worksheet.
            if (string.IsNullOrEmpty(sheet.Id?.Value)
                || sheet.Name?.Value is not { } sheetName
                || workbookPart.GetPartById(sheet.Id!.Value!) is not WorksheetPart worksheetPart
                || !worksheets.TryGetWorksheet(sheetName, out var worksheet))
            {
                continue;
            }

            yield return (worksheetPart, worksheet);
        }
    }

    /// <summary>
    /// Reads a part's root element without attaching it to the part.
    /// </summary>
    /// <remarks>
    /// This is the whole fidelity guarantee of this reader in three lines: the part is streamed, the
    /// element that comes back is detached, and <c>part.RootElement</c> stays unmaterialised, so the
    /// SDK has nothing to write back over the original bytes when the package is saved.
    /// </remarks>
    private static T? ReadDetached<T>(OpenXmlPart part) where T : OpenXmlElement
    {
        using var reader = new OpenXmlPartReader(part);

        // Create reads the XML declaration only, so the first Read lands on the root element.
        return reader.Read() ? reader.LoadCurrentElement() as T : null;
    }
}
