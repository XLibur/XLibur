using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel.Coordinates;
using XLibur.Excel.Tables;

namespace XLibur.Excel;

public partial class XLWorkbook
{
    #region Nested type: SaveContext

    internal sealed class SaveContext
    {
        public SaveContext()
        {
            DifferentialFormats = new Dictionary<XLStyleValue, int>();
            ColorFilterDxfIds = new Dictionary<(XLColorKey, bool), int>();
            RelIdGenerator = new RelIdGenerator();
            SharedFonts = new Dictionary<XLFontValue, FontInfo>();
            SharedNumberFormats = new Dictionary<XLNumberFormatValue, NumberFormatInfo>();
            SharedStyles = new Dictionary<XLStyleValue, StyleInfo>();
            TableId = 0;
            TableNames = [];
            PivotCacheIds = new Dictionary<XLPivotCache, uint>();
        }

        public Dictionary<XLStyleValue, int> DifferentialFormats { get; private set; }

        /// <summary>
        /// Maps (color key, byCellColor) to a dxf index for auto-filter color filters.
        /// </summary>
        public Dictionary<(XLColorKey Color, bool ByCellColor), int> ColorFilterDxfIds { get; private set; }

        public RelIdGenerator RelIdGenerator { get; private set; }
        public Dictionary<XLFontValue, FontInfo> SharedFonts { get; private set; }
        public Dictionary<XLNumberFormatValue, NumberFormatInfo> SharedNumberFormats { get; private set; }
        public Dictionary<XLStyleValue, StyleInfo> SharedStyles { get; private set; }
        public uint TableId { get; set; }
        public HashSet<string> TableNames { get; private set; }

        /// <summary>
        /// The id each pivot cache is written under, assigned while the <c>pivotCaches</c>
        /// element of workbook.xml is rebuilt and read back when each pivot table part writes
        /// its <c>cacheId</c> attribute. The id is a position in that rebuilt element, so it
        /// belongs to the save rather than to the cache, and is only meaningful within one.
        /// </summary>
        public Dictionary<XLPivotCache, uint> PivotCacheIds { get; private set; }

        /// <summary>
        /// The <c>table/@id</c> each table is written under, recorded as the table parts are
        /// generated and read back when a table slicer's cache names the table it filters.
        /// </summary>
        /// <remarks>
        /// Like <see cref="PivotCacheIds"/>, the id belongs to the save rather than to the table:
        /// it comes from <see cref="TableId"/>, a counter over the tables in write order, so it is
        /// only knowable once the part has been written and only meaningful within one save.
        /// </remarks>
        public Dictionary<XLTable, uint> TableIds { get; } = new();

        /// <summary>
        /// A map of shared string ids. The index is the actual index from sharedStringId, and
        /// the value is a mapped stringId to write to a file. The mapped stringId has no gaps
        /// between ids.
        /// </summary>
        public int[] SstMap { get; set; } = null!;

        /// <summary>
        /// 1-based index into <c>cellMetadata</c> records for the XLDAPR (dynamic array)
        /// metadata type. When <c>null</c>, no dynamic array formulas are present.
        /// </summary>
        public uint? DynamicArrayMetaIndex { get; set; }

        internal int GetSharedStringId(XLCell xlCell, string text)
        {
            return GetSharedStringId(xlCell.MemorySstId, xlCell.SheetPoint);
        }

        internal int GetSharedStringId(int memorySstId, Point point)
        {
            var sharedStringId = SstMap[memorySstId];
            if (sharedStringId < 0)
            {
                throw new UnreachableException($"Unable to find SST id {memorySstId} in shared string table for cell {point}. " +
                                               "That likely means reference counting is broken. As a stop-gap, try to set the " +
                                               "text value to an unused cell to increase number of references for the text.");
            }

            return sharedStringId;
        }

        /// <summary>
        /// The id <paramref name="pivotCache"/> is written under. Every cache reachable from a
        /// pivot table is registered while <c>pivotCaches</c> is rebuilt, which happens before
        /// any pivot table part is written, so a miss means the two have gone out of step.
        /// </summary>
        internal uint GetPivotCacheId(XLPivotCache pivotCache)
        {
            if (!PivotCacheIds.TryGetValue(pivotCache, out var cacheId))
            {
                throw new UnreachableException(
                    "Pivot cache was not assigned an id while the workbook pivotCaches element was rebuilt. " +
                    "Every cache used by a pivot table is registered there before any pivot table part is " +
                    "written, so the pivot table being written references a cache the workbook does not list.");
            }

            return cacheId;
        }

        internal int? GetNumberFormat(XLNumberFormatValue? numberFormat)
        {
            if (numberFormat is null)
                return null;

            return SharedNumberFormats.TryGetValue(numberFormat, out var customFormat)
                ? customFormat.NumberFormatId
                : numberFormat.NumberFormatId;
        }
    }

    #endregion Nested type: SaveContext

    #region Nested type: RelType

    internal enum RelType
    {
        Workbook
    }

    #endregion Nested type: RelType

    #region Nested type: RelIdGenerator

    internal sealed class RelIdGenerator
    {
        private readonly Dictionary<RelType, HashSet<string>> _relIds = new();

        private void AddValues(IEnumerable<string> values, RelType relType)
        {
            if (!_relIds.TryGetValue(relType, out var set))
            {
                set = new HashSet<string>();
                _relIds.Add(relType, set);
            }

            set.UnionWith(values);
        }

        /// <summary>
        /// Add all existing rel ids present on the parts or workbook to the generator, so they are not duplicated again.
        /// </summary>
        public void AddExistingValues(WorkbookPart workbookPart, XLWorkbook xlWorkbook)
        {
            AddValues(workbookPart.Parts.Select(p => p.RelationshipId), RelType.Workbook);
            AddValues(xlWorkbook.WorksheetsInternal.Cast<XLWorksheet>().Where(ws => !string.IsNullOrWhiteSpace(ws.RelId)).Select(ws => ws.RelId!), RelType.Workbook);
            AddValues(xlWorkbook.WorksheetsInternal.Cast<XLWorksheet>().Where(ws => !string.IsNullOrWhiteSpace(ws.LegacyDrawingId)).Select(ws => ws.LegacyDrawingId!), RelType.Workbook);
            AddValues(xlWorkbook.WorksheetsInternal
                .Cast<XLWorksheet>()
                .SelectMany(ws => ws.Tables.Cast<XLTable>())
                .Where(t => !string.IsNullOrWhiteSpace(t.RelId))
                .Select(t => t.RelId!), RelType.Workbook);

            foreach (var xlWorksheet in xlWorkbook.WorksheetsInternal.Cast<XLWorksheet>())
            {
                // if the worksheet is new, it doesn't have RelId yet.
                if (string.IsNullOrEmpty(xlWorksheet.RelId) || !workbookPart.TryGetPartById(xlWorksheet.RelId, out var part))
                    continue;

                var worksheetPart = (WorksheetPart)part;
                AddValues(worksheetPart.HyperlinkRelationships.Select(hr => hr.Id), RelType.Workbook);
                AddValues(worksheetPart.Parts.Select(p => p.RelationshipId), RelType.Workbook);
                if (worksheetPart.DrawingsPart != null)
                    AddValues(worksheetPart.DrawingsPart.Parts.Select(p => p.RelationshipId), RelType.Workbook);
            }
        }

        public string GetNext(RelType relType)
        {
            if (!_relIds.TryGetValue(relType, out var set))
            {
                set = [];
                _relIds.Add(relType, set);
            }

            var id = set.Count + 1;
            while (true)
            {
                var relId = string.Concat("rId", id);
                if (set.Add(relId))
                {
                    return relId;
                }
                id++;
            }
        }

        public void Reset(RelType relType)
        {
            _relIds.Remove(relType);
        }
    }

    #endregion Nested type: RelIdGenerator

    #region Nested type: FontInfo

    internal struct FontInfo
    {
        public XLFontValue Font;
        public uint FontId;
    }

    #endregion Nested type: FontInfo

    #region Nested type: FillInfo

    internal struct FillInfo
    {
        public XLFillValue Fill;
        public uint FillId;
    }

    #endregion Nested type: FillInfo

    #region Nested type: BorderInfo

    internal struct BorderInfo
    {
        public XLBorderValue Border;
        public uint BorderId;
    }

    #endregion Nested type: BorderInfo

    #region Nested type: NumberFormatInfo

    internal struct NumberFormatInfo
    {
        public XLNumberFormatValue NumberFormat;
        public int NumberFormatId;
    }

    #endregion Nested type: NumberFormatInfo

    #region Nested type: StyleInfo

    internal struct StyleInfo
    {
        public uint BorderId;
        public uint FillId;
        public uint FontId;
        public bool IncludeQuotePrefix;
        public int NumberFormatId;
        public XLStyleValue Style;
        public uint StyleId;
    }

    #endregion Nested type: StyleInfo
}
