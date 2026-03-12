using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace XLibur.Excel;

public partial class XLWorkbook
{
    #region Nested type: SaveContext

    internal sealed class SaveContext
    {
        public SaveContext()
        {
            DifferentialFormats = new Dictionary<XLStyleValue, int>();
            RelIdGenerator = new RelIdGenerator();
            SharedFonts = new Dictionary<XLFontValue, FontInfo>();
            SharedNumberFormats = new Dictionary<XLNumberFormatValue, NumberFormatInfo>();
            SharedStyles = new Dictionary<XLStyleValue, StyleInfo>();
            TableId = 0;
            TableNames = [];
            PivotSourceCacheId = 0;
            PivotSources = new Dictionary<Guid, PivotSourceInfo>();
        }

        public Dictionary<XLStyleValue, int> DifferentialFormats { get; private set; }
        public RelIdGenerator RelIdGenerator { get; private set; }
        public Dictionary<XLFontValue, FontInfo> SharedFonts { get; private set; }
        public Dictionary<XLNumberFormatValue, NumberFormatInfo> SharedNumberFormats { get; private set; }
        public Dictionary<XLStyleValue, StyleInfo> SharedStyles { get; private set; }
        public uint TableId { get; set; }
        public HashSet<string> TableNames { get; private set; }

        /// <summary>
        /// A free id that can be used by the workbook to reference to a pivot cache.
        /// The <c>PivotCaches</c> element in a workbook connects the parts with pivot
        /// cache parts.
        /// </summary>
        public uint PivotSourceCacheId { get; set; }

        /// <summary>
        /// A dictionary of extra info for pivot during saving. The key is <see cref="XLPivotCache.Guid"/>.
        /// </summary>
        public IDictionary<Guid, PivotSourceInfo> PivotSources { get; }

        /// <summary>
        /// A map of shared string ids. The index is the actual index from sharedStringId and
        /// value is an mapped stringId to write to a file. The mapped stringId has no gaps
        /// between ids.
        /// </summary>
        public int[] SstMap { get; set; } = null!;

        internal int GetSharedStringId(XLCell xlCell, string text)
        {
            var sharedStringId = SstMap[xlCell.MemorySstId];
            if (sharedStringId < 0)
            {
                throw new UnreachableException($"Unable to find text '{text}' in shared string table for cell {xlCell.SheetPoint}. " +
                                               "That likely means reference counting is broken. As a stop-gap, try to set the " +
                                               "text value to an unused cell to increase number of references for the text.");
            }

            return sharedStringId;
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
        Workbook//, Worksheet
    }

    #endregion Nested type: RelType

    #region Nested type: RelIdGenerator

    internal sealed class RelIdGenerator
    {
        private readonly Dictionary<RelType, HashSet<string>> _relIds = new();

        public void AddValues(IEnumerable<string> values, RelType relType)
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
                // if the worksheet is a new one, it doesn't have RelId yet.
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
                set = new HashSet<string>();
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
    };

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

    #region Nested type: Pivot tables

    internal struct PivotTableFieldInfo
    {
        public bool MixedDataType;
        public IReadOnlyList<XLCellValue> DistinctValues;
        public bool IsTotallyBlankField;
    }

    internal struct PivotSourceInfo
    {
        public IDictionary<string, PivotTableFieldInfo> Fields;

        public PivotSourceInfo(IDictionary<string, PivotTableFieldInfo> fields)
        {
            Fields = fields;
        }
    }

    #endregion Nested type: Pivot tables
}
