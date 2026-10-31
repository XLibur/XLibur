using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.Coordinates;
using XLibur.Extensions;
using XLibur.Utils;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace XLibur.Excel.IO;

internal static class PivotTableCacheDefinitionPartReader
{
    internal static void Load(WorkbookPart workbookPart, XLWorkbook workbook)
    {
        foreach (var pivotTableCacheDefinitionPart in workbookPart.GetPartsOfType<PivotTableCacheDefinitionPart>())
            LoadCache(workbookPart, workbook, pivotTableCacheDefinitionPart);
    }

    private static void LoadCache(WorkbookPart workbookPart, XLWorkbook workbook,
        PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart)
    {
        var cacheDefinition = pivotTableCacheDefinitionPart.PivotCacheDefinition;
        if (cacheDefinition?.CacheSource is not { } cacheSource)
            throw PartStructureException.RequiredElementIsMissing();

        var pivotSourceReference = ParsePivotSourceReference(cacheSource);

        // One cache per part, even where several parts name the same source. Excel writes a
        // part per cache and the save path reads the relationship back off each cache to find
        // it again, so folding them together here would write out fewer parts than came in.
        var pivotCache = workbook.PivotCachesInternal.Add(pivotSourceReference);
        pivotCache.WorkbookCacheRelId = workbookPart.GetIdOfPart(pivotTableCacheDefinitionPart);

        if (cacheDefinition.MissingItemsLimit?.Value is { } missingItemsLimit)
        {
            pivotCache.ItemsToRetainPerField = missingItemsLimit switch
            {
                0 => XLItemsToRetain.None,
                XLHelper.MaxRowNumber => XLItemsToRetain.Max,
                _ => XLItemsToRetain.Automatic,
            };
        }

        if (cacheDefinition.CacheFields is { } cacheFields)
        {
            ReadCacheFields(cacheFields, pivotCache);
            if (pivotTableCacheDefinitionPart.PivotTableCacheRecordsPart?.PivotCacheRecords is { } recordsPart)
            {
                PivotTableCacheRecordsPartReader.Read(recordsPart, pivotCache);
            }
        }

        pivotCache.SaveSourceData = cacheDefinition.SaveData?.Value ?? true;
        pivotCache.RefreshDataOnOpen = cacheDefinition.RefreshOnLoad?.Value ?? false;

        // A slicer cache binds to a pivot cache through this identifier, which lives in an
        // extension rather than on the element itself and is unrelated to the renumbered
        // pivotCache/@cacheId in workbook.xml. Read so that a slicer added to a loaded pivot
        // table quotes the id the file already uses instead of inventing a second one.
        pivotCache.PivotCacheId = cacheDefinition
            .Descendants<X14.PivotCacheDefinition>()
            .FirstOrDefault()?.PivotCacheId?.Value;
    }

    internal static IXLPivotSource ParsePivotSourceReference(CacheSource cacheSource)
    {
        // Cache source has several types. Each has a specific required format. Do not use different
        // combinations, Excel will crash or at least try to repair
        // [worksheet] uses a worksheet source:
        //   * An unnamed range in a sheet: Uses `sheet` and `ref`.
        //   * An table: Uses `name` that contains a name of the table.
        // [external]
        //   * `connectionId` link to external relationships.
        // [consolidation]
        //  * uses consolidation tag and a list of range sets plus optionally
        //    page fields to add a custom report fields that allow user to select
        //    ranges from rangeSet to calculate values.
        // [scenario]
        //  * only type attribute tag is specified, no other value. Likely linked
        //    through cacheField names (e.g. <cacheField name="$A$1 by">).

        // Not all sources are supported, but at least pipe the data through so the load/save works
        IEnumValue sourceType = cacheSource.Type?.Value ?? throw PartStructureException.MissingAttribute();
        if (sourceType.Equals(SourceValues.Worksheet))
            return ParseWorksheetSource(cacheSource);

        if (sourceType.Equals(SourceValues.External))
        {
            if (cacheSource.ConnectionId?.Value is not { } connectionId)
                throw PartStructureException.MissingAttribute("connectionId");

            return new XLPivotSourceConnection(connectionId);
        }

        if (sourceType.Equals(SourceValues.Consolidation))
            return ParseConsolidationSource(cacheSource);

        if (sourceType.Equals(SourceValues.Scenario))
        {
            return new XLPivotSourceScenario();
        }

        throw PartStructureException.InvalidAttributeValue(sourceType.Value);
    }

    private static IXLPivotSource ParseWorksheetSource(CacheSource cacheSource)
    {
        var sheetSource = cacheSource.WorksheetSource;
        if (sheetSource is null)
            throw PartStructureException.ExpectedElementNotFound("'worksheetSource' element is required for type 'worksheet'.");

        // If the source is a defined name, it must be a single area reference
        if (sheetSource.Name?.Value is { } tableOrName)
        {
            if (sheetSource.Id?.Value is { } externalWorkbookRelId)
                return new XLPivotSourceExternalWorkbook(externalWorkbookRelId, tableOrName);

            return new XLPivotSourceReference(tableOrName);
        }

        if (sheetSource.Sheet?.Value is { } sheetName &&
            sheetSource.Reference?.Value is { } areaRef &&
            Area.TryParse(areaRef.AsSpan(), out var sheetArea))
        {
            var area = new SheetArea(sheetName, sheetArea);
            if (sheetSource.Id?.Value is { } externalWorkbookRelId)
                return new XLPivotSourceExternalWorkbook(externalWorkbookRelId, area);

            // area is in this workbook
            return new XLPivotSourceReference(area);
        }

        throw PartStructureException.IncorrectElementFormat("worksheetSource");
    }

    private static XLPivotSourceConsolidation ParseConsolidationSource(CacheSource cacheSource)
    {
        if (cacheSource.Consolidation is not { } consolidation)
            throw PartStructureException.ExpectedElementNotFound("consolidation");

        var autoPage = consolidation.AutoPage?.Value ?? true;
        var xlPages = new List<XLPivotCacheSourceConsolidationPage>();
        if (consolidation.Pages is { } pages)
        {
            // There is 1..4 pages
            foreach (var page in pages.Cast<Page>())
            {
                var xlPageItems = new List<string>();
                foreach (var pageItem in page.Cast<PageItem>())
                {
                    var pageItemName = pageItem.Name?.Value ?? throw PartStructureException.MissingAttribute();
                    xlPageItems.Add(pageItemName);
                }

                xlPages.Add(new XLPivotCacheSourceConsolidationPage(xlPageItems));
            }
        }

        if (consolidation.RangeSets is not { } rangeSets)
            throw PartStructureException.RequiredElementIsMissing();

        var xlRangeSets = new List<XLPivotCacheSourceConsolidationRangeSet>();
        foreach (var rangeSet in rangeSets.Cast<RangeSet>())
            xlRangeSets.Add(GetRangeSet(rangeSet, xlPages));

        if (xlRangeSets.Count < 1)
            throw PartStructureException.IncorrectElementsCount();

        return new XLPivotSourceConsolidation
        {
            AutoPage = autoPage,
            Pages = xlPages,
            RangeSets = xlRangeSets
        };
    }

    private static XLPivotCacheSourceConsolidationRangeSet GetRangeSet(RangeSet rangeSet, List<XLPivotCacheSourceConsolidationPage> xlPages)
    {
        var pageIndexes = new[]
        {
            rangeSet.FieldItemIndexPage1?.Value,
            rangeSet.FieldItemIndexPage2?.Value,
            rangeSet.FieldItemIndexPage3?.Value,
            rangeSet.FieldItemIndexPage4?.Value,
        };

        ValidateRangeSetPageIndexes(pageIndexes, xlPages);

        if (rangeSet.Name?.Value is { } tableOrName)
        {
            return new XLPivotCacheSourceConsolidationRangeSet
            {
                Indexes = pageIndexes,
                RelId = rangeSet.Id?.Value,
                TableOrName = tableOrName,
            };
        }

        if (rangeSet.Sheet?.Value is { } sheet &&
            rangeSet.Reference?.Value is { } reference &&
            Area.TryParse(reference.AsSpan(), out var area))
        {
            return new XLPivotCacheSourceConsolidationRangeSet
            {
                Indexes = pageIndexes,
                RelId = rangeSet.Id?.Value,
                Area = new SheetArea(sheet, area)
            };
        }

        throw PartStructureException.IncorrectElementFormat("rangeSet");
    }

    private static void ValidateRangeSetPageIndexes(uint?[] pageIndexes, List<XLPivotCacheSourceConsolidationPage> xlPages)
    {
        // Validate that supplied indexes reference existing page and page items
        for (var i = 0; i < pageIndexes.Length; ++i)
        {
            var pageIndex = pageIndexes[i];

            // If there is a page and rangeSet doesn't define index to the page, it is displayed as blank
            if (pageIndex is null)
                continue;

            // Range set points to a non-existent page filter
            if (i >= xlPages.Count)
                throw PartStructureException.IncorrectAttributeValue();

            // Range set points to a non-existent item in a page filter
            var pageFilter = xlPages[i];
            if (pageIndex.Value >= pageFilter.PageItems.Count)
                throw PartStructureException.IncorrectAttributeValue();
        }
    }

    private static void ReadCacheFields(CacheFields cacheFields, XLPivotCache pivotCache)
    {
        foreach (var cacheField in cacheFields.Elements<CacheField>())
        {
            if (cacheField.Name?.Value is not { } fieldName)
                throw PartStructureException.MissingAttribute();

            if (pivotCache.ContainsField(fieldName))
            {
                // We don't allow duplicate field names... but what do we do if we find one? Let's just skip it.
                continue;
            }

            var fieldStats = ReadCacheFieldStats(cacheField);
            var fieldSharedItems = cacheField.SharedItems is not null
                ? ReadSharedItems(cacheField)
                : new XLPivotCacheSharedItems();

            var fieldValues = new XLPivotCacheValues(fieldSharedItems, fieldStats);
            pivotCache.AddCachedField(fieldName, fieldValues);

            var fieldIndex = pivotCache.FieldCount - 1;

            // A calculated field has a formula and DatabaseField=false. It doesn't have
            // records in the cache records part — its values are computed by Excel.
            if (cacheField.Formula?.Value is { } formula)
            {
                pivotCache.SetCalculatedField(fieldIndex, formula);
            }
            else if (cacheField.DatabaseField?.Value == false)
            {
                // A grouping field (e.g. months grouped from a date field) has
                // DatabaseField=false but no formula. It also has no records.
                pivotCache.SetNonDatabaseField(fieldIndex);
            }

            // Not modelled, but kept: a grouping field is nothing without its group, and the items
            // of a pivot field over a grouped field index the group items, not the shared items.
            if (cacheField.FieldGroup is { } fieldGroup)
                pivotCache.SetFieldGroupXml(fieldIndex, fieldGroup.OuterXml);
        }
    }

    private static XLPivotCacheValuesStats ReadCacheFieldStats(CacheField cacheField)
    {
        var sharedItems = cacheField.SharedItems;

        // Various statistics about the records of the field, not just shared items.
        var containsBlank = OpenXmlHelper.GetBooleanValueAsBool(sharedItems?.ContainsBlank, false);
        var containsNumber = OpenXmlHelper.GetBooleanValueAsBool(sharedItems?.ContainsNumber, false);
        var containsOnlyInteger = OpenXmlHelper.GetBooleanValueAsBool(sharedItems?.ContainsInteger, false);
        var minValue = sharedItems?.MinValue?.Value;
        var maxValue = sharedItems?.MaxValue?.Value;
        var containsDate = OpenXmlHelper.GetBooleanValueAsBool(sharedItems?.ContainsDate, false);
        var minDate = sharedItems?.MinDate?.Value;
        var maxDate = sharedItems?.MaxDate?.Value;
        var containsString = OpenXmlHelper.GetBooleanValueAsBool(sharedItems?.ContainsString, true);
        var longText = OpenXmlHelper.GetBooleanValueAsBool(sharedItems?.LongText, false);

        // The containsMixedTypes, containsNonDate and containsSemiMixedTypes are derived from primary stats.
        return new XLPivotCacheValuesStats
        {
            ContainsBlank = containsBlank,
            ContainsNumber = containsNumber,
            ContainsInteger = containsOnlyInteger,
            MinValue = minValue,
            MaxValue = maxValue,
            ContainsString = containsString,
            LongText = longText,
            ContainsDate = containsDate,
            MinDate = minDate,
            MaxDate = maxDate,
        };
    }

    private static XLPivotCacheSharedItems ReadSharedItems(CacheField cacheField)
    {
        var sharedItems = new XLPivotCacheSharedItems();

        // If there are no shared items, the cache record can't contain field items
        // referencing the shared items.
        if (cacheField.SharedItems is not { } fieldSharedItems)
            return sharedItems;

        foreach (var item in fieldSharedItems.Elements())
        {
            // Shared items can't contain element of type index (`x`),
            // because index references shared items.
            PivotCacheValueCodec.ReadValue(sharedItems, item);
        }

        return sharedItems;
    }
}
