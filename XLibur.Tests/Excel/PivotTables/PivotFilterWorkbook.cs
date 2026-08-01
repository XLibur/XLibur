using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;
using XLibur.Excel.IO;

namespace XLibur.Tests.Excel.PivotTables;

/// <summary>
/// Builds the smallest workbook a pivot filter can live in, so that a test can put one filter in
/// and read back what a load and save did to it.
/// </summary>
/// <remarks>
/// Written with the OpenXML SDK rather than XLibur because XLibur has no API for creating pivot
/// filters — which is the whole reason they only have to survive a round trip.
/// </remarks>
internal static class PivotFilterWorkbook
{
    /// <summary>
    /// Load the given filters into XLibur, save, and return the filters as written.
    /// </summary>
    internal static List<PivotFilter> RoundTrip(params PivotFilter[] filters)
    {
        using var input = new MemoryStream();
        Create(input, filters);
        input.Position = 0;

        using var wb = new XLWorkbook(input);
        using var output = new MemoryStream();
        wb.SaveAs(output);

        return ReadFilters(output);
    }

    /// <summary>
    /// The <c>filterColumn</c> of a filter after a round trip — the element most of the criteria
    /// tests are about.
    /// </summary>
    internal static FilterColumn RoundTripFilterColumn(FilterColumn filterColumn)
    {
        var autoFilter = new AutoFilter();
        autoFilter.Append(filterColumn);

        return RoundTripAutoFilter(autoFilter).Elements<FilterColumn>().Single();
    }

    /// <summary>
    /// The <c>autoFilter</c> of a filter after a round trip.
    /// </summary>
    internal static AutoFilter RoundTripAutoFilter(AutoFilter autoFilter)
    {
        var pivotFilter = new PivotFilter
        {
            Field = 0U,
            Id = 1U,
            Type = PivotFilterValues.CaptionEqual,
        };
        pivotFilter.Append(autoFilter);

        return RoundTrip(pivotFilter).Single().AutoFilter!;
    }

    internal static List<PivotFilter> ReadFilters(MemoryStream saved)
    {
        saved.Position = 0;
        using var doc = SpreadsheetDocument.Open(saved, false);
        var definition = doc.WorkbookPart!.WorksheetParts
            .SelectMany(wsp => wsp.GetPartsOfType<PivotTablePart>())
            .Single()
            .PivotTableDefinition;

        return definition.Elements<PivotFilters>()
            .SelectMany(f => f.Elements<PivotFilter>())
            .ToList();
    }

    /// <summary>
    /// A minimal workbook whose single pivot table carries the given filters.
    /// </summary>
    internal static void Create(Stream stream, params PivotFilter[] filters)
    {
        using var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        var workbookPart = doc.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());

        var dataSheetPart = workbookPart.AddNewPart<WorksheetPart>();
        sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(dataSheetPart), SheetId = 1, Name = "Data" });

        var sheetData = new SheetData();
        sheetData.Append(CreateRow(1, "Name", "Revenue"));
        sheetData.Append(CreateNumericRow(2, "Cookies", 500));
        sheetData.Append(CreateNumericRow(3, "Cake", 800));
        dataSheetPart.Worksheet = new Worksheet(sheetData);

        var pivotSheetPart = workbookPart.AddNewPart<WorksheetPart>();
        sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(pivotSheetPart), SheetId = 2, Name = "PivotSheet" });
        pivotSheetPart.Worksheet = new Worksheet(new SheetData());

        var cachePart = workbookPart.AddNewPart<PivotTableCacheDefinitionPart>();
        var cachePartId = workbookPart.GetIdOfPart(cachePart);

        var cacheDefinition = new PivotCacheDefinition
        {
            Id = "rId1",
            RefreshOnLoad = true,
            CreatedVersion = 5,
            RefreshedVersion = 5,
        };
        cacheDefinition.AddNamespaceDeclaration("r", OpenXmlConst.RelationshipsNs);

        var cacheSource = new CacheSource { Type = SourceValues.Worksheet };
        cacheSource.Append(new WorksheetSource { Sheet = "Data", Reference = "A1:B3" });
        cacheDefinition.Append(cacheSource);

        var cacheFields = new CacheFields();

        var nameField = new CacheField { Name = "Name" };
        var nameShared = new SharedItems { ContainsSemiMixedTypes = false, ContainsString = true, ContainsNumber = false, Count = 2 };
        nameShared.Append(new StringItem { Val = "Cookies" });
        nameShared.Append(new StringItem { Val = "Cake" });
        nameField.SharedItems = nameShared;
        cacheFields.Append(nameField);

        var revenueField = new CacheField { Name = "Revenue" };
        revenueField.SharedItems = new SharedItems
        {
            ContainsSemiMixedTypes = false,
            ContainsString = false,
            ContainsNumber = true,
            MinValue = 500,
            MaxValue = 800,
        };
        cacheFields.Append(revenueField);

        cacheFields.Count = 2;
        cacheDefinition.Append(cacheFields);

        var recordsPart = cachePart.AddNewPart<PivotTableCacheRecordsPart>();
        var records = new PivotCacheRecords { Count = 2 };
        records.Append(new PivotCacheRecord(new FieldItem { Val = 0 }, new NumberItem { Val = 500 }));
        records.Append(new PivotCacheRecord(new FieldItem { Val = 1 }, new NumberItem { Val = 800 }));
        recordsPart.PivotCacheRecords = records;

        cachePart.PivotCacheDefinition = cacheDefinition;

        var pivotCaches = new PivotCaches();
        pivotCaches.Append(new PivotCache { CacheId = 0, Id = cachePartId });
        workbookPart.Workbook.Append(pivotCaches);

        var pivotTablePart = pivotSheetPart.AddNewPart<PivotTablePart>();
        pivotTablePart.CreateRelationshipToPart(cachePart);

        var pivotTableDef = new PivotTableDefinition
        {
            Name = "PivotTable1",
            CacheId = 0,
            DataCaption = "Values",
            CreatedVersion = 5,
            UpdatedVersion = 5,
        };

        pivotTableDef.Append(new Location { Reference = "A1:B4", FirstHeaderRow = 1, FirstDataRow = 2, FirstDataColumn = 1 });

        var pivotFields = new PivotFields { Count = 2 };
        var pf0 = new PivotField { Axis = PivotTableAxisValues.AxisRow, ShowAll = false };
        var items0 = new Items { Count = 3 };
        items0.Append(new Item { Index = 0 });
        items0.Append(new Item { Index = 1 });
        items0.Append(new Item { ItemType = ItemValues.Default });
        pf0.Append(items0);
        pivotFields.Append(pf0);
        pivotFields.Append(new PivotField { DataField = true, ShowAll = false });
        pivotTableDef.Append(pivotFields);

        var rowFields = new RowFields { Count = 1 };
        rowFields.Append(new Field { Index = 0 });
        pivotTableDef.Append(rowFields);

        var rowItems = new RowItems { Count = 3 };
        rowItems.Append(new RowItem(new MemberPropertyIndex { Val = 0 }));
        rowItems.Append(new RowItem(new MemberPropertyIndex { Val = 1 }));
        rowItems.Append(new RowItem(new MemberPropertyIndex()) { ItemType = ItemValues.Grand });
        pivotTableDef.Append(rowItems);

        var colItems = new ColumnItems { Count = 1 };
        colItems.Append(new RowItem(new MemberPropertyIndex()) { ItemType = ItemValues.Grand });
        pivotTableDef.Append(colItems);

        var dataFields = new DataFields { Count = 1 };
        dataFields.Append(new DataField { Name = "Sum of Revenue", Field = 1 });
        pivotTableDef.Append(dataFields);

        pivotTableDef.Append(new PivotTableStyle { Name = "PivotStyleLight16", ShowRowHeaders = true, ShowColumnHeaders = true });

        if (filters.Length > 0)
        {
            // Schema order puts filters after pivotTableStyleInfo.
            var pivotFilters = new PivotFilters { Count = (uint)filters.Length };
            foreach (var filter in filters)
                pivotFilters.Append(filter);

            pivotTableDef.Append(pivotFilters);
        }

        pivotTablePart.PivotTableDefinition = pivotTableDef;
    }

    private static Row CreateRow(uint rowIndex, params string[] values)
    {
        var row = new Row { RowIndex = rowIndex };
        for (var i = 0; i < values.Length; i++)
        {
            row.Append(new Cell
            {
                CellReference = $"{(char)('A' + i)}{rowIndex}",
                DataType = CellValues.String,
                CellValue = new CellValue(values[i]),
            });
        }

        return row;
    }

    private static Row CreateNumericRow(uint rowIndex, string name, double value)
    {
        var row = new Row { RowIndex = rowIndex };
        row.Append(new Cell { CellReference = $"A{rowIndex}", DataType = CellValues.String, CellValue = new CellValue(name) });
        row.Append(new Cell { CellReference = $"B{rowIndex}", DataType = CellValues.Number, CellValue = new CellValue(value) });
        return row;
    }
}
