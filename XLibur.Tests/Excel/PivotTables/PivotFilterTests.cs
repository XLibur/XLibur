using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;
using XLibur.Excel.IO;

namespace XLibur.Tests.Excel.PivotTables;

/// <summary>
/// The <c>filters</c> collection holds the label, value, date and top-N filters applied to pivot
/// fields — what Excel offers from a field's dropdown. Dropping it on load silently un-filters
/// the pivot table on the next save, which changes what the workbook shows.
/// </summary>
/// <remarks>
/// Not the report-filter axis: that is <c>pageFields</c>, modelled by
/// <c>XLPivotTable.Filters</c>, and was already supported.
/// </remarks>
public class PivotFilterTests
{
    [Test]
    public async Task PivotFilters_SurviveARoundTrip()
    {
        using var input = new MemoryStream();
        CreateWorkbookWithPivotFilters(input);
        input.Position = 0;

        using var wb = new XLWorkbook(input);
        using var output = new MemoryStream();
        wb.SaveAs(output);

        var written = ReadPivotFilters(output);

        await Assert.That(written.Count).IsEqualTo(2);

        var caption = written[0];
        await Assert.That(caption.Field!.Value).IsEqualTo(0U);
        await Assert.That(caption.Id!.Value).IsEqualTo(1U);
        await Assert.That(caption.Type!.InnerText).IsEqualTo("captionEqual");
        await Assert.That(caption.StringValue1!.Value).IsEqualTo("Cookies");
        await Assert.That(caption.Name!.Value).IsEqualTo("Only cookies");

        var value = written[1];
        await Assert.That(value.Field!.Value).IsEqualTo(1U);
        await Assert.That(value.Id!.Value).IsEqualTo(2U);
        await Assert.That(value.Type!.InnerText).IsEqualTo("valueGreaterThan");
        await Assert.That(value.EvaluationOrder!.Value).IsEqualTo(-1);
    }

    /// <summary>
    /// <c>autoFilter</c> is required by the schema and carries the actual criteria, so a round
    /// trip that kept only the attributes would still lose what the filter does.
    /// </summary>
    [Test]
    public async Task PivotFilterAutoFilter_SurvivesARoundTrip()
    {
        using var input = new MemoryStream();
        CreateWorkbookWithPivotFilters(input);
        input.Position = 0;

        using var wb = new XLWorkbook(input);
        using var output = new MemoryStream();
        wb.SaveAs(output);

        var autoFilter = ReadPivotFilters(output)[0].AutoFilter;
        await Assert.That(autoFilter).IsNotNull();

        var filterColumn = autoFilter!.Elements<FilterColumn>().Single();
        await Assert.That(filterColumn.ColumnId!.Value).IsEqualTo(0U);

        var filterValue = filterColumn.Elements<Filters>().Single().Elements<Filter>().Single();
        await Assert.That(filterValue.Val!.Value).IsEqualTo("Cookies");
    }

    /// <summary>
    /// The top-N filter carries a <c>top10</c> child rather than a value list, so it exercises a
    /// different branch of the preserved sub-tree.
    /// </summary>
    [Test]
    public async Task PivotFilterTop10AutoFilter_SurvivesARoundTrip()
    {
        using var input = new MemoryStream();
        CreateWorkbookWithPivotFilters(input);
        input.Position = 0;

        using var wb = new XLWorkbook(input);
        using var output = new MemoryStream();
        wb.SaveAs(output);

        var autoFilter = ReadPivotFilters(output)[1].AutoFilter!;
        var top10 = autoFilter.Elements<FilterColumn>().Single().Elements<Top10>().Single();

        await Assert.That(top10.Top!.Value).IsTrue();
        await Assert.That(top10.Val!.Value).IsEqualTo(3D);
    }

    [Test]
    public async Task LoadedPivotFilters_AreExposedOnThePivotTable()
    {
        using var input = new MemoryStream();
        CreateWorkbookWithPivotFilters(input);
        input.Position = 0;

        using var wb = new XLWorkbook(input);
        var pt = (XLPivotTable)wb.Worksheet("PivotSheet").PivotTables.Single();

        await Assert.That(pt.PivotFilters.Count).IsEqualTo(2);
        await Assert.That(pt.PivotFilters[0].Type).IsEqualTo("captionEqual");
        await Assert.That(pt.PivotFilters[0].StringValue1).IsEqualTo("Cookies");
        await Assert.That(pt.PivotFilters[1].EvaluationOrder).IsEqualTo(-1);
        await Assert.That(pt.PivotFilters[1].AutoFilterXml).Contains("top10");
    }

    /// <summary>
    /// Most pivot tables have no filters, and an empty <c>filters</c> element is not valid, so
    /// nothing must be written for them.
    /// </summary>
    [Test]
    public async Task PivotTableWithoutFilters_WritesNoFiltersElement()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet("Data");
        var range = data.FirstCell().InsertData(new object[]
        {
            ("Pastry", "Sold"),
            ("Waffle", 3),
            ("Donut", 5),
        });

        var pivots = wb.AddWorksheet("Pivots");
        var pt = pivots.PivotTables.Add("pvt", pivots.Cell("A1"), range);
        pt.RowLabels.Add("Pastry");
        pt.Values.Add("Sold");

        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var definition = doc.WorkbookPart!.WorksheetParts
            .SelectMany(wsp => wsp.GetPartsOfType<PivotTablePart>())
            .Single()
            .PivotTableDefinition;

        await Assert.That(definition.Elements<PivotFilters>().Any()).IsFalse();
    }

    private static List<PivotFilter> ReadPivotFilters(MemoryStream saved)
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
    /// A minimal workbook whose pivot table carries two filters: a caption filter holding a value
    /// list, and a value filter holding a top-N.
    /// </summary>
    private static void CreateWorkbookWithPivotFilters(Stream stream)
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

        // The element under test. Schema order puts filters after pivotTableStyleInfo.
        var filters = new PivotFilters { Count = 2 };
        filters.Append(BuildCaptionFilter());
        filters.Append(BuildValueFilter());
        pivotTableDef.Append(filters);

        pivotTablePart.PivotTableDefinition = pivotTableDef;
    }

    /// <summary>
    /// A caption filter whose <c>autoFilter</c> holds an explicit value list.
    /// </summary>
    private static PivotFilter BuildCaptionFilter()
    {
        var filterColumn = new FilterColumn { ColumnId = 0U };
        var filterValues = new Filters();
        filterValues.Append(new Filter { Val = "Cookies" });
        filterColumn.Append(filterValues);

        var autoFilter = new AutoFilter { Reference = "A1:A3" };
        autoFilter.Append(filterColumn);

        var pivotFilter = new PivotFilter
        {
            Field = 0U,
            Id = 1U,
            Type = PivotFilterValues.CaptionEqual,
            StringValue1 = "Cookies",
            Name = "Only cookies",
        };
        pivotFilter.Append(autoFilter);
        return pivotFilter;
    }

    /// <summary>
    /// A value filter whose <c>autoFilter</c> holds a top-N instead, and a non-default
    /// <c>evalOrder</c>.
    /// </summary>
    private static PivotFilter BuildValueFilter()
    {
        var filterColumn = new FilterColumn { ColumnId = 0U };
        filterColumn.Append(new Top10 { Top = true, Val = 3D });

        var autoFilter = new AutoFilter { Reference = "B1:B3" };
        autoFilter.Append(filterColumn);

        var pivotFilter = new PivotFilter
        {
            Field = 1U,
            Id = 2U,
            Type = PivotFilterValues.ValueGreaterThan,
            EvaluationOrder = -1,
        };
        pivotFilter.Append(autoFilter);
        return pivotFilter;
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
