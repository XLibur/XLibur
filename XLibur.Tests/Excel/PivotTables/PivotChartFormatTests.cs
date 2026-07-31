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
/// A PivotChart keeps its manual per-series and per-point formatting in the chart part, but the
/// <c>chartFormats</c> collection on the pivot table is what ties each formatting record to the
/// pivot area it applies to. Losing the collection on load leaves the chart pointing at nothing,
/// so the formatting silently disappears on the next save.
/// </summary>
public class PivotChartFormatTests
{
    [Test]
    public async Task ChartFormats_SurviveARoundTrip()
    {
        using var input = new MemoryStream();
        CreateWorkbookWithChartFormats(input);
        input.Position = 0;

        using var wb = new XLWorkbook(input);
        using var output = new MemoryStream();
        wb.SaveAs(output);

        var written = ReadChartFormats(output);

        await Assert.That(written.Count).IsEqualTo(2);

        await Assert.That(written[0].Chart!.Value).IsEqualTo(0U);
        await Assert.That(written[0].Format!.Value).IsEqualTo(3U);
        await Assert.That(written[0].Series?.Value ?? false).IsTrue();
        await Assert.That(written[0].PivotArea).IsNotNull();

        await Assert.That(written[1].Chart!.Value).IsEqualTo(1U);
        await Assert.That(written[1].Format!.Value).IsEqualTo(7U);
        await Assert.That(written[1].Series?.Value ?? false).IsFalse();
    }

    /// <summary>
    /// The pivot area is the half of the record that says what is being formatted, so a
    /// round trip that kept only the ids would still lose the link.
    /// </summary>
    [Test]
    public async Task ChartFormatPivotArea_SurvivesARoundTrip()
    {
        using var input = new MemoryStream();
        CreateWorkbookWithChartFormats(input);
        input.Position = 0;

        using var wb = new XLWorkbook(input);
        using var output = new MemoryStream();
        wb.SaveAs(output);

        var pivotArea = ReadChartFormats(output)[0].PivotArea!;
        var reference = pivotArea.PivotAreaReferences!.Elements<PivotAreaReference>().Single();

        await Assert.That(reference.Field!.Value).IsEqualTo(0U);
        await Assert.That(reference.Elements<FieldItem>().Single().Val!.Value).IsEqualTo(1U);
    }

    [Test]
    public async Task LoadedChartFormats_AreExposedOnThePivotTable()
    {
        using var input = new MemoryStream();
        CreateWorkbookWithChartFormats(input);
        input.Position = 0;

        using var wb = new XLWorkbook(input);
        var pt = (XLPivotTable)wb.Worksheet("PivotSheet").PivotTables.Single();

        await Assert.That(pt.ChartFormats.Count).IsEqualTo(2);
        await Assert.That(pt.ChartFormats[0].Chart).IsEqualTo(0U);
        await Assert.That(pt.ChartFormats[0].Format).IsEqualTo(3U);
        await Assert.That(pt.ChartFormats[0].Series).IsTrue();
    }

    /// <summary>
    /// The overwhelming majority of pivot tables have no chart, and an empty
    /// <c>chartFormats</c> element is not valid, so nothing must be written for them.
    /// </summary>
    [Test]
    public async Task PivotTableWithoutChartFormats_WritesNoChartFormatsElement()
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

        await Assert.That(definition.Elements<ChartFormats>().Any()).IsFalse();
    }

    private static System.Collections.Generic.List<ChartFormat> ReadChartFormats(MemoryStream saved)
    {
        saved.Position = 0;
        using var doc = SpreadsheetDocument.Open(saved, false);
        var definition = doc.WorkbookPart!.WorksheetParts
            .SelectMany(wsp => wsp.GetPartsOfType<PivotTablePart>())
            .Single()
            .PivotTableDefinition;

        return definition.Elements<ChartFormats>()
            .SelectMany(cf => cf.Elements<ChartFormat>())
            .ToList();
    }

    /// <summary>
    /// A minimal workbook whose pivot table carries a <c>chartFormats</c> collection: one record
    /// formatting a whole series and one formatting a single point.
    /// </summary>
    private static void CreateWorkbookWithChartFormats(Stream stream)
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

        // The element under test. Schema order puts chartFormats after dataFields and before
        // pivotTableStyleInfo.
        var chartFormats = new ChartFormats { Count = 2 };
        chartFormats.Append(BuildChartFormat(chart: 0, format: 3, series: true, fieldItem: 1));
        chartFormats.Append(BuildChartFormat(chart: 1, format: 7, series: false, fieldItem: 0));
        pivotTableDef.Append(chartFormats);

        pivotTableDef.Append(new PivotTableStyle { Name = "PivotStyleLight16", ShowRowHeaders = true, ShowColumnHeaders = true });

        pivotTablePart.PivotTableDefinition = pivotTableDef;
    }

    private static ChartFormat BuildChartFormat(uint chart, uint format, bool series, uint fieldItem)
    {
        var pivotArea = new PivotArea { DataOnly = false, Outline = false, FieldPosition = 0U };
        var references = new PivotAreaReferences { Count = 1 };
        var reference = new PivotAreaReference { Field = 0U, Count = 1U, Selected = false };
        reference.Append(new FieldItem { Val = fieldItem });
        references.Append(reference);
        pivotArea.Append(references);

        var chartFormat = new ChartFormat { Chart = chart, Format = format, Series = series };
        chartFormat.Append(pivotArea);
        return chartFormat;
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
